#!/bin/bash

logandmetadir="/Library/Logs/Microsoft/IntuneScripts/"
log="$logandmetadir/cyberark_agent-install.log"
scriptname="CyberArk Agent Install"
securetoken="8675309"
company="COMPANY"

if [ -d $logandmetadir ]; then
    echo "# $(date) | Log directory already exists - $logandmetadir"
else
    echo "# $(date) | creating log directory - $logandmetadir"
    mkdir -p $logandmetadir
fi

exec 1>> $log 2>&1

echo ""
echo "##############################################################"
echo "# $(date) | Starting $scriptname"
echo "############################################################"
echo ""

if [ $(sudo profiles show | grep -c "CyberArk") -lt 1 ]; then
    echo "System extensions are not installed or incomplete."
    exit 0
else
    echo "System extensions are installed."
fi

echo "Setting up authentication..."
authBody=$(cat <<EOF
{
    "Username": "user@domain.com",
    "Password": "P@ssword123!",
    "ApplicationID": "COMPANY"
}
EOF
)

authResponse=$(curl -s -X POST -H "Content-Type: application/json" -d "$authBody" "https://login.epm.cyberark.com/EPM/API/Auth/EPM/Logon")
token=$(echo "$authResponse" | jq -r '.EPMAuthenticationResult')
managerURL=$(echo "$authResponse" | jq -r '.ManagerURL')

echo "Getting set lists..."
setsResponse=$(curl -s -X GET -H "Authorization: basic $token" "$managerURL/EPM/API/Sets")
sets=$(echo "$setsResponse" | jq -c ".Sets[] | select(.Name | startswith("$company"))")
setId=$(echo "$sets" | jq -r '.Id')

echo "Building API calls to gather download files..."
agentResponse=$(curl -s -X GET -H "Authorization: basic $token" "$managerURL/EPM/API/Sets/$setId/Computers/Packages?os=macos")
agentId=$(echo "$agentResponse" | grep -o '"Id":"[^"]*"' | head -n 1 | cut -d':' -f2 | tr -d '"')

echo "Acquiring the download URL for the agent..."
agentURL=$(curl -s -X GET -H "Authorization: basic $token" "$managerURL/EPM/API/Sets/$setId/Computers/Packages/$agentId/URL")

echo "Sending the agent URL to a string variable for CURL to use..."
agentURLCurl=$(echo "$agentURL" | sed 's/\\//g' | sed 's/\"//g')

echo "Downloading the agent..."
curl -s -o "/tmp/epmagent.zip" "$agentURLCurl"

echo "Downloading the agent file and installation key..."
installData=$(curl -D /tmp/epmkey.txt -s -X GET -H "Authorization: basic $token" "$managerURL/EPM/API/Sets/$setId/Computers/Packages/$agentId/Configuration")

echo "Extracting the configuration file from the response..."
configFile=$(echo "$installData" | jq -r '. | tostring')
echo "$configFile" > "/tmp/epm.config"

echo "Extracting the installation key from the header file..."
installKey=$(grep -i '^installationkey:' /tmp/epmkey.txt | awk -F': ' '{print $2}' | tr -d '\r')
if [ -z "$installKey" ]; then
    echo "Error: Installation key not found in /tmp/epmkey.txt."
    exit 1
fi


echo "Installing the agent..."
if [ -d "/tmp/epmagent" ]; then
    echo "Agent directory already exists. Removing it to force override..."
    rm -rf "/tmp/epmagent"
fi
unzip -o "/tmp/epmagent.zip" -d "/tmp/epmagent"
unzip -o /tmp/epmagent/Install\ CyberArk\ EPM.app.zip -d "/tmp/epmagent"

sudo /private/tmp/epmagent/Install\ CyberArk\ EPM.app/Contents/MacOS/CyberArk\ EPM -installationKey $installKey -configuration '/private/tmp/epm.config' -token $securetoken -cleanup

echo "Cleaning up..."
rm -rf "/tmp/epmagent" "/tmp/epmagent.zip" "/tmp/epmkey.txt"
exit 0