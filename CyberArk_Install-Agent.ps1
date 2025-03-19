# The following is based on documentation located at: https://docs.cyberark.com/epm/latest/en/content/installation/windows-installagents.htm#tabset-1-tab-3.
# Setup authentication.
$reinstallToken = '############################'
$AuthBody = @{
    Username      = '#########################'
    Password      = '#########################'
    ApplicationID = '######'
}
$AuthJsonBody = $AuthBody | ConvertTo-Json
$AuthParams = @{
    Body        = $AuthJsonBody
    ContentType = 'application/json'
    Method      = 'Post'
    Uri         = 'https://login.epm.cyberark.com/EPM/API/Auth/EPM/Logon'
}
$Token = Invoke-RestMethod @AuthParams

# Get set lists.
$Headers = @{
    Authorization = "basic $($token.EPMAuthenticationResult)"
}
$SetParams = @{
    Method = 'Get'
    Uri    = "$($token.ManagerURL)/EPM/API/Sets"
}
$Sets = Invoke-RestMethod @SetParams -Headers $Headers
$Sets = $Sets.Sets | Where-Object { $_.Name -like '######*' }

# Build API calls to gather download files.
$agentHeaders = @{
    Authorization = "basic $($token.EPMAuthenticationResult)"
}
$agentParams = @{
    Method = 'GET'
    Uri    = "$($token.ManagerURL)/EPM/API/Sets/$($Sets.Id)/Computers/Packages?os=windows" # https://<EPM_Server>/EPM/API/<Version>/Sets/{SetId}/Computers/Packages?os={windows}....macos
}
$agentWindows = Invoke-RestMethod @agentParams -Headers $agentHeaders

# Filter out on the latest agent by RealeaseDate.
$agent = $agentWindows.Packages | Where-Object { $_.ReleaseDate -eq ($agentWindows.Packages.ReleaseDate | Sort-Object ReleaseDate -Descending | Select-Object -First 1) }

# Filter current workstation for 32 or 64 bit, or ARM chipset.
$os = Get-WmiObject -Class Win32_OperatingSystem
$osArchitecture = $os.OSArchitecture

switch ( $osArchitecture ) {
    'ARM 64-bit Processor' { $Architecture = 'arm64' }
    '64-Bit' { $Architecture = 'x64' }
}

# Filter out the correct agent for the workstation.
$agent = $agent | Where-Object { $_.PackageArch -eq "$Architecture" }

# Aqcuires the download URL for the agent.
$agentParams = @{
    Method = 'GET'
    Uri    = "$($token.ManagerURL)/EPM/API/Sets/$($Sets.Id)/Computers/Packages/$($agent.Id)/URL"
}
$agentWindows = Invoke-RestMethod @agentParams -Headers $agentHeaders

# Download the agent.
Invoke-WebRequest -Uri $agentWindows -Method 'Get' -OutFile "$env:temp\epmagent.msi" -UseBasicParsing

# Download the agent file and installation key.
$confParams = @{
    Method = 'GET'
    Uri    = "$($token.ManagerURL)/EPM/API/Sets/$($Sets.Id)/Computers/Packages/$($agent.Id)/Configuration"
} 
$installData = Invoke-WebRequest @confParams -Headers $agentHeaders -UseBasicParsing

# Extract the configuration file from the response. The configuration file section starts with "{"iot":" and ends with "}".
$configFile = $installData.RawContent
$configFile.Substring($configFile.IndexOf('{'), $configFile.LastIndexOf('}') - $configFile.IndexOf('{') + 1) | Out-File "$env:temp\epm.config" -Encoding UTF8

# Extract the installation key from the $installData header response. This section starts with "installationkey":"
$installationKey = $installdata.headers.installationkey

# Install the agent.
MsiExec.exe /i "$env:temp\epmagent.msi" INSTALLATIONKEY=$installationKey CONFIGURATION="$env:temp\epm.config" SECURE_TOKEN=$reinstalltoken ISDEPLOYMENT=Yes /l*v "$env:temp\epm_install.log" REINSTALLMODE=vm /qn