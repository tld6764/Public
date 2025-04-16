#!/bin/sh

logandmetadir="/Library/Logs/Microsoft/IntuneScripts/"
log="$logandmetadir/cyberarkagentuninstall.log"
scriptname="CyberArk Agent Uninstall"
securetoken="8675309"

if [ -d $logandmetadir ]; then
    echo "# $(date) | Log directory already exists - $logandmetadir"
else
    echo "# $(date) | creating log directory - $logandmetadir"
    mkdir -p $logandmetadir
fi

sudo exec 1>> $log 2>&1

echo ""
echo "##############################################################"
echo "# $(date) | Starting $scriptname"
echo "############################################################"
echo ""

if [ "$securetoken" = "" ]; then
    echo "`date +"%Y-%m-%d_%H-%M-%S"` INFO: No Protection Token provided, Trying without it..." 
    sudo /Applications/CyberArk\ EPM.app/Contents/Helpers/CyberArkEPMUninstall -full
    if [ $? -eq 0 ] ; then
        echo "`date +"%Y-%m-%d_%H-%M-%S"` INFO: Successful uninstall completed, Exiting..."
        exit 0
    else    
        echo "`date +"%Y-%m-%d_%H-%M-%S"` ERROR: Failed to uninstall without agent protection token"
        exit 1
    fi
else
    echo "`date +"%Y-%m-%d_%H-%M-%S"` INFO: Protection Token provided, Trying with it..." 
    sudo /Applications/CyberArk\ EPM.app/Contents/Helpers/CyberArkEPMUninstall -token $securetoken -full
    if [ $? -eq 0 ] ; then
        echo "`date +"%Y-%m-%d_%H-%M-%S"` INFO: Successful uninstall completed with token, Exiting..."
        exit 0
    else    
        echo "`date +"%Y-%m-%d_%H-%M-%S"` ERROR: Failed to uninstall with token"
        exit 1
    fi
fi
exit 0