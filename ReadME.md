![example](https://github.com/creetz/ExchangeLogfileSearcher/blob/master/example.png)

    .SYNOPSIS
    Powershell Log File Search - ExchangeLogfileSearcher.ps1
   
    Christian Reetz
    (Updated by Christian Reetz)
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    22.02.2018
	
    .DESCRIPTION

    This script collects entries from multiple Exchange Log Files and create a single csv-file.
   	
    PARAMETER search
   
    PARAMETER start
    Start Date in english format

    PARAMETER end
    End Date in english format

    PARAMETER sourcelogfilepath
    Each row start with the sourceserver from the logfile

    PARAMETER germantimeformat
    $true or keep clear
         
     EXAMPLES
    .\ExchangeLogfileSearcher.ps1 -search "192.168.100.100" -start 01/31/2017 -end 12/31/2017
    .\ExchangeLogfileSearcher.ps1 -search "192.168.100.100" -start "31.01.2017 12:00" -end "31.12.2017 12:00" -germantimeformat $true
    .\ExchangeLogfileSearcher.ps1 -search "192.168.100.100" -start 01/31/2017 -end 12/31/2017 -sourcelogfilepath $true
