<#
    .SYNOPSIS
    Powershell Log File Search - ExchangeLogfileSearcher.ps1
   
   	Christian Reetz
    (Updated by Christian Reetz)
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	26.10.2018
	
    .DESCRIPTION

    This scripts is Part of the "Exchange Health Center" - Code
    
    This script generate a html view which show the exchange server status
   	
   	PARAMETER search
   
    PARAMETER lastdays
    0 is not possible, but you use for example 0.5 for an half day
    
    PARAMETER start
    Start Date in english format

    PARAMETER end
    End Date in english format

    PARAMETER sourcelogfilepath
    Each row start with the sourceserver from the logfile

    PARAMETER germantimeformat
    $true or keep clear

    PARAMETER targetexchangeserver
    specifie targetexchangeserver

    PARAMETER IISScopeLastLogFile
    $true or keep clear
    When $true file-scope is "select the last iis-logfile - then you can ignore -start and -end!"
         
	EXAMPLES
    .\ExchangeLogfileSearcher.ps1 -search "192.168.100.100" -start 01/31/2017 -end 12/31/2017
    .\ExchangeLogfileSearcher.ps1 -search "192.168.100.100" -start "31.01.2017 12:00" -end "31.12.2017 12:00" -germantimeformat $true
    .\ExchangeLogfileSearcher.ps1 -search "192.168.100.100" -start 01/31/2017 -end 12/31/2017 -sourcelogfilepath $true

    #>

[CmdletBinding()]
Param(
	[Parameter(Mandatory=$true)][string]$search,
    [Parameter(Mandatory=$false)] $sourcelogfilepath,
    [Parameter(Mandatory=$false)] $germantimeformat,
    [Parameter(Mandatory=$false)] $IISScopeLastLogFile,
    [Parameter(Mandatory=$false)] $targetexchangeserver,
    [Parameter(Mandatory=$false)][string]$start,
    [Parameter(Mandatory=$false)][string]$end
)

if (!($start) -and !($IISScopeLastLogFile))
{
    Write-Host "Please specifie a start date! (f.E. 10/26/2018)"
    $start = read-host "StartDate"
    #$start = get-date $start
}

if (!($end) -and !($IISScopeLastLogFile))
{
    Write-Host "Please specifie an end date! (f.E. 10/26/2018)"
    $end = read-host "EndDate"
    #$end = get-date $end
}

if (($germantimeformat) -and $start)
{
    $startday = $start.Split('.')[0]
    $startmonth = $start.Split('.')[1]
    $startyear = $start.Split('.')[2].split(' ')[0]
    $starthour = $start.Split('.')[2].split(' ')[1].split(':')[0]
    $startminute = $start.Split('.')[2].split(' ')[1].split(':')[1]
    $start = get-date -Year $startyear -Month $startmonth -Day $startday -Hour $starthour -Minute $startminute
}

if (($germantimeformat) -and $end)
{
    $endday = $end.Split('.')[0]
    $endmonth = $end.Split('.')[1]
    $endyear = $end.Split('.')[2].split(' ')[0]
    $endhour = $end.Split('.')[2].split(' ')[1].split(':')[0]
    $endminute = $end.Split('.')[2].split(' ')[1].split(':')[1]
    $end = get-date -Year $endyear -Month $endmonth -Day $endday -Hour $endhour -Minute $endminute
}


if ($targetexchangeserver)
{
    $exchsrvs = [pscustomobject]@{name="$targetexchangeserver"}
}
else
{
    $exchsrvs = get-exchangeserver | Sort-Object identity
}

cls
 
echo "---------------------------------------------------------"  
echo ""
echo "               Advanced Logs-File Searcher"
echo ""
echo "    1. Search HTTPProxy (EWS, MAPI, etc.) - Logs"
echo "    2. Search ExchangeTransportService - Logs"
echo "    3. Search MAPI CLIENT ACCESS - Logs"
echo "    4. Search EWS - Logs"
echo "    5. Search Calendar Repair Assistant"
echo "    6. IISLog"
#echo "    7. Remote Powershell [BETA]"
echo ""
echo "---------------------------------------------------------"  
echo ""  

echo ""  
$answer = read-host "Please Make a Selection"  
if ($answer -eq 1){$choice = 1}
if ($answer -eq 2){$choice = 2}
if ($answer -eq 3){$choice = 3}
if ($answer -eq 4){$choice = 4}
if ($answer -eq 5){$choice = 5}
if ($answer -eq 6){$choice = 6}
#if ($answer -eq 7){$choice = 7}
if ($answer -eq 0){break}


if ($choice -eq 1)
{
    if (($webservice -ne "mapi") -or ($webservice -ne "autodiscover") -or ($webservice -ne "eas") -or ($webservice -ne "ecp") -or ($webservice -ne "ews") -or ($webservice -ne "oab") -or ($webservice -ne "owa") -or ($webservice -ne "owacalendar") -or ($webservice -ne "powershell") -or ($webservice -ne "rpchttp"))
    {
        $webservice = read-host "Please specifie a vaild webservice (like mapi, autodiscover, eas, ecp, ews, oab, owa, owacalendar, powershell or rpchttp)"
    }

    $logtype = $webservice
    
    $logpath = "c$\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\$webservice\"
}

if ($choice -eq 2)
{
    if (($transportservice -ne "SmtpReceive") -or ($transportservice -ne "Smtpsend"))
    {
        $hubdORfrontend = read-host "Please specifie a vaild transportrole (hub or frontend)"
        $transportservice = read-host "Please specifie a vaild transportservice (like SmtpReceive or Smtpsend)"
    }

    $logtype = $transportservice
    
    $logpath = "c$\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\$hubdORfrontend\ProtocolLog\$transportservice\"
}

if ($choice -eq 3)
{
    $logpath = "c$\Program Files\Microsoft\Exchange Server\V15\Logging\MAPI Client Access\"

    $logtype = "MapiCAS"
}

if ($choice -eq 4)
{
    $logpath = "c$\Program Files\Microsoft\Exchange Server\V15\Logging\EWS\"   

    $logtype = "EWS"
}

if ($choice -eq 5)
{
    $logpath = "c$\Program Files\Microsoft\Exchange Server\V15\Logging\Calendar Repair Assistant\"   

    $logtype = "CRA"
}

if ($choice -eq 6)
{
    $logpath = "c$\inetpub\logs\LogFiles\W3SVC1\"
    $logpath2 = "c$\inetpub\logs\LogFiles\W3SVC2\"

    $logtype = "IISLogs"
}

<#
if ($choice -eq 7)
{
    $logpath = "c$\Program Files\Microsoft\Exchange Server\V15\Logging\CmdletInfra\Powershell-Proxy\Cmdlet\"   
}
#>

echo ""
write-host -ForegroundColor Magenta "Important! add 2h - LogFiles are GMT+0"
echo ""

$results = @()
$sourcelogfilepathserver = @()
Foreach ($exchsrv in $exchsrvs)
{
    $path = "\\$($exchsrv.name)\$logpath"
    write-host -ForegroundColor Yellow "Search in $path"

    if ($IISScopeLastLogFile -and ($choice -eq 6))
    {
        $files = (Get-ChildItem "$path" | sort-object lastwritetime -Descending)[0] 
    }
    else
    {
        $files = Get-ChildItem "$path" | ? {($_.LastWriteTime -gt $start) -and ($_.LastWriteTime -lt $end)}
    }

    if ($files.name.count -gt 0)
    {
        foreach ($file in $files)
        {
            $hit = $results.count
            write-host -ForegroundColor Cyan "Search in $file" -NoNewline
            if ($sourcelogfilepath)
            {
                $results += Get-Content $file.VersionInfo.filename | select-string "$search" 
            }
            else
            {
                $results += Get-Content $file.VersionInfo.filename | select-string "$search"
            }
            if ($results.count -gt $hit)
            {
                Write-Host -ForegroundColor White " - hit!"
                $sourcelogfilepathserver += "$($path)$($file)"
            }
            else
            {
                Write-Host ""
            }
        }
    }

    if ($choice -eq 6)
    {
        $results_iis_backend = @()

        $path2 = "\\$($exchsrv.name)\$logpath2"
        write-host -ForegroundColor Yellow "Search in $path2"
        
        if ($IISScopeLastLogFile)
        {
            $files2 = (Get-ChildItem "$path2" | sort-object lastwritetime -Descending)[0] 
        }
        else
        {
            $files2 = Get-ChildItem "$path2" | ? {($_.LastWriteTime -gt $start) -and ($_.LastWriteTime -lt $end)}
        }
        
        if ($files2.name.count -gt 0)
        {
            foreach ($file in $files2)
            {
                $hit = $results_iis_backend.count
                write-host -ForegroundColor Cyan "Search in $file" -NoNewline
                if ($sourcelogfilepath)
                {
                    $results_iis_backend += Get-Content $file.VersionInfo.filename | select-string "$search" 
                }
                else
                {
                    $results_iis_backend += Get-Content $file.VersionInfo.filename | select-string "$search"
                }
                if ($results_iis_backend.count -gt $hit)
                {
                    Write-Host -ForegroundColor White " - hit!"
                    $sourcelogfilepathserver += "$($path)$($file)"
                }
                else
                {
                    Write-Host ""
                }
            }
        }
            
    }
     
        
    if (!($csvheadline) -and ($files))
    {
        $csvheadline = Get-Content $files[-1].VersionInfo.filename -first 10 | select-string "#Fields:"
    }
}

Write-Host -ForegroundColor Green "$($results.count) Entries found!"
if ($choice -eq 6)
{
    Write-Host -ForegroundColor Cyan "$($results_iis_backend.count) IIS_BackEnd Entries found!"
}

if ($sourcelogfilepath)
{
    $new_results =@()
    $csvheadline = "sourcelogfilepath,$($csvheadline.line)"
    $i = 0

    foreach ($entry in $results)
    {
        $new_results += "$($sourcelogfilepathserver[$i]),$results"
        $i++
    }
    $results = $new_results

    if ($choice -eq 6)
    {
        $new_results_iis_backend =@()
        $csvheadline = "sourcelogfilepath,$($csvheadline.line)"
        $i = 0

        foreach ($entry in $results_iis_backend)
        {
            $new_results_iis_backend += "$($sourcelogfilepathserver[$i]),$results_iis_backend"
            $i++
        }
        $results_iis_backend = $new_results_iis_backend
    }

}
else
{
    $csvheadline = $csvheadline.line    
    $results = $results.line
}

#$org_results = $results

#Areyousure function. Alows user to select y or n when asked to exit. Y exits and N returns to main menu.  
 function areyousure {$areyousure = read-host "Are you sure you want to exit? (y/n)"  
           if ($areyousure -eq "y"){exit}  
           if ($areyousure -eq "n"){mainmenu}  
           else {write-host -foregroundcolor red "Invalid Selection"   
                 areyousure  
                }  
} 

<#undo changes and filters
function undo{
$results = $org_results
$fl = "0"
mainmenu 
}
#>

#fullist
function fulllist{
 cls
 $fl = "1"
mainmenu 
}

function gridview{
cls
$results | Out-GridView -Title "Logging Results"
    
    if ($results_iis_backend)
    {
        $results_iis_backend | Out-GridView -Title "Logging Results (IIS BackEnd)"
    }

mainmenu
}

function exportcsv{
cls
$filepath = "c:\temp\$(get-date -format yyyyMMdd_HHmm)exchangelogfilesearcher-$logtype-export.csv"
$csvheadline | Out-File -FilePath $filepath
$results | Out-File -FilePath $filepath -Append

    if ($results_iis_backend)
    {
        $filepath = "c:\temp\$(get-date -format yyyyMMdd_HHmm)exchangelogfilesearcher-IISLogsBackEnd-export.csv"
        $csvheadline | Out-File -FilePath $filepath
        $results_iis_backend | Out-File -FilePath $filepath -Append
    }

[bool]$csvexportdone = $true
mainmenu
echo ""
}

$wrap = "0";

#Mainmenu function. Contains the screen output for the menu and waits for and handles user input.  
function mainmenu{
#cls

Write-Host -ForegroundColor White "Found $($results.count) Results"
if ($choice -eq 6)
{
    Write-Host -ForegroundColor Cyan "$($results_iis_backend.count) IIS_BackEnd Entries found!"
}
 
if ($fl -eq "1")
{
    $results | fl
    $fl = "0"
}

 echo ""
 echo "---------------------------------------------------------"  
 echo ""
 echo "    1. Show FullList"
 echo "    2. Output to GridView"
 echo "    3. Export to CSV"
 #echo "    9. undo Filter"  
 echo "    0. Exit"  
 echo ""
 echo "---------------------------------------------------------"  
 echo ""  

 echo ""
 if ($csvexportdone)
 {
    write-host -ForegroundColor Yellow "CSV-Export done: " -NoNewline
    Write-Host -ForegroundColor White "$filepath" 
 }
 echo "" 
 $answer = read-host "Please Make a Selection"  
 if ($answer -eq 1){fulllist}
 if ($answer -eq 2){gridview}
 if ($answer -eq 3){exportcsv}
 #if ($answer -eq 9){undo}  
 if ($answer -eq 0){areyousure}
 else {write-host -ForegroundColor red "Invalid Selection"  
       sleep 5  
       mainmenu  
      }  
                }  
 mainmenu

