<#  
.SYNOPSIS  
	This script reports detailed queue information on Exchange 2016 servers

.PARAMETER ServerFilter
The name or partial name of the Servers to query valid values such as MSGSVR or MSGCONSVR2301

.PARAMETER ReportPath
Directory location to create the HTML report

.NOTES  
  Version      				: 0.1
  Rights Required			: Exchange View Only Admin/Local Server Administrator
  Exchange Version			: 2016/2013 (last tested on Exchange 2016 CU14/Windows 2012R2)
  Authors       			: Steven Snider (stevesn@microsoft.com) (HTML template borrowed from other examples on internet)
  Last Update               : Oct 24 2019

.VERSION
  0.1 - Initial Version for connecting Internal Exchange Servers
	
#>

Param(
   [Parameter(Mandatory=$false)] [string] $ServerFilter="MSGSVR",
   [Parameter(Mandatory=$false)] [string] $ReportPath=$env:TEMP

)


<#Needs:


#>


If (-Not($UserCredential)) {
    $UserCredential = Get-Credential
}

#region Verifying Administrator Elevation
Write-Host Verifying User permissions... -ForegroundColor Yellow
Start-Sleep -Seconds 2
#Verify if the Script is running under Admin privileges
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
  [Security.Principal.WindowsBuiltInRole] "Administrator")) 
{
  Write-Warning "You do not have Administrator rights to run this script.`nPlease re-run this script as an Administrator!"
  Write-Host 
  Break
}
#endregion

#region Script Information

Write-Host "--------------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host "Exchange Queue Report" -ForegroundColor Green
Write-Host "Version: 0.1" -ForegroundColor Green
Write-Host "--------------------------------------------------------------" -BackgroundColor DarkGreen
#endregion

$FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $ReportPath+"\ExQueueReport-"+$FileDate+".html"
[Void](New-Item -ItemType file $ServicesFileName -Force)

[string]$search = "(&(objectcategory=computer)(cn=$serverfilter*))"
$ExchangeServers = ([adsisearcher]$search).findall() | %{$_.properties.name} | sort

$ServersList = @()
$ServersList = $ExchangeServers

If ($serverslist.length -eq 0) {
    Write-Host "Filter returned zero servers.  Please adjust filter and try again." -ForegroundColor Red
    Exit
}

#### Building HTML File ####
Function writeHtmlHeader
{
    param($fileName)
    $date = ( get-date ).ToString('MM/dd/yyyy')
    Add-Content $fileName "<html>"
    Add-Content $fileName "<head>"
    Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
    Add-Content $fileName '<title>Exchange Queue Report</title>'
    add-content $fileName '<STYLE TYPE="text/css">'
    add-content $fileName  "<!--"
    add-content $fileName  "td {"
    add-content $fileName  "font-family: Segoe UI;"
    add-content $fileName  "font-size: 11px;"
    add-content $fileName  "border-top: 1px solid #1E90FF;"
    add-content $fileName  "border-right: 1px solid #1E90FF;"
    add-content $fileName  "border-bottom: 1px solid #1E90FF;"
    add-content $fileName  "border-left: 1px solid #1E90FF;"
    add-content $fileName  "padding-top: 0px;"
    add-content $fileName  "padding-right: 0px;"
    add-content $fileName  "padding-bottom: 0px;"
    add-content $fileName  "padding-left: 0px;"
    add-content $fileName  "}"
    add-content $fileName  "body {"
    add-content $fileName  "margin-left: 5px;"
    add-content $fileName  "margin-top: 5px;"
    add-content $fileName  "margin-right: 0px;"
    add-content $fileName  "margin-bottom: 10px;"
    add-content $fileName  ""
    add-content $fileName  "table {"
    add-content $fileName  "border: thin solid #000000;"
    add-content $fileName  "}"
    add-content $fileName  "-->"
    add-content $fileName  "</style>"
    add-content $fileName  "</head>"
    add-content $fileName  "<body>"
    add-content $fileName  "<table width='100%'>"
    add-content $fileName  "<tr bgcolor='#336699 '>"
    add-content $fileName  "<td colspan='7' height='25' align='center'>"
    add-content $fileName  "<font face='Segoe UI' color='#FFFFFF' size='4'>Exchange Queue Report - $date</font>"
    add-content $fileName  "</td>"
    add-content $fileName  "</tr>"
    add-content $fileName  "</table>"
}

Function writeTableHeader
{
    param($fileName)
    Add-Content $fileName "<tr bgcolor=#0099CC>"
    Add-Content $fileName "<td width='30%' align='center'><font color=#FFFFFF>Identity</font></td>"
    Add-Content $fileName "<td width='5%' align='center'><font color=#FFFFFF>Status</font></td>"
    Add-Content $fileName "<td width='5%' align='center'><font color=#FFFFFF>MessageCount</font></td>"
    Add-Content $fileName "<td width='5%' align='center'><font color=#FFFFFF>Velocity</font></td>"
    Add-Content $fileName "<td width='20%' align='center'><font color=#FFFFFF>NextHopDomain</font></td>"
    Add-Content $fileName "<td width='25%' align='center'><font color=#FFFFFF>LastError</font></td>"
    Add-Content $fileName "<td width='5%' align='center'><font color=#FFFFFF>RetryCount</font></td>"
    Add-Content $fileName "<td width='5%' align='center'><font color=#FFFFFF>LastRetryTime</font></td>"
    Add-Content $fileName "</tr>"
}

Function writeHtmlFooter
{
    param($fileName)
    Add-Content $fileName "</body>"
    Add-Content $fileName "</html>"
}

Function writeServiceInfo
{
    param($fileName,$Identity,$Status,$MessageCount,$Velocity,$NextHopDomain,$LastError,$RetryCount,$LastRetryTime)

    
     Add-Content $fileName "<tr>"
     Add-Content $fileName "<td align='center'>$Identity</td>"
     Add-Content $fileName "<td>$Status</td>"
     If ($MessageCount -gt 50) {
         Add-Content $fileName "<td bgcolor='#FF0000'>$MessageCount</td>"
     }
     ElseIf ($MessageCount -gt 10) {
         Add-Content $fileName "<td bgcolor='#FFFF00'>$MessageCount</td>"
     }
     Else {
         Add-Content $fileName "<td>$MessageCount</td>"
     }
     If ($Velocity -gt 0) {
         Add-Content $fileName "<td>$Velocity</td>"
     }
     ElseIf ($Velocity -lt 0) {
         Add-Content $fileName "<td bgcolor='#FF0000'>$Velocity</td>"
     }
     Else {
         Add-Content $fileName "<td>$MessageCount</td>"
     }
     Add-Content $fileName "<td>$NextHopDomain</td>"
     Add-Content $fileName "<td align='center'>$LastError</td>"
     Add-Content $fileName "<td align='center'>$RetryCount</td>"
     Add-Content $fileName "<td>$LastRetryTime</td>"
}

Function sendEmail
    { param($from,$to,$subject,$smtphost,$htmlFileName)
        $body = Get-Content $htmlFileName
        $smtp= New-Object System.Net.Mail.SmtpClient $smtphost
        $msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
        $msg.isBodyhtml = $true
        $smtp.send($msg)
    }

########################### Main Script ###################################
writeHtmlHeader $ServicesFileName

        Add-Content $ServicesFileName "<table width='100%'><tbody>"
        Add-Content $ServicesFileName "<tr bgcolor='#0099CC'>"
        Add-Content $ServicesFileName "</tr>"

        WriteTableHeader $ServicesFileName

foreach ($Server in $ServersList)
{       
        
        try
        {
            Write-Host "Querying queues on server " $Server
            $ExQueues = Get-Queue -Server $Server -Filter {MessageCount -gt -1 -and DeliveryType -ne "ShadowRedundancy"}

        
        } # end Try
        catch
        {
            Write-Host
            Write-Host "Error Connecting to server " $Server ", Please verify connectivity and permissions" -ForegroundColor Red
            Add-Content $ServicesFileName "<tr bgcolor='#0099CC'>"
            Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=11><font face='segoe ui' color='#FF0000' size='2'>Error Connecting to server $Server, Please verify connectivity and permissions</font></td>"
            Add-Content $ServicesFileName "</tr>"

            Continue
        } #end catch
        
        foreach ($item in $ExQueues)
        {
            writeServiceInfo $ServicesFileName $item.Identity $item.Status $item.MessageCount $item.Velocity $item.NextHopDomain $item.LastError $item.RetryCount $item.LastRetryTime
        }
        
}
  
  
       
   Add-Content $ServicesFileName "</table>"

writeHtmlFooter $ServicesFileName

### Configuring Email Parameters
#sendEmail from@domain.com to@domain.com "Queue State Report - $Date" SMTPS_SERVER $ServicesFileName

#Closing HTML
writeHtmlFooter $ServicesFileName
Write-Host "`n`nThe File was generated at the following location: $ServicesFileName `n`nOpenning file..." -ForegroundColor Cyan
Invoke-Item $ServicesFileName


