#.
# Look near the bottom of the script to configure alarm report emails / notifications.
#
# The report will be saved in this file:
param([string]$File = "C:\CMO\ADWorkstations_Report.html")

Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" # Import the ConfigurationManager.psd1 module 
Set-Location "EU1:" # Set the current location to be the site code.


"<html>
<head>
<link rel=stylesheet href=https://maxcdn.bootstrapcdn.com/bootstrap/3.3.2/css/bootstrap.min.css>
<title>$env:COMPUTERNAME</title>
<style>
html, body
{
font-family: tahoma, arial;
font-size: 14px;
margin: 5px;
padding: 0px;
color: #2E1E2E;
}
table
{
border-collapse: collapse;
    width: 100%;
}
th
{
border: 1px solid #8899AA;
padding: 3px 7px 2px 7px;
font-size: 1.1em;
text-align: left;
padding-top: 5px;
padding-bottom: 4px;
background-color: #AABBCC;
color: #ffffff;
}
td
{
border: 1px solid #8899AA;
padding: 3px 7px 2px 7px;
    overflow: hidden;
}
h2
{
    text-align: center;
font-size: 22px;
    text-shadow: 1px 1px 1px rgba(150, 150, 150, 0.5);
}
h1
{
    margin-top: 20px;
    text-align: center;
font-size: 25px;
    text-shadow: 1px 1px 1px rgba(150, 150, 150, 0.5);
}
pre {
    white-space: pre-wrap;
    white-space: -moz-pre-wrap;
    white-space: -pre-wrap;
    white-space: -o-pre-wrap;
    word-wrap: break-word;
}
#sysinfo
{
    width: 49% !important;
    float: left;
    margin-bottom: 0px;
}
#action
{
    width: 49% !important;
    float: right;
}
#menu 
{
    position: fixed;
    right: 0;
    left: 0;
    top: 0;
    width: 100%;
    height: 25px;
    background: #AABBCC;
    color: #FFFFFF;
    text-align: center;
    overflow: hidden;
}
#menu a
{
    color: #FFFFFF;
    font-weight: bold;
}
@media screen and (max-width: 1010px)
{
    #sysinfo
    {
        float: none;
        margin-bottom: 20px;
        width: 100% !important;
    }
    #action
    {
        width: 100% !important;
        float: none;
    }
}
</style>
</head>
<body>
<div id='menu'>
<a href=#Home>Home</a> | <a href=#Total>Total Computers</a> | <a href=#Disabled>Disabled Computers</a> | <a href=#Inactive>Inactive Computers</a> | <a href=#Desktop>Desktop Computers</a> | <a href=#Laptop>Laptop Computers</a> | <a href=#Incorrect>Incorrect Naming</a>  | <a href=#Incorrect>Inactive</a>  | <a href=#Incorrect>NoAgent</a>   
</div>
<a name='sysinfo'></a><h1>$env:COMPUTERNAME Active Directory Workstation Report</h1>
" > $File

#Active Directory Statistics
$d = [DateTime]::Today.AddDays(-30)
$date = [DateTime]::Today.AddDays(-2)

$staleADWS = Get-ADComputer -Filter 'PasswordLastSet -le $d -and OperatingSystem -notlike "Windows Server*"' -SearchBase "OU=AN1-SouthAfrica,DC=za,DC=if,DC=atcsg,DC=net" -Properties * | Sort-Object DistinguishedName | Select DNSHostName,OperatingSystem,Description,Created,LastLogonDate,@{n='lastlogon';e={[DateTime]::FromFileTime($_.LastLogon)}},Enabled,IPv4Address | Sort-Object LastLogonDate
$rebootReq = Get-ADComputer -Filter 'LastLogonDate -le $date -and OperatingSystem -notlike "Windows Server*"' -SearchBase "OU=AN1-SouthAfrica,DC=za,DC=if,DC=atcsg,DC=net" -Properties * | Sort-Object DistinguishedName | Select DNSHostName,OperatingSystem,Description,Created,LastLogonDate,Enabled,IPv4Address,DistinguishedName
$rebooted = Get-ADComputer -Filter 'LastLogonDate -ge $date -and OperatingSystem -notlike "Windows Server*"' -SearchBase "OU=AN1-SouthAfrica,DC=za,DC=if,DC=atcsg,DC=net" -Properties * | Sort-Object DistinguishedName | Select DNSHostName,OperatingSystem,Description,Created,LastLogonDate,Enabled,IPv4Address,DistinguishedName

$totalADWS = Get-ADComputer -Filter 'OperatingSystem -notlike "Windows Server*"' -SearchBase "OU=AN1-SouthAfrica,DC=za,DC=if,DC=atcsg,DC=net" -Properties * | Select DNSHostName,OperatingSystem,Description,Created,LastLogonDate,@{n='lastlogon';e={[DateTime]::FromFileTime($_.LastLogon)}},IPv4Address | Sort-Object IPv4Address, LastLogonDate

$ws = Get-ADComputer -Filter 'OperatingSystem -notlike "Windows Server*" -and DNSHostName -like "AN1D*"' -SearchBase "OU=AN1-SouthAfrica,DC=za,DC=if,DC=atcsg,DC=net" -Properties * | Select DNSHostName,OperatingSystem,Description,Created,LastLogonDate,@{n='lastlogon';e={[DateTime]::FromFileTime($_.LastLogon)}},IPv4Address | Sort-Object IPv4Address, LastLogonDate
$laptop = Get-ADComputer -Filter 'OperatingSystem -notlike "Windows Server*" -and DNSHostName -like "AN1LT*"' -SearchBase "OU=AN1-SouthAfrica,DC=za,DC=if,DC=atcsg,DC=net" -Properties * | Select DNSHostName,OperatingSystem,Description,Created,LastLogonDate,@{n='lastlogon';e={[DateTime]::FromFileTime($_.LastLogon)}},IPv4Address | Sort-Object IPv4Address, LastLogonDate
$other = Get-ADComputer -Filter 'OperatingSystem -notlike "Windows Server*" -and DNSHostName -notlike "AN1LT*" -and DNSHostName -notlike "AN1DT*"' -SearchBase "OU=AN1-SouthAfrica,DC=za,DC=if,DC=atcsg,DC=net" -Properties * | Select DNSHostName,OperatingSystem,Description,Created,LastLogonDate,@{n='lastlogon';e={[DateTime]::FromFileTime($_.LastLogon)}},IPv4Address | Sort-Object IPv4Address, LastLogonDate

$disabledWS = Get-ADComputer -Filter 'OperatingSystem -notlike "Windows Server*"' -SearchBase "OU=AN1-SouthAfrica,DC=za,DC=if,DC=atcsg,DC=net"| where-object {$_.Enabled -notlike "True"} | Select DNSHostName,OperatingSystem,Description,Created,LastLogonDate,Enabled,IPv4Address,DistinguishedName

$active = $totalADWS.count - ($disabledWS.count + $staleADWS.Count) 

#Windows 7
$totalWin7 = Get-ADComputer -Filter 'OperatingSystem -notlike "Windows Server*"' -SearchBase "OU=AN1-SouthAfrica,DC=za,DC=if,DC=atcsg,DC=net" -Properties * | Where-Object {$_.OperatingSystem -Like 'Windows 7*'} | Select DNSHostName,OperatingSystem,Description,Created,LastLogonDate,Enabled,IPv4Address,DistinguishedName


#Windows 10
$totalWin10 = Get-ADComputer -Filter 'OperatingSystem -notlike "Windows Server*"' -SearchBase "OU=AN1-SouthAfrica,DC=za,DC=if,DC=atcsg,DC=net" -Properties * | Where-Object {$_.OperatingSystem -Like 'Windows 10*'} | Select DNSHostName,OperatingSystem,Description,Created,LastLogonDate,Enabled,IPv4Address,DistinguishedName

#Percentage of Healthy Agents
$adWSHealthyNUM = ($active / $totalADWS.count)

$adWSperc = "{0:p}" -f $adWSHealthyNUM

#Waiting Approval
$WSnotApproved = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.IsApproved -eq '0' -and $_.DeviceOS -like '*Workstation*'}

#Obselete
$WSObselete = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.IsObselete -like 'True' -and $_.DeviceOS -like '*Workstation*'} | Select Name

#SCCM Healthy Agent Statistics
$healthyAgentsWS = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.ClientActiveStatus -like '1' -and $_.DeviceOS -notlike '*Server*'} | Select Name

#SCCM Inactive Agent Statistics
$inActiveWSAgent = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.ClientActiveStatus -eq '0' -and $_.DeviceOS -like '*Workstation*' -and $_.DeviceOS -notlike '*Mac*' -and $_.DeviceOS -notlike '*SLES 11*'} | Sort LastActiveTime | Select Name, UserName, SiteCode, LastMPServerName, LastInstallationError, LastActiveTime

#SCCM No Agent Statistics
$noAgentsWS = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.IsClient -like 'False' -and $_.DeviceOS -like '*Workstation*' -and $_.DeviceOS -notlike '*Mac*' -and $_.DeviceOS -notlike '*SLES 11*'} |  Sort Name | Select Name, SiteCode, LastInstallationError


#SCCM Totals
$cmWSTotals = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.DeviceOS -like '*Workstation*' -and $_.DeviceOS -notlike '*Mac*' -and $_.DeviceOS -notlike '*SLES 11*'}

#DeviceOS
$cmWin7 = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.DeviceOS -like '*Workstation 6.1*' -and $_.DeviceOS -notlike '*Mac*' -and $_.DeviceOS -notlike '*SLES 11*'}
$cmWin10 = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.DeviceOS -like '*Workstation 10*' -and $_.DeviceOS -notlike '*Mac*' -and $_.DeviceOS -notlike '*SLES 11*'}

#Platform
$cmWorkstation = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.Name -like 'AN1D*' -and $_.DeviceOS -notlike '*Mac*' -and $_.DeviceOS -notlike '*SLES 11*'}
$cmLaptop = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.Name -like 'AN1LT*'}
$cmIncorrect = $cmLaptop = Get-CMDevice -CollectionName 'ZA_AN1-SOUTHAFRICA-UK-I-SA-Computers' | Where-Object {$_.Name -notlike 'AN1LT*' -and $_.Name -notlike 'AN1D*' -and $_.DeviceOS -notlike '*Server*' -and $_.DeviceOS -notlike '*SLES 11*'}

#Percentage of Healthy Agents
$WSHealthyNUM = ($healthyAgentsWS.count / $cmWSTotals.count)

$WSperc = "{0:p}" -f $WSHealthyNUM

##########################Start of Overall SCCM Statistics Table##########################################

$tabName = “SCCMStatistics”

#Create Table object
$table = New-Object system.Data.DataTable “$tabName”

#Define Columns

$col1 = New-Object system.Data.DataColumn 'Machine Type',([string])
$col2 = New-Object system.Data.DataColumn 'Total in AD',([string])
$col3 = New-Object system.Data.DataColumn 'Active',([string])
$col4 = New-Object system.Data.DataColumn 'Inactive 30 Days',([string])
$col5 = New-Object system.Data.DataColumn 'Disabled',([string])
$col6 = New-Object system.Data.DataColumn 'Reboot Required',([string])
$col7 = New-Object system.Data.DataColumn 'Rebooted in 2 days',([string])
$col8 = New-Object system.Data.DataColumn 'Percentage Active Device',([string])
$col9 = New-Object system.Data.DataColumn 'Windows 7',([string])
$col10 = New-Object system.Data.DataColumn 'Windows 10',([string])
$col11= New-Object system.Data.DataColumn 'Desktop',([string])
$col12= New-Object system.Data.DataColumn 'Laptop',([string])
$col13= New-Object system.Data.DataColumn 'Incorrect Name',([string])

#Add the Columns
$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)
$table.columns.add($col7)
$table.columns.add($col8)
$table.columns.add($col9)
$table.columns.add($col10)
$table.columns.add($col11)
$table.columns.add($col12)
$table.columns.add($col13)

#Create a row
$row1 = $table.NewRow()


$row1.'Machine Type' = “Windows Workstations” 
$row1.'Total in AD' = $totalADWS.Count
$row1.'Active' = $active
$row1.'Inactive 30 Days' = $StaleADWS.Count
$row1.'Disabled' = $disabledWS.count
$row1.'Reboot Required' = $rebootReq.count
$row1.'Rebooted in 2 days' = $rebooted.count
$row1.'Percentage Active Device' = $adWSperc
$row1.'Windows 7' = $totalWin7.count
$row1.'Windows 10' = $totalWin10.count
$row1.'Desktop' = $ws.count
$row1.'Laptop' = $laptop.Count
$row1.'Incorrect Name' = $other.Count

#Add the row to the table
$table.Rows.Add($row1)


##########################Start of SCCM Site Statistics Table##########################################

$tabNames = “SCCM Site Statistics”

#Create Table object
$tables = New-Object system.Data.DataTable “$tabNames”

#Define Columns

$cols1 = New-Object system.Data.DataColumn 'Site Name',([string])
$cols2 = New-Object system.Data.DataColumn 'Total in SCCM',([string])
$cols3 = New-Object system.Data.DataColumn 'Healthy in SCCM',([string])
$cols4 = New-Object system.Data.DataColumn 'Inactive in SCCM',([string]) 
$cols5 = New-Object system.Data.DataColumn 'No Agents in SCCM',([string])
$cols6 = New-Object system.Data.DataColumn 'Waiting Approval',([string])
$cols7 = New-Object system.Data.DataColumn 'Obselete',([string])
$cols8 = New-Object system.Data.DataColumn 'Percentage of Healthy Cients',([string])
$cols9 = New-Object system.Data.DataColumn 'Windows 7',([string])
$cols10 = New-Object system.Data.DataColumn 'Windows 10',([string])
$cols11 = New-Object system.Data.DataColumn 'Desktop',([string])
$cols12 = New-Object system.Data.DataColumn 'Laptop',([string])
$cols13 = New-Object system.Data.DataColumn 'Incorrect Name',([string])

#Add the Columns
$tables.columns.add($cols1)
$tables.columns.add($cols2)
$tables.columns.add($cols3)
$tables.columns.add($cols4)
$tables.columns.add($cols5)
$tables.columns.add($cols6)
$tables.columns.add($cols7)
$tables.columns.add($cols8)
$tables.columns.add($cols9)
$tables.columns.add($cols10)
$tables.columns.add($cols11)
$tables.columns.add($cols12)
$tables.columns.add($cols13)


#Create a row
$rows = $tables.NewRow()

#Enter data in the row
$rows.'Site Name' = “SGPCP-South Africa” 
$rows.'Total in SCCM' = $cmWSTotals.count
$rows.'Healthy in SCCM' = $healthyAgentsWS.count
$rows.'Inactive in SCCM' = $inActiveWSAgent.count
$rows.'No Agents in SCCM' = $noAgentsWS.count
$rows.'Waiting Approval' = $WSnotApproved.count
$rows.'Obselete' = $WSObselete.count
$rows.'Percentage of Healthy Cients' = $WSperc
$rows.'Windows 7' = $cmWin7.count 
$rows.'Windows 10' = $cmWin10.count
$rows.'Desktop' = $cmWorkstation.count 
$rows.'Laptop' = $cmLaptop.count
$rows.'Incorrect Name' = $cmIncorrect.count  

#Add the row to the table
$tables.Rows.Add($rows)

#########################END TABLE###########################

"<a name='Home'></a><font color=red><b><h2>SGCP - Active Directory Overall Statistics</h2><font color=red><b>" >> $File
Write-Output "* SCCM Overall Statistics"
$table | Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | ConvertTo-Html -Fragment >> $File

"<a name='Home'></a><font color=red><b><h2>SGCP - CMCB Overall Statistics</h2><font color=red><b>" >> $File
Write-Output "* SCCM Overall Statistics"
$tables | Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | ConvertTo-Html -Fragment >> $File

"<a name='Disabled'></a><h2>Disabled Workstations</h2>" >> $File
Write-Output "* Workstation Servers Disabled in Active Directory"
$disabledWS | ConvertTo-Html -Fragment >> $File

"<a name='Inactive'></a><h2>Workstations Inactive in AD</h2>" >> $File
Write-Output "* Workstation Inactive in Active Directory for over 30 Days"
$staleADWS | ConvertTo-Html -Fragment >> $File

"<a name='Total'></a><h2>Total Workstations</h2>" >> $File
Write-Output "* Workstation Inactive in Active Directory for over 35 Days"
$totalADWS | ConvertTo-Html -Fragment >> $File

"<a name='Desktop'></a><h2>Desktops</h2>" >> $File
Write-Output "* Workstation Inactive in Active Directory for over 35 Days"
$ws | ConvertTo-Html -Fragment >> $File

"<a name='Laptop'></a><h2>Laptops</h2>" >> $File
Write-Output "* Workstation Inactive in Active Directory for over 35 Days"
$laptop | ConvertTo-Html -Fragment >> $File

"<a name='Incorrect'></a><h2>Incorrect Naming Convention</h2>" >> $File
Write-Output "* Workstation Inactive in Active Directory for over 35 Days"
$other | ConvertTo-Html -Fragment >> $File

"<a name='Inactive'></a><h2>Inactive Agents</h2>" >> $File
Write-Output "* Workstation Inactive in Active Directory for over 35 Days"
$inActiveWSAgent | ConvertTo-Html -Fragment >> $File

"<a name='NoAgent'></a><h2>No Agents</h2>" >> $File
Write-Output "* Workstation Inactive in Active Directory for over 35 Days"
$noAgentsWS | ConvertTo-Html -Fragment >> $File

"<a name='Troubleshooting'></a><font color=red><b><h2>UCS Solutions</h2></b></font>" >> $File
"<p><i>Powered by CMO</i></p>" >> $File

$date = Get-Date
"<p><i>Report produced: $date</i></p>" >> $File

if((Get-Content $File | Select-String -Pattern "color=red"))
{
    Write-Output "*** Alarms were raised!"

#Send Mail
$body = Get-Content C:\CMO\ADWorkstations_Report.html -Raw
$params = @{ 
    Attachments = 'C:\CMO\ADWorkstations_Report.html'
    Body = $body
    BodyAsHtml = $true 
    Subject = 'SGCP - AD Workstations Report' 
    From = 'CMO.Reporting@saint-gobain.com' 
    To = 'igshaan.anthony@bcx.co.za'
    Cc = 'adolph.nkosi@bcx.co.za','Shweta.Yermal@saint-gobain.com'
    SmtpServer = '10.155.65.59' 
    Port = 25 
 } 
 
Send-MailMessage @params
}

Write-Output "Done! Report at: $File"
