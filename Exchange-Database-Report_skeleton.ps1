#Creates the password object
$exchuname = "domain\service_account"
$AESKey = Get-Content "\\<servername>\<folder>\<subfolder>\Key_service_account.key"
$pass = Get-Content "\\<servername>\<folder>\<subfolder>\service_account_pw.txt"
$securePwd = $pass | ConvertTo-SecureString -Key $AESKey
$exadmin = New-Object System.Management.Automation.PSCredential -ArgumentList $exchuname, $securePwd



#Imports Exchange
$Session = New-PSSession -Name Exchange -ConfigurationName Microsoft.Exchange -ConnectionUri "http://yourexchangeserver.domain.com/powershell" -Authentication Kerberos -Credential $exadmin
Import-PSSession $Session >> $null -AllowClobber

$Mailboxes = Get-Mailbox -ResultSize Unlimited

#Sets the size for "unlimited" mailboxes; can be changed if needs be 
$Unlimited = 3221225472
$UnlimitedWarning = 2147483648

#Starts Mailbox Report
$Report = Foreach ($Mail in $Mailboxes) {
    $Policy = Get-CASMailbox -Identity $Mail.Alias
    $Devices = Get-MobileDevice -Mailbox $Mail.alias
    $Stats = $Mail | Get-MailboxStatistics

    #sets mailbox length in bytes
    if ($mail.ProhibitSendQuota -like "Unlimited") {
        $mail.ProhibitSendQuota = $Unlimited
    }
    else {
        $mail.ProhibitSendQuota = ($Mail.ProhibitSendQuota -replace '.*\(', '') -replace '\D+', ''
    }
    if ($mail.ProhibitSendReceiveQuota -like "Unlimited") {
        $mail.ProhibitSendReceiveQuota = $Unlimited   
    }
    else {
        $mail.ProhibitSendReceiveQuota = ($Mail.ProhibitSendQuota -replace '.*\(', '') -replace '\D+', ''
    }
    if ($mail.IssueWarningQuota -like "Unlimited") {
        $mail.IssueWarningQuota = $UnlimitedWarning  
    }
    else {
        $mail.IssueWarningQuota = ($Mail.IssueWarningQuota -replace '.*\(', '') -replace '\D+', ''
    }
    if ($Null -eq $Stats.totalitemsize.value) {
        $stats.totalitemsize.value = 0     
    }
    else {
        $stats.totalitemsize.value = ($Stats.totalitemsize.value -replace '.*\(', '') -replace '\D+', ''
    }
    if ($Null -eq $Stats.TotalDeletedItemSize.value) {
        $stats.TotalDeletedItemSize.value = 0     
    }
    else {
        $stats.TotalDeletedItemSize.value = ($Stats.TotalDeletedItemSize.value -replace '.*\(', '') -replace '\D+', ''
    }
    
    #Creates percentage of mailbox.
    $PercentFull = [math]::Round((($Stats.totalitemsize.value / $Mail.ProhibitSendReceiveQuota) * 100), 2)
    [pscustomobject]@{
        Name                     = $mail.Name 
        Username                 = $Mail.Alias
        Database                 = $Stats.Database
        Policy                   = $Policy.ActiveSyncMailboxPolicy
        MobileDeviceCount        = [double]$Devices.count
        ProhibitSendQuota        = [double]$Mail.ProhibitSendQuota
        ProhibitSendReceiveQuota = [double]$Mail.ProhibitSendReceiveQuota
        IssueWarningQuota        = [double]$Mail.IssueWarningQuota
        TotalItemSize            = [double]$stats.totalitemsize.value
        TotalDeletedItemSize     = [double]$stats.TotalDeletedItemSize.value
        ItemCount                = [double]$Stats.ItemCount 
        PercentFull              = $PercentFull
    }
}

#Grabs the database info for the database report.
$DatabaseInfo = Get-MailboxDatabase -Status | Select-Object Name, @{l = 'DatabaseSize'; e = { ($_.DatabaseSize -replace '.*\(', '') -replace '\D+', '' } }, @{l = 'AvailableNewMailboxSpace'; e = { ($_.AvailableNewMailboxSpace -replace '.*\(', '') -replace '\D+', '' } }, @{l = 'UserCount'; e = { ($report.Database.name -like $_.name).count } }

#Starts the database report setup.
$Databasereport = foreach ($DBI in $DatabaseInfo) {
    [pscustomobject]@{
        DatabaseName             = $dbi.Name
        DatabaseSize             = [Math]::Round(($DBI.DatabaseSize / 1gb), 2)
        AvailableNewMailboxSpace = [Math]::Round(($DBI.AvailableNewMailboxSpace / 1gb), 2)
        UserCount                = $DBI.UserCount
        MailboxAverage           = [math]::Round((($DBI.DatabaseSize / $DBI.UserCount) / 1gb), 2)
        LargestMailboxUsername   = ($Report | Where-Object { $_.Database.name -contains $DBI.name } | Sort-Object -Property TotalItemSize -Descending | Select-Object -First 1).Username
        LargestMailboxUserSize   = "$([math]::Round((($Report | Where-Object {$_.Database.name -contains $DBI.name} | Sort-Object -Property TotalItemSize -Descending | Select-Object -First 1).TotalItemSize/1gb),2))/gb"
    }
}

$DBReport = $Databasereport | ConvertTo-HTML -As Table -Fragment
$userReport = $Report | ConvertTo-HTML -As Table -Fragment

$Htmlbody = @" 
<html> 
<head>
<style>
body {
    Color: #252525;
    font-family: Verdana,Arial;
    font-size:11pt;
}
table {border: 1px solid rgb(104,107,112); text-align: left;}
th {background-color: #d2e3f7;border-bottom:2px solid rgb(79,129,189);text-align: left;}
tr {border-bottom:2px solid rgb(71,85,112);text-align: left;}
td {border-bottom:1px solid rgb(99,105,112);text-align: left;}
h1 {
    text-align: left;
    color:#5292f9;
    Font-size: 34pt;
    font-family: Verdana, Arial;
}
h2 {
    text-align: left;
    color:#323a33;
    Font-size: 20pt;
}
h3 {
    text-align: center;
    color:#211b1c;
    Font-size: 15pt;
}
h4 {
    text-align: left;
    color:#2a2d2a;
    Font-size: 15pt;
}
h5 {
    text-align: center;
    color:#2a2d2a;
    Font-size: 12pt;
}
a:link {
    color:#0098e5;
    text-decoration: underline;
    cursor: auto;
    font-weight: 500;
}
a:visited {
    color:#05a3b7;
    text-decoration: underline;
    cursor: auto;
    font-weight: 500;
}
</style>
</head>
<body>
<h3>Exchange DB Report</h3> 
<br>
$DBReport
<br>


<br>
<br>
Thank you<br>
Exchange Administrator<br>

</body> 
</html> 
"@  

$Htmlbody2 = @" 
<html> 
<head>
<style>
body {
    Color: #252525;
    font-family: Verdana,Arial;
    font-size:11pt;
}
table {border: 1px solid rgb(104,107,112); text-align: left;}
th {background-color: #d2e3f7;border-bottom:2px solid rgb(79,129,189);text-align: left;}
tr {border-bottom:2px solid rgb(71,85,112);text-align: left;}
td {border-bottom:1px solid rgb(99,105,112);text-align: left;}
h1 {
    text-align: left;
    color:#5292f9;
    Font-size: 34pt;
    font-family: Verdana, Arial;
}
h2 {
    text-align: left;
    color:#323a33;
    Font-size: 20pt;
}
h3 {
    text-align: center;
    color:#211b1c;
    Font-size: 15pt;
}
h4 {
    text-align: left;
    color:#2a2d2a;
    Font-size: 15pt;
}
h5 {
    text-align: center;
    color:#2a2d2a;
    Font-size: 12pt;
}
a:link {
    color:#0098e5;
    text-decoration: underline;
    cursor: auto;
    font-weight: 500;
}
a:visited {
    color:#05a3b7;
    text-decoration: underline;
    cursor: auto;
    font-weight: 500;
}
</style>
</head>
<body>
<h3>Exchange User Report</h3> 
<br>
$userReport
<br>


<br>
<br>
Thank you<br>
Exchange Administrator<br>

</body> 
</html> 
"@  

$mailadmins = "someemail@yourdomain.com","anotheremail@yourdomain.com"

Send-MailMessage -To $mailadmins -From "service_account@yourdomain.com" -Subject "Exchange DB Report" -BodyAsHtml $Htmlbody -SmtpServer your.smtpserver.yourdomain.com
Send-MailMessage -To $mailadmins -From "service_account@yourdomain.com" -Subject "Exchange User Mailbox Report" -BodyAsHtml $Htmlbody2 -SmtpServer your.smtpserver.yourdomain.com