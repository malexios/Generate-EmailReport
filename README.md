# Description

Accepts an object, converts it to a table and generates an HTML report containing that table. The returned object can be used as the body of an email or saved as a document.

# Example

$ReportDescription = "List of users in the 'Administrators' AD group."
$EmailFooter = "This report ran from $env:computername as a scheduled task as the user $env:username at $(Get-Date -Format hh:MM:ss ).<br>Please contact administrator@domain.com with any questions."

$HTML = Generate-EmailReport -ReportName "Domain Admins"`
    -ReportDescription $ReportDescription`
    -EmailFooterText $EmailFooter`
    -HTMLBodyBackgroundColor "white"`
    -TableBorderColor "white"`
    -TableHeaderBackgroundColor "lightgrey"`
    -DataObject $DomainAdmins
    -TableDescription "Domain Administrators"

Send-MailMessage -SmtpServer smtp.domain.com -to mike@domain.com -from reportserver@domain.com -subject $ReportDescription -body $HTML -bodyashtml
