
<#
.DESCRIPTION
    Template to format data as HTML and include it in a emailed report.

.NOTES
    Version:        1.0
    Author:         Michael Alexios
    Creation Date:  2/8/19
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------

##### Report Variables #####
$ReportName = "Very Important Stuff That You Need To Know"
$EmailHeaderText = "Short blurb about the table/tables below.<br>It's going to be HTML, so feel free to include HTML tags."
$EmailFooterText = "This report ran from xxxx server as a scheduled task and took xxx seconds to complete"

##### Email variables #####
$SMTPServer = "smtp.domain.com"
$from = "REPORTSERVER01@domain.com"
$to = "to@domain.com"

$HTML = $null

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Build-EmailBody ($ReportName,$EmailHeaderText,$EmailFooterText,$DataObject1,$DataObject2){
    $HTMLHeader = Begin-HTML -ReportName $ReportName -EmailHeaderText $EmailHeaderText
    $HTMLTable1 = Add-TableToHtml -Data $DataObject1 -TableDescription "Here is a table of useful information."
    $HTMLTable2 = Add-TableToHtml -Data $DataObject2 -TableDescription "Here is more useful information."
    $HTMLClose = Close-HTML
    $EmailBody = $HTMLHeader + $HTMLTable1 + $HTMLTable2 + $HTMLClose
    return $EmailBody
}

function Begin-HTML {
    Param(
        [Parameter(Mandatory)]
        $ReportName,
        [Parameter(Mandatory)]
        $EmailHeaderText
    )

    $HTML = "<!DOCTYPE html>"
    $HTML += "<style>"
    $HTML += "BODY{background-color:white;}"
    $HTML += "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
    $HTML += "TH{border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:PowderBlue}"
    $HTML += "TD{border-width: 1px;padding: 4px;border-style: solid;border-color: black;background-color:LightGrey}"
    $HTML += "</style>"
    $HTML += "<body>"
    $HTML += "<H2>$ReportName - $(get-date -format d)</H2>"
    $HTML += "<p>$EmailHeaderText</p>"
    return $HTML
}

function Add-TableToHtml {
    Param(
        [Parameter(Mandatory)]
        $Data,
        [Parameter(Mandatory)]
        $TableDescription
    )

    $HTML += "<b>$TableDescription"
    $HTML += $Data | ConvertTo-Html -Fragment
    return $HTML
}

function Close-HTML {
    Param(
        [Parameter()]
        $EmailFooterText
    )
    $html += $EmailFooterText
    $html += "</body>"
    $html += "</html>"
    return $html
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

$DataObject1 = Collect-SomeData
$DataObject2 = Collect-MoreData

$EmailBody = Build-EmailBody $ReportName $EmailHeaderText $EmailFooterText $DataObject1 $DataObject2

Send-MailMessage -smtpserver $smtpserver -from $from -to $to -subject $ReportName -body $EmailBody -bodyashtml
