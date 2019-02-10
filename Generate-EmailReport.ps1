
<#
.DESCRIPTION
    Template to format data as HTML and include it in an emailed report.
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------

##### Report Variables #####
$ReportName = "Very Important Stuff That You Need To Know"
$EmailHeaderText = "Short blurb about the table/tables below.<br>It's going to be HTML, so feel free to include HTML tags."
$AuthorEmailAddress = "author@domain.com"
# You may not want to expose the server name and user. On the other hand, it may be useful eight years from now if the report is still running and no one knows where it's coming from.
$EmailFooterText = "This report ran from $env:computername as a scheduled task as the user $env:username at $(Get-Date -Format hh:MM:ss ).<br>Please contact $AuthorEmailAddress with any questions."

##### Report and Table Colors #####
# Use a supported HTML color name
# Some common colors: Red,Cyan,Blue,DarkBlue,LightBlue,Purple,Yellow,Lime,Magenta,White,Silver,Grey,Black,Orange,Brown,Maroon,Green,Olive
$BodyBackgroundColor = "White"
$TableBorderColor = "Black"
$TableHeaderBackgroundColor =  "PowderBlue"
$TableCellBackgroundColor =  "LightGrey"

##### Email Variables #####
$SMTPServer = "smtp.domain.com"
$from = "REPORTSERVER01@domain.com"
$to = "recipient@domain.com"

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Build-EmailBody{
    Param(
        [Parameter(Mandatory=$true)]
        $ReportName,
        [Parameter(Mandatory=$true)]
        $EmailHeaderText,
        [Parameter()]
        $EmailFooterText,
        [Parameter()]
        $BodyBackgroundColor = "white",
        [Parameter()]
        $TableBorderColor = "black",
        [Parameter()]
        $TableHeaderBackgroundColor = "white",
        [Parameter()]
        $TableCellBackgroundColor = "white",
        [Parameter()]
        $DataObject1,
        [Parameter()]
        $DataObject2
    )

    $EmailBody = Begin-HTML `
        -ReportName $ReportName `
        -EmailHeaderText $EmailHeaderText `
        -BodyBackgroundColor $BodyBackgroundColor `
        -TableBorderColor $TableBorderColor `
        -TableHeaderBackgroundColor $TableHeaderBackgroundColor `
        -TableCellBackgroundColor $TableCellBackgroundColor

    $EmailBody += Add-TableToHtml -Data $DataObject1 -TableDescription "Here is a table of useful information."
    $EmailBody += Add-TableToHtml -Data $DataObject2 -TableDescription "Here is more useful information."
    $EmailBody += Close-HTML -EmailFooterText $EmailFooterText
    return $EmailBody
}

function Begin-HTML {
    Param(
        [Parameter(Mandatory=$true)]
        $ReportName,
        [Parameter(Mandatory=$true)]
        $EmailHeaderText,
        [Parameter()]
        $BodyBackgroundColor,
        [Parameter()]
        $TableBorderColor,
        [Parameter()]
        $TableHeaderBackgroundColor,
        [Parameter()]
        $TableCellBackgroundColor
    )

    $HTML = "<!DOCTYPE html>"
    $HTML += "<style>"
    $HTML += "BODY{background-color:$BodyBackgroundColor;}"
    $HTML += "TABLE{border-width: 1px;border-style: solid;border-color: $TableBorderColor;border-collapse: collapse;}"
    $HTML += "TH{border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:$TableHeaderBackgroundColor}"
    $HTML += "TD{border-width: 1px;padding: 4px;border-style: solid;border-color: black;background-color:$TableCellBackgroundColor}"
    $HTML += "</style>"
    $HTML += "<body>"
    $HTML += "<H2>$ReportName - $(get-date -format d)</H2>"
    $HTML += "<p>$EmailHeaderText</p>"
    return $HTML
}

function Add-TableToHtml {
    Param(
        [Parameter(Mandatory=$true)]
        $Data,
        [Parameter(Mandatory=$true)]
        $TableDescription
    )
    $HTML += "<br/><br/>"
    $HTML += "<b><h3>$TableDescription</h3>"
    $HTML += $Data | ConvertTo-Html -Fragment
    return $HTML
}

function Close-HTML {
    Param(
        [Parameter()]
        $EmailFooterText
    )

    $HTML += "<br/><br/>"
    $html += $EmailFooterText
    $html += "</body>"
    $html += "</html>"
    return $html
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

$DataObject1 = Get-ChildItem "c:\program files"
$DataObject2 = Get-ChildItem "C:\Program Files (x86)"

$EmailBody = Build-EmailBody `
    -ReportName $ReportName `
    -EmailHeaderText $EmailHeaderText `
    -EmailFooterText $EmailFooterText `
    -BodyBackgroundColor $BodyBackgroundColor `
    -TableBorderColor $TableBorderColor `
    -TableHeaderBackgroundColor $TableHeaderBackgroundColor `
    -TableCellBackgroundColor $TableCellBackgroundColor `
    -DataObject1 $DataObject1 `
    -DataObject2 $DataObject2

Send-MailMessage -smtpserver $smtpserver -from $from -to $to -subject $ReportName -body $EmailBody -bodyashtml
