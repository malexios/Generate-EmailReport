
<#
.SYNOPSIS
    Accepts an object and generates am HTML report. The data in the object is converted to a table in the report. The returned object can be used as the body of an email or saved as a document.

.DESCRIPTION
    There is a bug in ConvertTo-HTML that returns an asterisk for the table header when the data object only has one property.

.PARAMETER ReportName
    Mandatory
    string
    Used as the main title of the report. Appears at the top of the returned HTML string with the current date appended.

.PARAMETER ReportDescription
    Mandatory
    string
    Short description of the report. It will appear under ReportName.

.PARAMETER EmailFooterText
    string
    Text to appear below the table. Runas user, server, time, etc. 
    You may not want to expose the server name and user. On the other hand, it may be useful eight years from now if the report is still running and no one knows where it's coming from.
    ex:  
    "This report ran from $env:computername as a scheduled task as the user $env:username at $(Get-Date -Format hh:MM:ss ).<br>Please contact $SupportEmailAddress with any questions."

.PARAMETER HTMLBodyBackgroundColor
    string
    Default: white
    The background color of the returned HTML. Must be a valid HTML color.

.PARAMETER TableBorderColor
    string
    Default: black
    The border color for the table in the returned HTML. Must be a valid HTML color.

.PARAMETER TableHeaderBackgroundColor
    string
    Default: white
    The background color for the table header in the returned HTML. Must be a valid HTML color.

.PARAMETER TableCellBackgroundColor
    string
    Default: white
    The background color for the cells in the returned HTML. Must be a valid HTML color.

.PARAMETER DataObject
    Used to generate the table

.PARAMETER TableDescription
    Appears just above the table in the returned HTML.

.INPUTS
  none

.OUTPUTS
  none

.NOTES
    Version:        1.0
    Author:         Michael Alexios
    Creation Date:  2/8/19
  
.EXAMPLE
  $HTML = Generate-EmailReport -ReportName "Domain Admins"`
    -ReportDescription "List of users in the 'Administrators' AD group."`
    -EmailFooterText "This report ran from $env:computername as a scheduled task as the user $env:username at $(Get-Date -Format hh:MM:ss ).<br>Please contact administrator@domain.com with any questions."`
    -HTMLBodyBackgroundColor "white"`
    -TableBorderColor "white"`
    -TableHeaderBackgroundColor "lightgrey"`
    -DataObject $DomainAdmins
    -TableDescription "Domain Administrators"

.TODO
    Work around bug in ConvertTo-HTML that returns an asterisk for the table header when the data object only has one property.
    Workaround: $serverObjects | ConvertTo-HTML -Property 'Server Name' -Head $style
#>

Param(
    [Parameter(Mandatory)]
    $ReportName,
    [Parameter(Mandatory)]
    $ReportDescription,
    [Parameter()]
    $EmailFooterText,
    [Parameter()]
    $HTMLBodyBackgroundColor = "white",
    [Parameter()]
    $TableBorderColor = "black",
    [Parameter()]
    $TableHeaderBackgroundColor = "white",
    [Parameter()]
    $TableCellBackgroundColor = "white",
    [Parameter()]
    $DataObject,
    [Parameter()]
    $TableDescription

)

#----------------------------------------------------------[Declarations]----------------------------------------------------------
$ValidHTMLColors = @("AliceBlue","AntiqueWhite","Aqua","Aquamarine","Azure","Beige","Bisque","Black","BlanchedAlmond","Blue","BlueViolet","Brown","BurlyWood","CadetBlue","Chartreuse","Chocolate","Coral","CornflowerBlue","Cornsilk","Crimson","Cyan","DarkBlue","DarkCyan","DarkGoldenRod","DarkGray","DarkGrey","DarkGreen","DarkKhaki","DarkMagenta","DarkOliveGreen","DarkOrange","DarkOrchid","DarkRed","DarkSalmon","DarkSeaGreen","DarkSlateBlue","DarkSlateGray","DarkSlateGrey","DarkTurquoise","DarkViolet","DeepPink","DeepSkyBlue","DimGray","DimGrey","DodgerBlue","FireBrick","FloralWhite","ForestGreen","Fuchsia","Gainsboro","GhostWhite","Gold","GoldenRod","Gray","Grey","Green","GreenYellow","HoneyDew","HotPink","IndianRed","","Indigo","","Ivory","Khaki","Lavender","LavenderBlush","LawnGreen","LemonChiffon","LightBlue","LightCoral","LightCyan","LightGoldenRodYellow","LightGray","LightGrey","LightGreen","LightPink","LightSalmon","LightSeaGreen","LightSkyBlue","LightSlateGray","LightSlateGrey","LightSteelBlue","LightYellow","Lime","LimeGreen","Linen","Magenta","Maroon","MediumAquaMarine","MediumBlue","MediumOrchid","MediumPurple","MediumSeaGreen","MediumSlateBlue","MediumSpringGreen","MediumTurquoise","MediumVioletRed","MidnightBlue","MintCream","MistyRose","Moccasin","NavajoWhite","Navy","OldLace","Olive","OliveDrab","Orange","OrangeRed","Orchid","PaleGoldenRod","PaleGreen","PaleTurquoise","PaleVioletRed","PapayaWhip","PeachPuff","Peru","Pink","Plum","PowderBlue","Purple","RebeccaPurple","Red","RosyBrown","RoyalBlue","SaddleBrown","Salmon","SandyBrown","SeaGreen","SeaShell","Sienna","Silver","SkyBlue","SlateBlue","SlateGray","SlateGrey","Snow","SpringGreen","SteelBlue","Tan","Teal","Thistle","Tomato","Turquoise","Violet","Wheat","White","WhiteSmoke","Yellow","YellowGreen")

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Begin-HTML {
    Param(
        [Parameter(Mandatory)]
        $ReportName,
        [Parameter(Mandatory)]
        $ReportDescription,
        [Parameter()]
        $HTMLBodyBackgroundColor,
        [Parameter()]
        $TableBorderColor,
        [Parameter()]
        $TableHeaderBackgroundColor,
        [Parameter()]
        $TableCellBackgroundColor
    )

    $HTML = "<!DOCTYPE html>"
    $HTML += "<style>"
    $HTML += "BODY{background-color:$HTMLBodyBackgroundColor;}"
    $HTML += "TABLE{border-width: 1px;border-style: solid;border-color: $TableBorderColor;border-collapse: collapse;}"
    $HTML += "TH{border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:$TableHeaderBackgroundColor}"
    $HTML += "TD{border-width: 1px;padding: 4px;border-style: solid;border-color: black;background-color:$TableCellBackgroundColor}"
    $HTML += "</style>"
    $HTML += "<body>"
    $HTML += "<H2>$ReportName - $(get-date -format d)</H2>"
    $HTML += "<p>$ReportDescription</p>"
    return $HTML
}

function Add-TableToHtml {
    Param(
        [Parameter(Mandatory)]
        $Data,
        [Parameter(Mandatory)]
        $TableDescription
    )

    $HTML += "<br/><br/>"
    $HTML += "<b><h3>$TableDescription</h3>"
    $HTML += $Data | ConvertTo-Html -Fragment

    # There is a bug in ConvertTo-HTML that returns an asterisk for the table header when the data object only has one property.
    # This will not work correctly if there is more them one property in the object
    # ($ErrorServers | Get-Member -MemberType NoteProperty).count #A note in case I want to make sure there is only one property in the object
    $PropertyName = (($Data | Get-Member)[-1]).Name
    $HTML = $HTML -replace '<th>\*</th>',"<th>$PropertyName</th>"


    return $HTML
}

function Add-TableToHtml ($TableData,$TableDescription){
    $HTML += "<h3>$TableDescription</h3>"
    if ($TableData.count -gt 0){
        $HTML += $TableData | ConvertTo-Html -Fragment
        # There is a bug in ConvertTo-HTML that returns an asterisk for the table header when the data object only has one property.
        # ($ErrorServers | Get-Member -MemberType NoteProperty).count #in case I want to make sure there is only one property in the object
        $PropertyName = (($TableData | Get-Member)[-1]).Name
        $HTML = $HTML -replace '<th>\*</th>',"<th>$PropertyName</th>"
    }
    else {$HTML += "<br>No problems found."}

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
if ($ValidHTMLColors -notcontains $HTMLBodyBackgroundColor) {throw "$($HTMLBodyBackgroundColor) is not a valid HTML color"}
if ($ValidHTMLColors -notcontains $TableBorderColor) {throw "$($TableBorderColor) is not a valid HTML color"}
if ($ValidHTMLColors -notcontains $TableHeaderBackgroundColor) {throw "$($TableHeaderBackgroundColor) is not a valid HTML color"}
if ($ValidHTMLColors -notcontains $TableCellBackgroundColor) {throw "$($TableCellBackgroundColor) is not a valid HTML color"}



$EmailBody = Begin-HTML `
-ReportName $ReportName `
-ReportDescription $ReportDescription `
-HTMLBodyBackgroundColor $HTMLBodyBackgroundColor `
-TableBorderColor $TableBorderColor `
-TableHeaderBackgroundColor $TableHeaderBackgroundColor `
-TableCellBackgroundColor $TableCellBackgroundColor

$EmailBody += Add-TableToHtml -Data $DataObject -TableDescription $TableDescription
$EmailBody += Close-HTML -EmailFooterText $EmailFooterText


return $EmailBody
