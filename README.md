.SYNOPSIS
    Accepts an object, converts it to a table and generates an HTML report. The returned object can be used as the body of an email or saved as a document.

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
  
.EXAMPLE
  $HTML = Generate-EmailReport -ReportName "Domain Admins"`
    -ReportDescription "List of users in the 'Administrators' AD group."`
    -EmailFooterText "This report ran from $env:computername as a scheduled task as the user $env:username at $(Get-Date -Format hh:MM:ss ).<br>Please contact administrator@domain.com with any questions."`
    -HTMLBodyBackgroundColor "white"`
    -TableBorderColor "white"`
    -TableHeaderBackgroundColor "lightgrey"`
    -DataObject $DomainAdmins
    -TableDescription "Domain Administrators"
