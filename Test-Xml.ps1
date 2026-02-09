<#
.SYNOPSIS
Verify XML file using W3C XML Schema.
.DESCRIPTION
The contents of the XML file is verified using 'xmllint.exe'
from https://gitlab.gnome.org/GNOME/libxml2.
.EXAMPLE
Verifying PS help XML file:
test-xml .\Connect-PSSession.xml .\helpItems.xsd
.INPUTS
None
.OUTPUTS
None
#>Param (
    [Parameter(Mandatory=$true,
               Position=0)]
    [System.String] $XMLPath,
    [Parameter(Mandatory=$true,
               Position=1)]
    [System.String] $SchemaPath
)
D:\Users\Adam\bin\APPS\XML\xmllint.exe --noout --schema $SchemaPath $XMLPath 2>&1 | ./Show-Item
