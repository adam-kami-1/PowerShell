<#
.SYNOPSIS
Verify XML file using W3C XML Schema.
.DESCRIPTION
The contents of the XML file is verified using 'xmllint.exe'
from https://gitlab.gnome.org/GNOME/libxml2.
.PARAMETER XMLPath
Specifies XML file to validate.
.PARAMETER SchemaPath
Specifies file containing W3C XML Schema.
.PARAMETER DtdPath
Specifies file containing Document Type Definition.
.EXAMPLE
Verifying PS help XML file using W3C XML Schema:
test-xml .\Connect-PSSession.xml -SchemaPath .\helpItems.xsd
.INPUTS
None
.OUTPUTS
None
#>
Param (
    [Parameter(Mandatory=$true,
               Position=0)]
    [System.String] $XMLPath,

    [Parameter(Mandatory=$false)]
    [System.String] $SchemaPath = '',

    [Parameter(Mandatory=$false)]
    [System.String] $DtdPath = ''
)
if ($SchemaPath -ne '')
{
    if ($DtdPath -ne '')
    {
        D:\Users\Adam\bin\APPS\XML\xmllint.exe --noout --valid --dtdvalid $DtdPath --schema $SchemaPath $XMLPath 2>&1 | Show-Item.ps1
    }
    else
    {
        D:\Users\Adam\bin\APPS\XML\xmllint.exe --noout --valid --schema $SchemaPath $XMLPath 2>&1 | Show-Item.ps1
    }
}
else
{
    if ($DtdPath -ne '')
    {
        D:\Users\Adam\bin\APPS\XML\xmllint.exe --noout --valid --dtdvalid $DtdPath $XMLPath 2>&1 | Show-Item.ps1
    }
    else
    {
        D:\Users\Adam\bin\APPS\XML\xmllint.exe --noout --valid $XMLPath 2>&1 | Show-Item.ps1
    }
}
