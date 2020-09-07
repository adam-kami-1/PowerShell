<#
.SYNOPSIS
   Display the contents of XML file.
.DESCRIPTION
   The `Show-Xml.ps1` script displays the contents of XML file. The contents is presented as a list of indented nodes.
.EXAMPLE
Show-Xml -InpuObject .\example.xml
.EXAMPLE
$XML = [xml](Get-Content .\Get-Help.xml)
Show-Xml -InputObject $XML

    This shows whole XML file.
.EXAMPLE
$XML = [xml](Get-Content .\Get-Help.xml)
Show-Xml -InputObject $XML.ChildNodes[1].ChildNodes[0].ChildNodes[0]
Show-Xml -InputObject $XML.helpItems.command[0].details

    Both calls to Show-Xml shows the the same piece of XML file.
.INPUTS
   None
.OUTPUTS
   System.String[]
   `Show-Xml` returns collection of strings.
.NOTES
   Errors in XML file could cause corruption of `Show-Xml` script.
#>
[OutputType([String[]])]
param (

    # Specifies how many levels of the XML tree will be displayed. 0 means that only input element will be displayed. The default value is -1, meaning all objects.
    #
    [Parameter(Mandatory=$false)]
    [ValidateRange(-1,500)]
    [System.Int32] $Depth = -1,

    # If parameter is a System.String then it specifies the path to the XML file.
    # If parameter is one of types:
    # - System.Xml.XmlDocument,
    # - System.Xml.XmlDeclaration,
    # - System.Xml.XmlElement,
    # then it specifies the XML object.
    #
    [Parameter(Mandatory=$true,
               Position=0,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [System.Management.Automation.PSObject] $InputObject = $null,

    # Specifies the number of characters in each line of output. If this parameter is not used, the width is determined by the characteristics of the host. The default for the PowerShell console is 80 characters.
    #
    [Parameter(Mandatory=$false)]
    [ValidateRange(50,500)]
    [System.Int32] $Width = 0
)

if ($Width -eq 0)
{
    if ( $psISE -ne $null )
    {
        # This is PowerShell ISE
        $Width = 80
    }
    else
    {
        $Width = [System.Console]::BufferWidth
    }
}
$MaxWidth = $Width-1



function DisplayXMLItem ( [System.Int32] $IndentLevel, $Item)
{
    if (($Depth -ne -1) -and ($IndentLevel -gt $Depth))
    {
        return
    }
    switch ($Item.GetType().FullName)
    {
        {($_ -eq 'System.Xml.XmlElement') -or
         ($_ -eq 'System.Xml.XmlDeclaration') -or
         ($_ -eq 'System.Xml.XmlDocument')}
            {
                ###############
                # Name and type
                Write-Output ('    '*$IndentLevel+$Item.LocalName+' <'+$Item.GetType().Name+'>')

                ############
                # Attributes
                if ($Item.GetType().FullName -eq 'System.Xml.XmlDeclaration')
                {
                    $Attributes = @()
                    if ($Item.OuterXml.IndexOf(' version=') -ne -1)
                    {
                        $Attributes += @{Name='version'; Value=$Item.Version}
                    }
                    if ($Item.OuterXml.IndexOf(' encoding=') -ne -1)
                    {
                        $Attributes += @{Name='encoding'; Value=$Item.Encoding}
                    }
                }
                else
                {
                    $Attributes = $Item.Attributes
                }
                foreach ($Attribute in $Attributes)
                {
                    $Width = $MaxWidth-('    '*($IndentLevel+2)).Length
                    $Value = $Attribute.Name+'="'+$Attribute.Value+'"'
                    if ($Value.Length -gt ($Width-1))
                    {
                        $Value = $Value.Substring(0,$Width-3)+'..."'
                    }
                    Write-Output ('    '*($IndentLevel+2)+$Value)
                }

                ##########
                # Children
                if ($Item.HasChildNodes)
                {
                    for ($i = 0; $i -lt $Item.ChildNodes.Count; $i++)
                    {
                        DisplayXMLItem ($IndentLevel+1) $Item.ChildNodes[$i]
                    }
                }
            }
        'System.Xml.XmlText'
            {
                $Width = $MaxWidth-('    '*$IndentLevel).Length
                $Pos = $Item.Value.IndexOf("`n")
                if ($Pos -ne -1)
                {
                    $Value = '"'+$Item.Value.Substring(0,$Item.Value.IndexOf("`n"))+'..."'
                }
                else
                {
                    $Value = '"'+$Item.Value+'"'
                }
                if ($Value.Length -gt $Width)
                {
                    $Value = $Value.Substring(0,$Width-3)+'..."'
                }
                Write-Output ('    '*$IndentLevel+$Value)
            }
        default
            {
                Write-Output "`n`nUnsupported type $_`n`n"
            }
    }
} # DisplayXMLItem #



Write-Output ("`n")
switch ($InputObject.GetType().FullName)
{
    'System.String'
        {
            if (Test-Path $InputObject)
            {
                $XML = [System.Xml.XmlDocument](Get-Content $InputObject)
                DisplayXMLItem 0 $XML
            }
            else
            {
                $XML = [System.Xml.XmlDocument]$InputObject
                DisplayXMLItem 0 $XML
            }
        }
    {($_ -eq 'System.Xml.XmlDocument') -or
     ($_ -eq 'System.Xml.XmlDeclaration') -or
     ($_ -eq 'System.Xml.XmlElement')}
        {
            DisplayXMLItem 0 $InputObject
        }
    default
        {
            Write-Error "Parameter -InputObject is not an XML document or element"
        }
}

