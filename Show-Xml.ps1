﻿<#
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
    [ValidateRange(-1,[System.Int32]::MaxValue)]
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

    # Select format of XML tree elements connectors
    #
    [Parameter(Mandatory=$false)]
    [ValidateSet('none', 'ascii', 'box')]
    [System.String] $Tree = 'none',
    
    # Specifies the number of characters in each line of output. If this parameter is not used, the width is determined by the characteristics of the host..
    #
    [Parameter(Mandatory=$false)]
    [ValidateRange(50,[System.Int32]::MaxValue)]
    [System.Int32] $Width = 0
)

if ($Width -eq 0)
{
    $Width = [System.Console]::WindowWidth
}
$MaxWidth = $Width-1
if ($Depth -eq -1)
{
    $Depth = [System.Int32]::MaxValue
}
switch ($Tree)
{
    'none'
        {
            $TreeChild = '    '
            $TreeLast  = '    '
            $TreeCont  = '    '
        }
    'ascii'
        {
            $TreeChild = '  +-'
            $TreeLast  = '  +-'
            $TreeCont  = '  | '
        }
    'box'
        {
            $TreeChild = '  ├─'
            $TreeLast =  '  └─'
            $TreeCont =  '  │ '
        }
}


function DisplayXMLItem ( [System.String] $Indent, $Item)
{
    $IndentLevel = $Indent.Length / 4
    if ($IndentLevel -gt $Depth)
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
                Write-Output ($Indent+$Item.LocalName+' <'+$Item.GetType().Name+'>')
                if (($Indent.Length -gt 0))
                {
                    if ($Item -eq $Item.ParentNode.LastChild)
                    {
                        $Indent = $Indent.Substring(0,$Indent.Length-4)+'    '
                    }
                    else
                    {
                        $Indent = $Indent.Substring(0,$Indent.Length-4)+$TreeCont
                    }
                }

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
                if (($Item.FirstChild -eq $null) -or (($IndentLevel+1) -gt $Depth))
                {
                    $AttrIndent = '    '
                }
                else
                {
                    $AttrIndent = $TreeCont
                }
                foreach ($Attribute in $Attributes)
                {
                    $Width = $MaxWidth-($Indent+$AttrIndent+'    ').Length
                    $Value = $Attribute.Name+'="'+$Attribute.Value+'"'
                    if ($Value.Length -gt ($Width-1))
                    {
                        $Value = $Value.Substring(0,$Width-3)+'..."'
                    }
                    Write-Output ($Indent+$AttrIndent+'    '+$Value)
                }

                ##########
                # Children
                if ($Item.HasChildNodes)
                {
                    for ($i = 0; $i -lt $Item.ChildNodes.Count; $i++)
                    {
                        if ($i -eq ($Item.ChildNodes.Count-1))
                        {
                            DisplayXMLItem ($Indent+$TreeLast) $Item.ChildNodes[$i]
                        }
                        else
                        {
                            DisplayXMLItem ($Indent+$TreeChild) $Item.ChildNodes[$i]
                        }
                    }
                }
            }
        'System.Xml.XmlText'
            {
                $Width = $MaxWidth-$Indent.Length
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
                Write-Output ($Indent+$Value)
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
                DisplayXMLItem '' $XML
            }
            else
            {
                $XML = [System.Xml.XmlDocument]$InputObject
                DisplayXMLItem '' $XML
            }
        }
    {($_ -eq 'System.Xml.XmlDocument') -or
     ($_ -eq 'System.Xml.XmlDeclaration') -or
     ($_ -eq 'System.Xml.XmlElement')}
        {
            DisplayXMLItem '' $InputObject
        }
    default
        {
            Write-Error "Parameter -InputObject is not an XML document or element"
        }
}
