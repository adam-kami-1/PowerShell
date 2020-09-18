<#
.SYNOPSIS
pomoc displays full help on the selected item.
.DESCRIPTION
If -Test switch is not used then pomoc uses Get-Help -Full to retrieve help information.
.EXAMPLE
pomoc -Name Get_Command
#>
[CmdletBinding(DefaultParameterSetName='AllUsersView', HelpUri='https://go.microsoft.com/fwlink/?LinkID=113316')]
param (

    # Gets help about the specified command or concept. Enter the name of a cmdlet, function, provider, script, or workflow, such as `Get-Member`, a conceptual article name, such as `about_Objects`, or an alias, such as `ls`. Wildcard characters are permitted in cmdlet and provider names, but you can't use wildcard characters to find the names of function help and script help articles.
    #
    # To get help for a script that isn't located in a path that's listed in the `$env:Path` environment variable, type the script's path and file name.
    #
    # If you enter the exact name of a help article, `Get-Help` displays the article contents.
    #
    # If you enter a word or word pattern that appears in several help article titles, `Get-Help` displays a list of the matching titles.
    #
    # If you enter a word that doesn't match any help article titles, `Get-Help` displays a list of articles that include that word in their contents.
    #
    # The names of conceptual articles, such as `about_Objects`, must be entered in English, even in non-English versions of PowerShell.
    [Parameter(Position=0, ValueFromPipelineByPropertyName=$true)]
    [System.String] $Name = 'default',

    # Gets help that explains how the cmdlet works in the specified provider path. Enter a PowerShell provider path.
    #
    # This parameter gets a customized version of a cmdlet help article that explains how the cmdlet works in the specified PowerShell provider path. This parameter is effective only for help about a provider cmdlet and only when the provider includes a custom version of the provider cmdlet help article in its help file. To use this parameter, install the help file for the module that includes the provider.
    #
    # To see the custom cmdlet help for a provider path, go to the provider path location and enter a `Get-Help` command or, from any path location, use the Path parameter of `Get-Help` to specify the provider path. You can also find custom cmdlet help online in the provider help section of the help articles.
    #
    # For more information about PowerShell providers, see about_Providers (./About/about_Providers.md).
    [System.String] ${Path},

    # Displays help only for items in the specified category and their aliases. Conceptual articles are in the HelpFile category.
    [ValidateSet('Alias','Cmdlet','Provider','General','FAQ','Glossary','HelpFile','ScriptCommand','Function','Filter','ExternalScript','All','DefaultHelp','Workflow','DscResource','Class','Configuration')]
    [System.String[]] ${Category},

    # Displays commands with the specified component value, such as Exchange . Enter a component name. Wildcard characters are per mitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [System.String[]] ${Component},

    # Displays help for items with the specified functionality. Enter the functionality. Wildcard characters are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [System.String[]] ${Functionality},

    # Displays help customized for the specified user role. Enter a role. Wildcard characters are permitted.
    #
    # Enter the role that the user plays in an organization. Some cmdlets display different text in their help files based on the value of this parameter. This parameter has no effect on help for the core cmdlets.
    [System.String[]] ${Role},

    # Displays the online version of a help article in the default browser. This parameter is valid only for cmdlet, function, workflow, and script help articles. You can't use the Online parameter with `Get-Help` in a remote session.
    #
    # For information about supporting this feature in help articles that you write, see about_Comment_Based_Help (./About/about_Comment_Based_Help.md), and Supporting Online Help (/powershell/scripting/developer/module/supporting-online-help), and Writing Help for PowerShell Cmdlets (/powershell/scripting/developer/help/writing-help-for-windows-powershell-cmdlets).
    [Parameter(ParameterSetName='Online', Mandatory=$true)]
    [System.Management.Automation.SwitchParameter] ${Online},

    # Displays the help topic in a window for easier reading. The window includes a Find search feature and a Settings box that lets you set options for the display, including options to display only selected sections of a help topic.
    #
    # The ShowWindow parameter supports help topics for commands (cmdlets, functions, CIM commands, workflows, scripts) and conceptual About articles. It does not support provider help.
    #
    # This parameter was introduced in PowerShell 3.0.
    [Parameter(ParameterSetName='ShowWindow', Mandatory=$true)]
    [System.Management.Automation.SwitchParameter] ${ShowWindow},

    # Specifies the number of characters in each line of output. If this parameter is not used, the width is determined by the characteristics of the host. The default for the PowerShell console is 80 characters.
    #
    [Parameter(Mandatory=$false)]
    [System.Int32] $Width,

    # Run my version of help
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.SwitchParameter] $Test,

    # If -Test then rescan help files
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.SwitchParameter] $Rescan
)


#!/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe

#    # Adds parameter descriptions and examples to the basic help display. This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
#    [Parameter(ParameterSetName='DetailedView', Mandatory=$true)]
#    [System.Management.Automation.SwitchParameter] ${Detailed},

#    # Displays the entire help article for a cmdlet. Full includes parameter descriptions and attributes, examples, input and output object types, and additional notes.
#    #
#    # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
#    [Parameter(ParameterSetName='AllUsersView')]
#    [System.Management.Automation.SwitchParameter] ${Full},

#    # Displays only the name, synopsis, and examples. To display only the examples, type `(Get-Help <cmdlet-name>).Examples`.
#    #
#    # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
#    [Parameter(ParameterSetName='Examples', Mandatory=$true)]
#    [System.Management.Automation.SwitchParameter] ${Examples},

#    # Displays only the detailed descriptions of the specified parameters. Wildcards are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
#    [Parameter(ParameterSetName='Parameters', Mandatory=$true)]
#    [System.String] ${Parameter},



function ToCamelCase ( [System.String] $Str )
{
    $ChArr = $Str.ToCharArray()
    $Str = ''
    for ($i = 0; $i -lt $ChArr.Length; $i++)
    {
        if (($i -eq 0) -or ('.-_ '.Contains($ChArr[$i-1])))
        {
            $Str += (($ChArr[$i]).ToString()).ToUpper()
        }
        else
        {
            $Str += $ChArr[$i]
        }
    }
    return $Str
} # ToCamelCase #

function Max ( $Val1, $Val2 )
{
    if ($Val1 -ge $Val2)
    {
        return $Val1
    }
    else
    {
        return $Val2
    }
} # Max #


function Min ( $Val1, $Val2 )
{
    if ($Val1 -le $Val2)
    {
        return $Val1
    }
    else
    {
        return $Val2
    }
} # Min #


Function VerCmp ( [System.String] $Version1, [System.String] $Version2 )
{
    $Ver1 = $Version1.Split('.')
    $Ver2 = $Version2.Split('.')
    for ($i = 0; $i -lt (Min $Ver1.Count $Ver2.Count); $i++)
    {
        if ($Ver1[$i] -ne $Ver2[$i])
        {
            return $Ver1[$i] - $Ver2[$i]
        }
    }
    if ($Ver1.Count -eq $Ver2.Count)
    {
        return 0
    }
    return $Ver1.Count - $Ver2.Count
} # VerCmp #


#####################################################################
# Displaying stuff
#####################################################################

$HelpInfo = @{
    Items = @();
    ItemIndex = @{};
    }

$Work = @{
    # Used only during building $HelpInfo
    Functions = @{};

    # Used only during displaying results
    Colors = @{};
    Output = [System.String[]] @();
    }


function DisplayParagraph ( [System.Int32] $IndentLevel, [System.String] $Format, [System.String] $Text = '')
{
    $Indent = 4 * $IndentLevel
    $TextWidth = $Width-$Indent
    if ($Text -eq '')
    {
        $Format = 'empty'
    }
    $DisplayedLines = 0
    switch ($Format)
    {
        'empty'
            {
                $Work.Output += @('')
                $DisplayedLines++
            }
        {($_ -eq 'para') -or
         ($_ -eq 'hangpara') -or
         ($_ -eq 'comppara') -or
         ($_ -eq 'listpara')}
            {
                $Text = $Text.Trim()
                while ($Text.IndexOf('  ') -ne -1)
                {
                    $Text = $Text.Replace('  ', ' ')
                }
                while ($Text.IndexOf(' .') -ne -1)
                {
                    $Text = $Text.Replace(' .', '.')
                }
                $Lines = @()
                if ($Text.Length -gt 0)
                {
                    while ($Text.Length -gt $TextWidth)
                    {
                        $Line = $Text.Substring(0, $TextWidth)
                        $NextLine = ''
                        if ($Line.IndexOf(' ') -ne -1)
                        {
                            while (($Line.Length -gt 0) -and ($Line.Substring($Line.Length-1, 1) -ne ' '))
                            {
                                $NextLine = $NextLine.Insert(0, $Line.Substring($Line.Length-1, 1))
                                $Line = $Line.Substring(0, $Line.Length-1)
                            }
                        }
                        $Lines += @($Line.TrimEnd())
                        $Text = $NextLine+$Text.Substring($TextWidth)
                        if ((($_ -eq 'hangpara') -or ($_ -eq 'listpara')) -and ($Lines.Count -eq 1))
                        {
                            if ($_ -eq 'listpara')
                            {
                                $TextWidth -= 2
                            }
                            else
                            {
                                $TextWidth -= 4
                            }
                        }
                    }
                    $Lines += @($Text)
                    $no = 0
                    foreach ($Line in $Lines)
                    {
                        $Work.Output += @((' ' * $Indent)+$Line)
                        $DisplayedLines++
                        $no++
                        if ((($_ -eq 'hangpara') -or ($_ -eq 'listpara')) -and ($no -eq 1))
                        {
                            if ($_ -eq 'listpara')
                            {
                                $Indent += 2
                            }
                            else
                            {
                                $Indent += 4
                            }
                        }
                    }
                }
                if (($_ -ne 'comppara') -and ($_ -ne 'listpara'))
                {
                    $Work.Output += @('')
                    $DisplayedLines++
                }
            }
        'code'
            {
                $Lines = $Text.Split("`n")
                foreach ($Line in $Lines)
                {
                    if ($Line.Length -gt $TextWidth)
                    {
                        $Line = $Line.Substring(0, $TextWidth-3)+'...'
                    }
                    $Work.Output += @((' ' * $Indent)+$Line)
                    $DisplayedLines++
                }
            }
        'sect'
            {
                $Work.Output += @((' ' * $Indent)+$Work.Colors.Section+$Text+$Work.Colors.Default)
                $DisplayedLines++
            }
        'subsect'
            {
                $Work.Output += @((' ' * $Indent)+$Work.Colors.ExtraSection+$Text+$Work.Colors.Default)
                $DisplayedLines++
            }
    }
    return $DisplayedLines
} # DisplayParagraph #


function ExtractParagraphText ( [System.Xml.XmlElement] $Para )
{
    $Text = ''
    foreach ($Child in $Para.ChildNodes)
    {
        switch ($Child.GetType())
        {
            'System.Xml.XmlText'
                {
                    $Text += $Child.Value
                }
            'System.Xml.XmlElement'
                {
                    $Text += ExtractParagraphText $Child
                }
        }
    }
    return $Text
} # ExtractParagraphText #


function DisplayCollectionOfParagraphs ( [System.Int32] $IndentLevel,
                                         [System.Object[]] $collection,
                                         [System.Boolean] $WasColon = $false )
{
    if (($collection.Count -eq 0) -or ($collection[0].Length -eq 0))
    {
        return 0
    }
    #$WasColon = $false
    $DisplayedLines = 0
    foreach ($Para in $collection)
    {
        if ($Para.GetType().FullName -eq 'System.Xml.XmlElement')
        {
            # This is a temporary solution. In target version we will need reverse operation:
            # gues the links even wen they are not fully tagged.
            $Para = ExtractParagraphText $Para
        }
        if ($Para.Length -eq 0)
        {
            continue
        }
        if ($Para.IndexOf("`n") -ne -1)
        {
            # Instead of $WasColon there should be used info about colon in last paragraph
            $DisplayedLines += DisplayCollectionOfParagraphs $IndentLevel ($Para.Split("`n")) $WasColon
        }
        else
        {
            $Para = $Para.TrimEnd()
            if (-not $WasColon)
            {
                if ($Para.Substring($Para.Length-1, 1) -eq ':')
                {
                    # List heading paragraph
                    $WasColon = $true
                    $DisplayedLines += DisplayParagraph $IndentLevel comppara $Para
                }
                else
                {
                    # Regular paragraph
                    $DisplayedLines += DisplayParagraph $IndentLevel para $Para
                }
            }
            else
            {
                if ($Para.Substring(0, 2) -eq '- ')
                {
                    # List item
                    $DisplayedLines += DisplayParagraph $IndentLevel listpara $Para
                }
                elseif ($Para.Substring(0, 2) -eq '-- ')
                {
                    # List item
                    $DisplayedLines += DisplayParagraph $IndentLevel listpara ('- '+$Para.Substring(3))
                }
                elseif ($Para.Substring(0, 2) -eq '--')
                {
                    # List item
                    $DisplayedLines += DisplayParagraph $IndentLevel listpara ('- '+$Para.Substring(2))
                }
                else
                {
                    # End of the list
                    $DisplayedLines += DisplayParagraph $IndentLevel empty
                    $WasColon = $false
                    # Regular paragraph
                    $DisplayedLines += DisplayParagraph $IndentLevel para $Para
                }
            }
        }
    }
    if ($WasColon)
    {
        $DisplayedLines += DisplayParagraph $IndentLevel empty
    }
    return $DisplayedLines
} # DisplayCollectionOfParagraphs #


function BuildLinkValue ( [System.String] $LinkText, [System.String] $URI)
{
    if ($URI -ne '')
    {
        if (($URI.Substring(0,8) -ne 'https://') -and ($URI.Substring(0,7) -ne 'http://'))
        {
            Write-Error "Unknown link for ${LinkText}: $URI"
        }
    }
    else
    {
        $version = "{0}.{1}" -f $PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor
        if ($HelpInfo.ItemIndex[$LinkText] -ne $null)
        {
            $URI = $HelpInfo.ItemIndex[$LinkText]
        }
        elseif ($LinkText.IndexOf(' ') -eq -1)
        {
            $URI = "https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/${LinkText}"#?view=powershell-$version"
        }
        else
        {
            $Page = $LinkText.ToLower().Replace(' ', '-')
            $URI = "https://docs.microsoft.com/en-us/powershell/scripting/developer/help/${Page}"#?view=powershell-$version"
        }
    }
    if ($URI -match '^[0-9]+$')
    {
        $LinkValue = '['+$LinkText+']'
    }
    else
    {
        if ($LinkText.Substring($LinkText.Length-1,1) -eq ':')
        {
            $LinkText = $LinkText.Substring(0,$LinkText.Length-1)
        }
        $LinkValue = $LinkText+': '+$URI
    }
    return $LinkValue
} # BuildLinkValue #


function DisplaySingleSyntax ( [System.Xml.XmlElement] $syntaxItem, [System.Boolean] $CommonParameters )
{
    $Para = $syntaxItem.name
    foreach ($parameter in $syntaxItem.parameter)
    {
        $Required = $parameter.required -eq 'true'
        $Position = $parameter.position -ne 'named'
        $Para += ' '
        if (-not $Required)
        {
            $Para += '['
        }
        if ($Position)
        {
            $Para += '['
        }
        $Para += '-'+$parameter.name
        if ($Position)
        {
            $Para += ']'
        }
        $TypeName = $parameter.parameterValue.FirstChild.InnerText
        if (($TypeName -eq '') -or ($TypeName -eq $null))
        {
            $TypeName = $parameter.type.name
        }
        if (($TypeName -ne 'System.Management.Automation.SwitchParameter') -and
            ($TypeName -ne '') -and ($TypeName -ne $null))
        {
            $Para += ' '
            $Para += '<'+$TypeName+'>'
        }
        if (-not $Required)
        {
            $Para += ']'
        }
    }
    if ($CommonParameters)
    {
        $Para += ' [<CommonParameters>]'
    }
    $null = DisplayParagraph 1 hangpara $Para
} # DisplaySingleSyntax #


function DisplaySingleParameter ( [System.Xml.XmlElement] $Parameter )
{
    $Required = $Parameter.required
    $Position = $Parameter.position
    $DefVal = $Parameter.defaultValue
    $PipelineInput = $Parameter.pipelineInput
    $Globbing = $Parameter.globbing
    $Aliases = $Parameter.aliases
    $null = DisplayParagraph 1 comppara ('-'+$Parameter.name+' <'+$Parameter.type.name+'>')

    $null = DisplayCollectionOfParagraphs 2 $Parameter.Description.para

    $null = DisplayParagraph 2 code "Required?                    $Required"
    $null = DisplayParagraph 2 code "Position?                    $Position"
    $null = DisplayParagraph 2 code "Default value                $DefVal"
    $null = DisplayParagraph 2 code "Accept pipeline input?       $PipelineInput"
    $null = DisplayParagraph 2 code "Accept wildcard characters?  $Globbing"
    if (($Aliases -ne '') -and ($Aliases -ne 'none'))
    {
        $null = DisplayParagraph 2 code "Aliases                      $Aliases"
    }
    $null = DisplayParagraph 0 empty
} # DisplaySingleParameter #


function DisplaySingleExample ( [System.Xml.XmlElement] $Example )
{
    $null = DisplayParagraph 1 para $Example.title
    $null = DisplayParagraph 2 code $Example.code
    $null = DisplayParagraph 2 empty
    $null = DisplayCollectionOfParagraphs 2 $Example.remarks.para
} # DisplaySingleExample #


function DisplayXmlHelpFile ( [System.Collections.Hashtable] $Item, [System.Xml.XmlElement] $command )
{
    $null = DisplayParagraph 0 empty

    # ========================================
    # Section NAME
    $null = DisplayParagraph 0 sect "NAME"
    $null = DisplayParagraph 1 para $command.details.name

    # ========================================
    # Section SYNOPSIS
    $null = DisplayParagraph 0 sect "SYNOPSIS"
    $null = DisplayParagraph 1 para $command.details.description.para

    # ========================================
    # Section SYNTAX
    if (($command.syntax -ne $null) -and
        ($command.syntax.syntaxItem -ne $null))
    {
        $null = DisplayParagraph 0 sect "SYNTAX"

        if ($command.syntax.syntaxItem.Count -ne 0)
        {
            foreach ($syntaxItem in $command.syntax.syntaxItem)
            {
                DisplaySingleSyntax $syntaxItem $Item.CommonParameters
            }
        }
        else
        {
            DisplaySingleSyntax $command.syntax.syntaxItem $Item.CommonParameters
        }
    }

    # ========================================
    # Section DESCRIPTION
    if ($command.description.para.Count -gt 0)
    {
        $null = DisplayParagraph 0 sect "DESCRIPTION"
        $null = DisplayCollectionOfParagraphs 1 $command.description.para
    }
    if ($command.description.section -ne $null)
    {
        if ($command.description.section.Count -gt 1)
        {
            foreach ($Section in $command.description.section)
            {
                $null = DisplayParagraph 0 'subsect' $Section.name
                $null = DisplayCollectionOfParagraphs 1 $Section.para
            }
        }
        else
        {
            $null = DisplayParagraph 0 'subsect' $command.description.section.name
            $null = DisplayCollectionOfParagraphs 1 $command.description.section.para
        }
    }
    
    # ========================================
    # Section PARAMETERS
    $HeaderDisplayed = $false
    if (($command.parameters -ne $null) -and
        ($command.parameters.parameter -ne $null))
    {
        $null = DisplayParagraph 0 sect "PARAMETERS"
        $HeaderDisplayed = $true
        if ($command.parameters.parameter.Count -ne 0)
        {
            foreach ($Parameter in $command.parameters.parameter)
            {
                DisplaySingleParameter $Parameter
            }
        }
        else
        {
            DisplaySingleParameter $command.parameters.parameter
        }
    }
    if ($Item.CommonParameters)
    {
        if (-not $HeaderDisplayed)
        {
            $null = DisplayParagraph 0 sect "PARAMETERS"
        }
        $null = DisplayParagraph 1 comppara ('<CommonParameters>')

        $null = DisplayParagraph 2 para ('This cmdlet supports the common parameters: '+
            'Verbose, Debug, ErrorAction, ErrorVariable, WarningAction, '+
            'WarningVariable, OutBuffer, PipelineVariable, and OutVariable. '+
            'For more information, see about_CommonParameters '+
            '(https:/go.microsoft.com/fwlink/?LinkID=113216).')
    }

    # ========================================
    # Section INPUTS
    if (($command.InputTypes -ne $null) -and
        ($command.InputTypes.InputType -ne $null))
    {
        $DisplayedLines = 0
        if ($command.InputTypes.InputType.Count -ne $null)
        {
            $null = DisplayParagraph 0 sect "INPUTS"
            foreach ($InputType in $command.InputTypes.InputType)
            {
                $null = DisplayParagraph 1 comppara $InputType.type.name
                $DisplayedLines += DisplayCollectionOfParagraphs 2 $InputType.description.para
            }
        }
        else
        {
            if ((([String]$command.InputTypes.InputType.type.name) -ne '') -or
                (([String]$command.InputTypes.InputType.description.para) -ne ''))
            {
                $null = DisplayParagraph 0 sect "INPUTS"
                $null = DisplayParagraph 1 comppara $command.InputTypes.InputType.type.name
                $DisplayedLines += DisplayCollectionOfParagraphs 2 $command.InputTypes.InputType.description.para
            }
        }
        if ($DisplayedLines -eq 0)
        {
            $null = DisplayParagraph 2 empty
        }
    }

    # ========================================
    # Section OUTPUTS
    if (($command.returnValues -ne $null) -and
        ($command.returnValues.returnValue -ne $null))
    {
        $DisplayedLines = 0
        if ($command.returnValues.returnValue.Count -ne $null)
        {
            $null = DisplayParagraph 0 sect "OUTPUTS"
            foreach ($returnValue in $command.returnValues.returnValue)
            {
                $null = DisplayParagraph 1 comppara $returnValue.type.name
                $DisplayedLines = DisplayCollectionOfParagraphs 2 $returnValue.description.para
            }
        }
        else
        {
            if ((([String]$command.returnValues.returnValue.type.name) -ne '') -or
                (([String]$command.returnValues.returnValue.description.para) -ne ''))
            {
                $null = DisplayParagraph 0 sect "OUTPUTS"
                $null = DisplayParagraph 1 comppara $command.returnValues.returnValue.type.name
                $DisplayedLines = DisplayCollectionOfParagraphs 2 $command.returnValues.returnValue.description.para
            }
        }
        if ($DisplayedLines -eq 0)
        {
            $null = DisplayParagraph 2 empty
        }
    }
    
    # ========================================
    # Section NOTES
    if (($command.alertSet -ne $null) -and
        ($command.alertSet.alert -ne $null) -and
        ($command.alertSet.alert.para.Count -ne 0) -and
        ($command.alertSet.alert.para[0] -ne $null))
    {
        $null = DisplayParagraph 0 sect "NOTES"
        $null = DisplayCollectionOfParagraphs 1 $command.alertSet.alert.para
    }

    # ========================================
    # Section EXAMPLES
    if (($command.examples -ne $null) -and
        ($command.examples.example -ne $null))
    {
        $null = DisplayParagraph 0 sect "EXAMPLES"
        if ($command.examples.example.Count -ne 0)
        {
            foreach ($Example in $command.examples.example)
            {
                DisplaySingleExample $Example
            }
        }
        else
        {
            DisplaySingleExample $command.examples.example
        }
    }

    # ========================================
    # Section RELATED LINKS
    if (($command.relatedLinks -ne $null) -and
        ($command.relatedLinks.navigationLink -ne $null))
    {
        $null = DisplayParagraph 0 sect "RELATED LINKS"
        $Links = @()
        if (($command.relatedLinks.navigationLink -is [System.Object[]]) -and
            ($command.relatedLinks.navigationLink.Count -gt 0))
        {
            foreach ($NavigationLink in $command.relatedLinks.navigationLink)
            {
                $LinkValue = BuildLinkValue $NavigationLink.linkText $NavigationLink.uri
                $Links += @('* '+$LinkValue)
            }
        }
        else
        {
            $LinkValue = BuildLinkValue $command.relatedLinks.navigationLink.linkText $command.relatedLinks.navigationLink.uri
            $Links += @('* '+$LinkValue)
        }
        $null = DisplayCollectionOfParagraphs 1 $Links
    }

    # ========================================
    # Section REMARKS
    if ($false)
    {
        $null = DisplayParagraph 0 sect "REMARKS"
        $null = DisplayParagraph 1 para 'Remarks will be described later !!!'
    }
} # DisplayXmlHelpFile #


function CleanParagraph ( [System.String] $Paragraph )
{
    # Replace all new lines with spaces
    $Paragraph = $Paragraph.Replace("`n", ' ')

    # Compress every spaces string into single space
    while ($Paragraph.IndexOf('  ') -ne -1)
    {
        $Paragraph = $Paragraph.Replace('  ', ' ')
    }
    return $Paragraph.Trim()
} # CleanParagraph #

function AddLinesToNewChild ( [System.Xml.XmlDocument] $XML, 
                              [System.Xml.XmlElement] $Parent, 
                              [System.String] $ChildName,
                              [System.Int32] $SkipLines,
                              [System.String] $Paragraph)
{
    # Ignore first $SkipLines lines from the $Paragraph
    while (($SkipLines -gt 0) -and ($Paragraph.IndexOf("`n") -gt 0))
    {
        $Paragraph = $Paragraph.Substring($Paragraph.IndexOf("`n")+1)
        $SkipLines--
    }
    if ($Clean)
    {
        $Paragraph = CleanParagraph $Paragraph
    }
    else
    {
        $Paragraph = $Paragraph.Replace("`n", ' ')
    }
    $Child = $Parent.AppendChild($XML.CreateElement($ChildName))
    $Child.Set_innerText($Paragraph)
} # AddLinesToNewChild #


function AddNavigationLink ( [System.Xml.XmlDocument] $XML, 
                             [System.Xml.XmlElement] $Parent,
                             [System.String] $Line)
{
    $Line = $Line.Trim()
    if ($Line.Length -lt 1)
    {
        return
    }
    if ($Line.Substring(0,1) -eq '-')
    {
        $Line = $Line.Substring(1).Trim()
    }
    if ($Line.Length -lt 1)
    {
        return
    }
    # Line could be in a form:
    # Link Text (http://...)
    # or
    # "Link Text" (http://...)
    # or
    # Link Text (https://...)
    # or
    # "Link Text" (https://...)
    # or
    # Link Text
    if ($Line -like '*(http*')
    {
        $Text = ($Line -replace '^ *(.*) *\((https?://.*)\)$', '$1').Trim('" ')
        $Url = ($Line -replace '^ *(.*) *\((https?://.*)\)$', '$2').Trim()
    }
    elseif ($Line -like '*http*')
    {
        $Text = ($Line -replace '^ *(.*) *(https?://.*)$', '$1').Trim('" ')
        $Url = ($Line -replace '^ *(.*) *(https?://.*)$', '$2').Trim()
    }
    else
    {
        $Text = $Line
        $Url = ''
    }
    if (($Url -ne '') -and ($Text -eq '') -and ($Parent.LastChild -ne $null) -and ($Parent.LastChild.uri -eq ''))
    {
        # LinkText and uri are splitted into two lines, so add uri
        # to the previous navigationLink
        $Parent.LastChild.ChildNodes[1].Set_innerText($Url)
    }
    else
    {
        # Create navigationLink link elements with two children: linkText and uri.
        # Set values of both children to $Text and $Url
        $NavigationLink = $Parent.AppendChild($XML.CreateElement('navigationLink'))
        $LinkText = $NavigationLink.AppendChild($XML.CreateElement('LinkText'))
        $LinkText.Set_innerText($Text)
        $Uri = $NavigationLink.AppendChild($XML.CreateElement('uri'))
        $Uri.Set_innerText($Url)
    }
} # AddNavigationLink #


function AddLinesToRelatedLinks ( [System.Xml.XmlDocument] $XML,
                                  [System.Int32] $SkipLines,
                                  [System.String] $Paragraph )
{
    # Ignore first $SkipLines lines from the $Paragraph
    while (($SkipLines -gt 0) -and ($Paragraph.IndexOf("`n") -gt 0))
    {
        $Paragraph = $Paragraph.Substring($Paragraph.IndexOf("`n")+1)
        $SkipLines--
    }
    $RelatedLinks = (Select-XML -Xml $XML -XPath '/helpItems/command/relatedLinks').Node
    $Lines = $Paragraph.Split("`n")
    foreach ($Line in $Lines)
    {
        AddNavigationLink $XML $RelatedLinks $Line
    }
} # AddLinesToRelatedLinks #


function AddDescriptionParagraph ( [System.Xml.XmlDocument] $XML,
                                   [System.String] $Paragraph )
{
    $Description = (Select-XML -Xml $XML -XPath '/helpItems/command/description').Node
    if ($Description -eq $null)
    {
        # There was no LONG DESCRIPTION header in file
        $Description = $Command.AppendChild($XML.CreateElement('description'))
    }
    AddLinesToNewChild $XML $Description 'para' 0 $Paragraph
}


function ExtraSectionHeading ( [System.String] $Paragraph )
{
    if (($Paragraph.IndexOf("`n") -ne -1) -or
        ($Paragraph.Trim().Length -eq 0)) 
    {
        return ''
    }
    $Paragraph = $Paragraph.TrimEnd()
    if ($Paragraph.Substring(0,1) -in @(' ','-'))
    {
        return ''
    }
    $Paragraph = $Paragraph.TrimStart()
    if ($Paragraph -in @('or', 'and'))
    {
        return ''
    }
    if ($Paragraph.Substring($Paragraph.Length-1,1) -in @('.', ':'))
    {
        return ''
    }
    if (-not ($Paragraph -match '^[-:, a-z0-9]+$'))
    {
        return ''
    }
    return $Paragraph
} # ExtraSectionHeading #


function StoreRegularparagraph ( [System.Collections.Hashtable] $Item,
                                 [System.Xml.XmlDocument] $XML,
                                 [System.String] $Paragraph )
{
    if ($Item.CurrentSectionName -eq 'NOTES')
    {
        $Alert = (Select-XML -Xml $XML -XPath '/helpItems/command/alertSet/alert').Node
        AddLinesToNewChild $XML $Alert 'para' 0 $Paragraph
    }
    elseif ($Item.CurrentExtraSectionNode -ne $null)
    {
        AddLinesToNewChild $XML $Item.CurrentExtraSectionNode 'para' 0 $Paragraph
    }
    else
    {
        AddDescriptionParagraph $XML $Paragraph
    }
} # StoreRegularparagraph #


function ParseRegularParagraph ( [System.Collections.Hashtable] $Item,
                                 [System.Xml.XmlDocument] $XML,
                                 [System.String] $Paragraph )
{
    #Write-Host $Paragraph
    switch ($Item.CurrentSectionName)
    {
        {($_ -eq 'NAME') -or
         ($_ -eq 'SYNOPSIS')}
            {
                $Details = (Select-XML -Xml $XML -XPath '/helpItems/command/details').Node
                if ($Details -eq $null)
                {
                    # There was no TOPIC nor item name. Put name into XML.
                    $Details = $Command.AppendChild($XML.CreateElement('details'))
                    $Name = $Details.AppendChild($XML.CreateElement('name'))
                    $Name.Set_innerText($Item.DisplayName)
                }
                if ($XML.helpItems.command.details.description -eq $null)
                {
                    # The synopsis was separated with empty line from SHORT DESCRIPTION.
                    $Description = $Details.AppendChild($XML.CreateElement('description'))
                    $Para = $Description.AppendChild($XML.CreateElement('para'))
                    $Paragraph = CleanParagraph $Paragraph
                    $Para.Set_innerText($Paragraph)
                }
                else
                {
                    AddDescriptionParagraph $XML $Paragraph
                    $Item.CurrentSectionName = 'DESCRIPTION'
                }
            }
        'RELATED LINKS'
            {
                $RelatedLinks = (Select-XML -Xml $XML -XPath '/helpItems/command/relatedLinks').Node
                if ($Paragraph.IndexOf("`n") -eq -1)
                {
                    AddNavigationLink $XML $RelatedLinks $Paragraph
                }
                else
                {
                    AddLinesToRelatedLinks $XML 0 $Paragraph
                }
            }
        {($_ -eq 'DESCRIPTION') -or
         ($_ -eq 'NOTES')}
            {
                $ExtraSection = ExtraSectionHeading $Paragraph
                if ($Item.Name -eq 'about_Comment_Based_Help')
                {
                    if (($ExtraSection -eq 'Name') -or
                        ($ExtraSection -eq 'Syntax') -or
                        ($ExtraSection -eq 'Description') -or
                        ($ExtraSection -eq 'Parameters') -or
                        ($ExtraSection -eq 'Parameter List') -or
                        ($ExtraSection -eq 'Common Parameters') -or
                        ($ExtraSection -eq 'Parameter Attribute Table') -or
                        ($ExtraSection -eq 'Inputs') -or
                        ($ExtraSection -eq 'Outputs') -or
                        ($ExtraSection -eq 'Remarks') -or
                        ($ExtraSection -eq 'Related Links'))
                    {
                        $ExtraSection = ''
                    }
                }
                switch -Regex ($ExtraSection)
                {
                    '^$'
                        {
                            # Really regular paragraph, add to appropriate section
                            StoreRegularparagraph $Item $XML $Paragraph
                        }
                    '^(NOTES|REMARKS)$'
                        {
                            $Command = (Select-XML -Xml $XML -XPath '/helpItems/command').Node
                            $AlertSet = $Command.AppendChild($XML.CreateElement('alertSet'))
                            $Alert = $AlertSet.AppendChild($XML.CreateElement('alert'))
                            $Item.CurrentSectionName = 'NOTES'
                            $Item.CurrentExtraSectionNode = $null
                        }
                    default
                        {
                            $Item.CurrentExtraSectionName = $ExtraSection.ToUpper()
                            #Write-Host "ExtraSection: " $Item.CurrentExtraSectionName
                            $Description = (Select-XML -Xml $XML -XPath '/helpItems/command/description').Node
                            $Item.CurrentExtraSectionNode = $Description.AppendChild($XML.CreateElement('section'))
                            $Name = $Item.CurrentExtraSectionNode.AppendChild($XML.CreateElement('name'))
                            $Name.Set_innerText($Item.CurrentExtraSectionName)
                            $Item.CurrentSectionName = 'DESCRIPTION'
                        }
                }
            }
    }
} # ParseRegularParagraph #


function ParseTxtHelpFile ( [System.Collections.Hashtable] $Item )
{
    # Convert contents of the HelpFile into array of paragraphs
    $File = Get-Content $Item.File
    $Paragraphs = @()
    $Paragraph = ''
    foreach ($Line in $File)
    {
        if ($Line.Trim() -ne '')
        {
            if ($Paragraph -eq '')
            {
                $Paragraph = $Line
            }
            else
            {
                $Paragraph += "`n"+$Line
            }
        }
        elseif ($Paragraph -ne '')
        {
            $Paragraphs += $Paragraph
            $Paragraph = ''
        }
    }
    if ($Paragraph -ne '')
    {
        $Paragraphs += $Paragraph
    }
    Clear-Variable File
    if (($Item.Name -eq 'about_PowerShell_exe'))
    {
        #Write-Host "Breakpoint"
    }

    # Convert file name into help item name
    $Item.DisplayName = $Item.Name.Replace('_', ' ')
    $Item.DisplayName = ToCamelCase $Item.DisplayName
    

    # parse paragraphs converting them into MAML like XML object
    ############################################################
    $XML = [System.Xml.XmlDocument]'<?xml version="1.0" encoding="utf-8"?><helpItems />'
    $Command = $XML.ChildNodes[1].AppendChild($XML.CreateElement('command'))
    $Item.CurrentSectionName = ''
    foreach ($Paragraph in $Paragraphs)
    {
        if ($Paragraph.IndexOf("`n") -ne -1)
        {
            $FirstLine = $Paragraph.Substring(0,$Paragraph.IndexOf("`n")).Trim()
        }
        else
        {
            $FirstLine = $Paragraph.Trim()
        }
        switch -Regex ($FirstLine)
        {
            ('^(TOPIC|'+$Item.DisplayName+')$')
                {
                    if ($Item.CurrentSectionName -eq '')
                    {
                        # Found TOPIC or item name. Put it into XML ignoring value in file.
                        $Details = $Command.AppendChild($XML.CreateElement('details'))
                        $Name = $Details.AppendChild($XML.CreateElement('name'))
                        $Name.Set_innerText($Item.DisplayName)
                        $Item.CurrentSectionName = 'NAME'
                    }
                    else
                    {
                        # Found item name in a firt line of a paragraph.
                        # Treat this a regular paragraph, probably extra section name.
                        ParseRegularParagraph $Item $XML $Paragraph
                    }
                }
            # In about_Type_Accelerators.help.txt there is typo: 'SHORT DESRIPTION'
            '^(SHORT DES(C)?RIPTION|SYNOPSIS)$'
                {
                    if ($Item.CurrentSectionName -eq 'DESCRIPTION')
                    {
                        ParseRegularParagraph $Item $XML $Paragraph
                        break
                    }
                    if ($XML.helpItems.command.details -eq $null)
                    {
                        # There was no TOPIC nor item name. Put name into XML.
                        $Details = $Command.AppendChild($XML.CreateElement('details'))
                        $Name = $Details.AppendChild($XML.CreateElement('name'))
                        $Name.Set_innerText($Item.DisplayName)
                    }
                    if ($Paragraph.IndexOf("`n") -ne -1)
                    {
                        # SHORT DESCRIPTION and synopsis are in one paragraph, but in two lines.
                        # Extract synopsis and put it into XML.
                        $Description = $Details.AppendChild($XML.CreateElement('description'))
                        # !!! Need trim and remove duplicate spaces
                        AddLinesToNewChild $XML $Description 'para' 1 $Paragraph
                    }
                    $Item.CurrentSectionName = 'SYNOPSIS'
                }
            '^((LONG|DETAILED) )?DESCRIPTION$'
                {
                    if ($Item.CurrentSectionName -eq 'DESCRIPTION')
                    {
                        $Description = (Select-XML -Xml $XML -XPath '/helpItems/command/description').Node
                        if ($Description.para.Count -eq 1)
                        {
                            # There was only one extra paragraph before section header,
                            # So we can move it to Synopsis
                            $Text = $XML.helpItems.command.details.description.para + ' ' +
                                    $XML.helpItems.command.description.para
                            $Synopsis = (Select-XML -Xml $XML -XPath '/helpItems/command/details/description/para').Node
                            $Synopsis.Set_innerText($Text)
                            $Para = (Select-XML -Xml $XML -XPath '/helpItems/command/description/para').Node
                            $Description.RemoveChild($Para) | Out-Null
                        }
                        else
                        {
                            ParseRegularParagraph $Item $XML $Paragraph
                            break
                        }
                    }
                    else
                    {
                        $Description = $Command.AppendChild($XML.CreateElement('description'))
                    }
                    if ($Paragraph.IndexOf("`n") -ne -1)
                    {
                        AddLinesToNewChild $XML $Description 'para' 1 $Paragraph
                    }
                    $Item.CurrentSectionName = 'DESCRIPTION'
                    $Item.CurrentExtraSectionNode = $null
                }
            '^SEE ALSO$'
                {
                    $RelatedLinks = $Command.AppendChild($XML.CreateElement('relatedLinks'))
                    if ($Item.OnlineURI -ne '')
                    {
                        AddNavigationLink $XML $RelatedLinks ('Online Version: '+$Item.OnlineURI)
                    }
                    if ($Paragraph.IndexOf("`n") -gt 0)
                    {
                        AddLinesToRelatedLinks $XML 1 $Paragraph
                    }
                    $Item.CurrentSectionName = 'RELATED LINKS'
                }
            default
                {
                    ParseRegularParagraph $Item $XML $Paragraph
                }
        }
    }
    return $XML
} # ParseTxtHelpFile #


function DisplayHelpItem ( [System.Collections.Hashtable] $Item )
{
    switch ($Item.Format)
    {
        "txt"
            {
                $XML = ParseTxtHelpFile $Item
                show-XML.ps1 $XML -Tree ascii -Width ([System.Console]::WindowWidth-5)| Out-String | Write-Verbose
                DisplayXmlHelpFile $Item $XML.ChildNodes[1].ChildNodes[0]
            }
        "xml"
            {
                # Write-Host ("Displaying file: "+$Item.File+" Item no: "+$Item.Index)
                $XML = [System.Xml.XmlDocument](Get-Content $Item.File)
                DisplayXmlHelpFile $Item ($XML.helpItems.command)[$Item.Index]
            }
    }
} # DisplayHelpItem #


#####################################################################
# File searching stuff
#####################################################################


function AddItem ( [System.Boolean] $MarkFunc, [System.Collections.Hashtable] $Item )
{
    # $Item = [pscustomobject]$I
    # Write-Host ("Adding Item "+$Item.Name)
    if ($HelpInfo.ItemIndex[$Item.Name] -eq $null)
    {
        if ($Work.Functions[$Item.Name] -ne $null)
        {
            $Item.Category = 'Function'
            if ($MarkFunc)
            {
                $Work.Functions[$Item.Name] = $null
            }
        }
        $HelpInfo.ItemIndex[$Item.Name] = $HelpInfo.Items.Count
        $HelpInfo.Items += $Item
    }
} # AddItem #


$TxtHelpFileModule = @{
    'about_ActivityCommonParameters'       = 'PSWorkflow';
    'about_Certificate_Provider'           = 'Microsoft.PowerShell.Security';
    'about_Checkpoint-Workflow'            = 'PSWorkflow';
    'about_Classes_and_DSC'                = 'PSDesiredStateConfiguration';
    'about_Escape_Characters'              = 'drop it !!!'; # there is about_Special_Characters
    'about_ForEach-Parallel'               = 'PSWorkflow';
    'about_InlineScript'                   = 'PSWorkflow';
    'about_PSReadline'                     = 'PSReadLine';
    'about_Parallel'                       = 'PSWorkflow';
    'about_Parsing_LocTest'                = 'drop it !!!'; # there is about_Parsing
    'about_PowerShell.exe'                 = 'drop it !!!'; # there is newer about_PowerShell_exe
    'about_PowerShell_Ise.exe'             = 'drop it !!!'; # there is newer about_PowerShell_Ise_exe
    'about_Scheduled_Jobs'                 = 'PSScheduledJob';
    'about_Scheduled_Jobs_Advanced'        = 'PSScheduledJob';
    'about_Scheduled_Jobs_Basics'          = 'PSScheduledJob';
    'about_Scheduled_Jobs_Troubleshooting' = 'PSScheduledJob';
    'about_Sequence'                       = 'PSWorkflow';
    'about_Suspend-Workflow'               = 'PSWorkflow';
    'about_WS-Management_Cmdlets'          = 'Microsoft.WSMan.Management';
    'about_WSMan_Provider'                 = 'Microsoft.WSMan.Management';
    'about_Windows_PowerShell_5.0'         = 'drop it !!!'; # This is old
    'about_WorkflowCommonParameters'       = 'PSWorkflow';
    'about_Workflows'                      = 'PSWorkflow';
    'default'                              = '';
    'WSManAbout'                           = 'drop it !!!'; # This is not a help file
    }


function GetModuleAndOnlineURI ( [System.String] $Name, [System.String] $ModuleName )
{
    $URI = ''
    if ($ModuleName -eq '')
    {
        $ModuleName = $TxtHelpFileModule[$Name]
        if ($ModuleName -ne 'drop it !!!')
        {
            if ($ModuleName -eq '')
            {
                $ModuleName = 'Microsoft.PowerShell.Core'
            }
            $version = "{0}.{1}" -f $PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor

            #$a0 = $ModuleName.Value
            #$a1 = $a0.ToLower()
            #$a2 = ($ModuleName.Value).ToLower()
            #$a3 = $Name.ToLower()

            if ($Name -ne 'default')
            {
                $URI = 'https://docs.microsoft.com/en-us/powershell/module/'+$ModuleName.ToLower()+
                       '/about/'+$Name.ToLower()+'?view=powershell-'+$version
            }
            return @{'Module'=$ModuleName;
                     'URI'=$URI}
        }
    }
    return $URI = @{'Module'=$ModuleName;
                    'URI'=$URI}
} # GetModuleAndOnlineURI #


function CheckTxtHelpFiles ( [System.String] $ModuleName, [System.String] $Path )
{
    if ( -not (Test-Path -Path $Path -PathType Container))
    {
        return
    }
    $Files = (Get-ChildItem $Path\*.help.txt).Name
    if ($Files.Count -gt 0)
    {
        foreach ($File in $Files)
        {
            $Name = $File -replace '.help.txt',''
            $URI = GetModuleAndOnlineURI $Name $ModuleName
            if ($URI.Module -eq 'drop it !!!')
            {
                continue
            }
            $Item = @{Name = $Name;
                      ModuleName = $URI.Module;
                      File = "$Path\$File";
                      OnlineURI = $URI.URI;
                      Format = 'txt';
                      Index = -1;
                      Category = 'HelpFile';
                      Synopsis = '';
                      CommonParameters = $false}
            $XML = ParseTxtHelpFile $Item
            $Item.Synopsis = $XML.helpItems.command.details.description.para
            AddItem $true $Item
        }
    }
    return
} # CheckTxtHelpFiles #


function CheckXMLFile ( [System.Collections.Hashtable] $LocalFuncs, [System.String] $ModuleName, [System.String] $Path, [System.String] $File )
{
    $XML = [System.Xml.XmlDocument](Get-Content "$Path\$File")
    if ($XML.helpItems -ne $null)
    {
        if ($XML.helpItems.command -ne $null)
        {
            if (($XML.helpItems.command).GetType().Name -eq "Object[]")
            {
                for ($Index = 0; $Index -lt ($XML.helpItems.command).Count; $Index++)
                {
                    $command = ($XML.helpItems.command)[$Index]
                    $Category = 'Cmdlet'
                    if ($LocalFuncs[$command.details.name] -ne $null)
                    {
                        $Category = 'Function'
                        $LocalFuncs[$command.details.name] = $null
                    }
                    $CmdInfo = Get-Command -Name $command.details.name -ErrorAction SilentlyContinue
                    if ($CmdInfo -ne $null)
                    {
                        $CommonParameters = $CmdInfo.Definition.IndexOf('<CommonParameters>') -ne -1
                    }
                    else
                    {
                        $CommonParameters = $false
                    }
                    AddItem $true @{Name = $command.details.name;
                                    ModuleName = $ModuleName;
                                    File = "$Path\$File";
                                    OnlineURI = '';
                                    Format = 'xml';
                                    Index = $Index;
                                    Category = $Category;
                                    Component = '';
                                    Functionality = '';
                                    Role = '';
                                    Synopsis = $command.details.description.para;
                                    CommonParameters = $CommonParameters}
                }
            }
        }
    }
    foreach ($Function in $LocalFuncs.keys)
    {
        if ($LocalFuncs[$Function] -ne $null)
        {
            $CmdInfo = Get-Command -Name $Function -ErrorAction SilentlyContinue
            if ($CmdInfo -ne $null)
            {
                $CommonParameters = $CmdInfo.Definition.IndexOf('<CommonParameters>') -ne -1
            }
            else
            {
                $CommonParameters = $false
            }
            AddItem $false @{Name = $Function;
                             ModuleName = $ModuleName;
                             File = '';
                             OnlineURI = '';
                             Format = '';
                             Index = -1;
                             Category = 'Function';
                             Component = '';
                             Functionality = '';
                             Role = '';
                             Synopsis = '...';
                             CommonParameters = $CommonParameters}
        }
    }
} # CheckXMLFile #


function CheckXmlHelpFiles ( [System.String[]] $LocalFunctions, [System.String] $ModuleName, [System.String] $Path, [System.String] $Pattern )
{
    if ( -not (Test-Path -Path $Path -PathType Container))
    {
        return
    }
    $LocalFuncs = @{}
    if ($LocalFunctions -ne $null)
    {
        for ($i = 0; $i -lt $LocalFunctions.Count; $i++)
        {
            $LocalFuncs[$LocalFunctions[$i]] = 'Function'
        }
    }
    $Files = (Get-ChildItem $Path\*$Pattern).Name
    if ($Files.Count -gt 0)
    {
        foreach ($File in $Files)
        {
            if ($ModuleName -eq '')
            {
                $MN = $File.Remove($File.Length - $Pattern.Length)
                switch ($MN)
                {
                    'System.Management.Automation'
                        {
                            $MN = 'Microsoft.PowerShell.Core'
                        }
                    'Microsoft.PowerShell.Consolehost'
                        {
                            $MN = 'Microsoft.PowerShell.Host'
                        }
                    'Microsoft.PowerShell.Commands.Diagnostics'
                        {
                            $MN = 'Microsoft.PowerShell.Diagnostics'
                        }
                    'Microsoft.PowerShell.Commands.Management'
                        {
                            $MN = 'Microsoft.PowerShell.Management'
                        }
                    'Microsoft.PowerShell.Commands.Utility'
                        {
                            $MN = 'Microsoft.PowerShell.Utility'
                        }
                }
                CheckXMLFile $LocalFuncs $MN $Path $File
            }
            else
            {
                CheckXMLFile $LocalFuncs $ModuleName $Path $File
            }
        }
    }
} # CheckXmlHelpFiles #


function CheckModule ( [System.String] $Path, [System.String] $ModuleName, [System.String] $Version)
{
    if ($Version -eq '')
    {
        $Path = "$Path\$ModuleName"
    }
    else
    {
        $Path = "$Path\$ModuleName\$Version"
    }
    $LocalFuncs = @()
    if (Test-Path -Path "$Path\$ModuleName.psd1")
    {
        $LocalFuncs = (Import-PowerShellDataFile -Path "$Path\$ModuleName.psd1" -ErrorAction SilentlyContinue).FunctionsToExport
    }
    CheckTxtHelpFiles $ModuleName "$Path\$PSUICulture"
    #CheckXmlHelpFiles $LocalFuncs $ModuleName "$Path\$PSUICulture" '.dll-help.xml'
    #CheckXmlHelpFiles $LocalFuncs $ModuleName "$Path\$PSUICulture" '.cdxml-help.xml'
    #CheckXmlHelpFiles $LocalFuncs $ModuleName "$Path\$PSUICulture" '.psd1-help.xml'
    #CheckXmlHelpFiles $LocalFuncs $ModuleName "$Path\$PSUICulture" '.psm1-help.xml'
    CheckXmlHelpFiles $LocalFuncs $ModuleName "$Path\$PSUICulture" '-help.xml'
} # CheckModule #


function FindHelpFiles ()
{
    Get-ChildItem function: | ForEach-Object { $Work.Functions[$_.Name] = 'Function' }
    CheckTxtHelpFiles '' $PSHOME\$PSUICulture
    $LocalFuncs = @()
    CheckXmlHelpFiles $LocalFuncs '' $PSHOME\$PSUICulture '.dll-help.xml'
    CheckXmlHelpFiles $LocalFuncs '' $PSHOME\$PSUICulture '-help.xml'
    foreach ($ModulePath in (($env:PSModulePath).Split(';')))
    {
        foreach ($Module in ((Get-Childitem $ModulePath -Directory).Name))
        {
            $Version = ''
            foreach ($SubDir in ((Get-Childitem $ModulePath\$Module -Directory).Name))
            {
                if ($SubDir -eq $PSUICulture)
                {
                    CheckModule $ModulePath $Module ''
                }
                elseif ($SubDir -match '^([0-9]\.)+[0-9]+$')
                {
                    if ((VerCmp $Version $SubDir) -le 0)
                    {
                        $Version = $SubDir
                    }
                }
            }
            if ($Version -ne '')
            {
                #Write-Host "Checking module: $Module\$SubDir"
                CheckModule $ModulePath $Module $Version
            }
        }
    }
    Get-ChildItem alias: | 
        ForEach-Object `
        {
            # Alias  $_.Name  ->  $_.Definition
            #
            if ($HelpInfo.ItemIndex[$_.Definition] -ne $null)
            {
                $Item = $HelpInfo.Items[$HelpInfo.ItemIndex[$_.Definition]]
                AddItem $true @{Name = $_.Name;
                                ModuleName = $Item.ModuleName;
                                File = $Item.File;
                                OnlineURI = $Item.OnlineURI;
                                Format = $Item.Format;
                                Index = $Item.Index;
                                Category = 'Alias';
                                Component = $Item.Component;
                                Functionality = $Item.Functionality;
                                Role = $Item.Role;
                                Synopsis = $_.Definition;
                                CommonParameters = $_.CommonParameters}
            }
        }
    foreach ($Function in $Work.Functions.keys)
    {
        if ($Work.Functions[$Function] -ne $null)
        {
            $CmdInfo = Get-Command -Name $Function -ErrorAction SilentlyContinue
            if ($CmdInfo -ne $null)
            {
                $CommonParameters = $CmdInfo.Definition.IndexOf('<CommonParameters>') -ne -1
            }
            else
            {
                $CommonParameters = $false
            }
            AddItem $false @{Name = $Function;
                             ModuleName = '';
                             File = '';
                             OnlineURI = '';
                             Format = '';
                             Index = -1;
                             Category = 'Function';
                             Component = '';
                             Functionality = '';
                             Role = '';
                             Synopsis = '...';
                             CommonParameters = $CommonParameters}
        }
    }
    $Work.Functions = @{}
} # FindHelpFiles #


#####################################################################
# Main body
#####################################################################

# Below two lines are in function help ()
#Set the outputencoding to Console::OutputEncoding. More.com doesn't work well with Unicode.
#$outputEncoding=[System.Console]::OutputEncoding

$Test = $true
#$Rescan = $true
#$Name = 'Add-Computer'
#$Name = 'Get-Help'
#$Name = '*'
#$Name = 'about_do'


###########################################################
# Call standard Get-Help if no -Test
if (-not $Test)
{
    Get-Help @PSBoundParameters | C:\Git\usr\bin\less.exe
    return
}


###########################################################
# Find all *.help.txt and  *.dll-help.xml files HelpFiles
if ((Test-Path -Path $env:USERPROFILE\.PS-pomoc.xml) -and -not $Rescan)
{
    Write-Verbose ('Importing stored info about help files to '+$env:USERPROFILE+'\.PS-pomoc.xml')
    $HelpInfo = Import-Clixml -Path $env:USERPROFILE\.PS-pomoc.xml
}
else
{
    Write-Verbose "Find all *.help.txt and  *.dll-help.xml files HelpFiles"
    FindHelpFiles
    Write-Verbose ('Storing info about help files from '+$env:USERPROFILE+'\.PS-pomoc.xml')
    Export-Clixml -Path $env:USERPROFILE\.PS-pomoc.xml -Encoding UTF8 -InputObject $HelpInfo
}


###########################################################

if ($Width -eq 0)
{
    $Width = [System.Console]::WindowWidth
}
if ($psISE -eq $null)
{
    $Esc=[char]0x1B;
    $F_Default = "${Esc}[39m";
    $F_Black = "${Esc}[30m";
    $F_DarkRed = "${Esc}[31m";
    $F_DarkGreen = "${Esc}[32m";
    $F_DarkYellow = "${Esc}[33m";
    $F_DarkBlue = "${Esc}[34m";
    $F_DarkMagenta = "${Esc}[35m";
    $F_DarkCyan = "${Esc}[36m";
    $F_Gray = "${Esc}[37m";
    $F_DarkGray = "${Esc}[90m";
    $F_Red = "${Esc}[91m";
    $F_Green = "${Esc}[92m";
    $F_Yellow = "${Esc}[93m";
    $F_Blue = "${Esc}[94m";
    $F_Magenta = "${Esc}[95m";
    $F_Cyan = "${Esc}[96m";
    $F_White = "${Esc}[97m";

    $Work.Colors = @{
        Section = $F_Magenta;
        ExtraSection = $F_Cyan;
        Parameter = $F_Yellow;
        Default = $F_Default;
        }
}
else
{
    $Work.Colors = @{
        Section = '';
        ExtraSection = '';
        Parameter = '';
        Default = '';
        }
}
$found = @()
for ($i = 0; $i -lt $HelpInfo.Items.Count; $i++)
{
    if (($HelpInfo.Items[$i].Name -like $Name) -and
        (($Category.Count -eq 0) -or ($Category -contains $HelpInfo.Items[$i].Category)) -and
        (($Component.Count -eq 0) -or ($Component -contains $HelpInfo.Items[$i].Component)) -and
        (($Functionality.Count -eq 0) -or ($Functionality -contains $HelpInfo.Items[$i].Functionality)) -and
        (($Role.Count -eq 0) -or ($Role -contains $HelpInfo.Items[$i].Role)))
    {
        $found += $i
    }
}
switch ($found.Count)
{
    0
        {
            Write-Error "Unable to find $Name"
        }
    1
        {
            DisplayHelpItem $HelpInfo.Items[$found[0]]
            $Work.Output
        }
    default
        {
            for ($i = 0; $i -lt $found.Count; $i++)
            {
                [pscustomobject]@{Name = $HelpInfo.Items[$found[$i]].Name;
                                  Category = $HelpInfo.Items[$found[$i]].Category;
                                  Module = $HelpInfo.Items[$found[$i]].ModuleName;
                                  Synopsis = $HelpInfo.Items[$found[$i]].Synopsis;
                                  #File = $HelpInfo.Items[$found[$i]].File;
                                  #Index = $HelpInfo.Items[$found[$i]].Index
                                 }
            }
        }
}
