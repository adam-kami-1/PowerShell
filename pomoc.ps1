﻿<#
.SYNOPSIS
pomoc displays full help on the selected item.
.DESCRIPTION
This is planed as a replacement of Get-Help cmdlet.
.EXAMPLE
pomoc -Name Get_Command
#>

[CmdletBinding(DefaultParameterSetName='AllUsersView')]
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
    [System.String] $Path,

    # Displays help only for items in the specified category and their aliases. Conceptual articles are in the HelpFile category.
    [ValidateSet('Alias','Cmdlet','Provider','General','FAQ','Glossary','HelpFile','ScriptCommand','Function','Filter','ExternalScript','All','DefaultHelp','Workflow','DscResource','Class','Configuration')]
    [System.String[]] $Category,

    # Displays commands with the specified component value, such as Exchange . Enter a component name. Wildcard characters are per mitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [System.String[]] $Component,

    # Displays help for items with the specified functionality. Enter the functionality. Wildcard characters are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [System.String[]] $Functionality,

    # Displays help customized for the specified user role. Enter a role. Wildcard characters are permitted.
    #
    # Enter the role that the user plays in an organization. Some cmdlets display different text in their help files based on the value of this parameter. This parameter has no effect on help for the core cmdlets.
    [System.String[]] $Role,

    # Adds parameter descriptions and examples to the basic help display. This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
    #
    # In fact in script pomoc.ps1 this parameter is ignored. Always full info is displayed.
    [Parameter(ParameterSetName='DetailedView', Mandatory=$true)]
    [System.Management.Automation.SwitchParameter] ${Detailed},

    # Displays the entire help article for a cmdlet. Full includes parameter descriptions and attributes, examples, input and output object types, and additional notes.
    #
    # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
    #
    # In fact in script pomoc.ps1 this parameter is ignored. Always full info is displayed.
    [Parameter(ParameterSetName='AllUsersView')]
    [System.Management.Automation.SwitchParameter] $Full,

    # Displays only the name, synopsis, and examples. To display only the examples, type `(Get-Help <cmdlet-name>).Examples`.
    #
    # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
    #
    # In fact in script pomoc.ps1 this parameter is ignored. Always full info is displayed.
    [Parameter(ParameterSetName='Examples', Mandatory=$true)]
    [System.Management.Automation.SwitchParameter] $Examples,

    # Displays only the detailed descriptions of the specified parameters. Wildcards are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [Parameter(ParameterSetName='Parameters', Mandatory=$true)]
    [System.String] $Parameter,

    # Displays the online version of a help article in the default browser. This parameter is valid only for cmdlet, function, workflow, and script help articles. You can't use the Online parameter with `Get-Help` in a remote session.
    #
    # For information about supporting this feature in help articles that you write, see about_Comment_Based_Help (./About/about_Comment_Based_Help.md), and Supporting Online Help (/powershell/scripting/developer/module/supporting-online-help), and Writing Help for PowerShell Cmdlets (/powershell/scripting/developer/help/writing-help-for-windows-powershell-cmdlets).
    [Parameter(ParameterSetName='Online', Mandatory=$true)]
    [System.Management.Automation.SwitchParameter] $Online,

    # Displays the help topic in a window for easier reading. The window includes a Find search feature and a Settings box that lets you set options for the display, including options to display only selected sections of a help topic.
    #
    # The ShowWindow parameter supports help topics for commands (cmdlets, functions, CIM commands, workflows, scripts) and conceptual About articles. It does not support provider help.
    #
    # This parameter was introduced in PowerShell 3.0.
    [Parameter(ParameterSetName='ShowWindow', Mandatory=$true)]
    [System.Management.Automation.SwitchParameter] $ShowWindow,

    # Specifies the number of characters in each line of output. If this parameter is not used, the width is determined by the characteristics of the host. The default for the PowerShell console is 80 characters.
    [Parameter(Mandatory=$false)]
    [System.Int32] $Width,

    # Rescan help files. This is required when there are added new or newer help files.
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.SwitchParameter] $Rescan,

    # Hide progress information during help files rescan.
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.SwitchParameter] $HideProgress
)



<#
.SYNOPSIS
pomoc displays full help on the selected item.
.DESCRIPTION
This is planed as a replacement of Get-Help cmdlet.
.EXAMPLE
pomoc -Name Get_Command
.EXAMPLE
[PSCustomObject]@{Name="Get-Help"} | pomoc
#>
#################
# function Main #
#################
function Main
{
    [CmdletBinding(DefaultParameterSetName='AllUsersView')]
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
        [System.String] $Path,

        # Displays help only for items in the specified category and their aliases. Conceptual articles are in the HelpFile category.
        [ValidateSet('Alias','Cmdlet','Provider','General','FAQ','Glossary','HelpFile','ScriptCommand','Function','Filter','ExternalScript','All','DefaultHelp','Workflow','DscResource','Class','Configuration')]
        [System.String[]] $Category,

        # Displays commands with the specified component value, such as Exchange . Enter a component name. Wildcard characters are per mitted. This parameter has no effect on displays of conceptual ( About_ ) help.
        [System.String[]] $Component,

        # Displays help for items with the specified functionality. Enter the functionality. Wildcard characters are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
        [System.String[]] $Functionality,

        # Displays help customized for the specified user role. Enter a role. Wildcard characters are permitted.
        #
        # Enter the role that the user plays in an organization. Some cmdlets display different text in their help files based on the value of this parameter. This parameter has no effect on help for the core cmdlets.
        [System.String[]] $Role,

        # Adds parameter descriptions and examples to the basic help display. This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
        #
        # In fact in script pomoc.ps1 this parameter is ignored. Always full info is displayed.
        [Parameter(ParameterSetName='DetailedView', Mandatory=$true)]
        [System.Management.Automation.SwitchParameter] ${Detailed},

        # Displays the entire help article for a cmdlet. Full includes parameter descriptions and attributes, examples, input and output object types, and additional notes.
        #
        # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
        #
        # In fact in script pomoc.ps1 this parameter is ignored. Always full info is displayed.
        [Parameter(ParameterSetName='AllUsersView')]
        [System.Management.Automation.SwitchParameter] $Full,

        # Displays only the name, synopsis, and examples. To display only the examples, type `(Get-Help <cmdlet-name>).Examples`.
        #
        # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
        #
        # In fact in script pomoc.ps1 this parameter is ignored. Always full info is displayed.
        [Parameter(ParameterSetName='Examples', Mandatory=$true)]
        [System.Management.Automation.SwitchParameter] $Examples,

        # Displays only the detailed descriptions of the specified parameters. Wildcards are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
        [Parameter(ParameterSetName='Parameters', Mandatory=$true)]
        [System.String] $Parameter,

        # Displays the online version of a help article in the default browser. This parameter is valid only for cmdlet, function, workflow, and script help articles. You can't use the Online parameter with `Get-Help` in a remote session.
        #
        # For information about supporting this feature in help articles that you write, see about_Comment_Based_Help (./About/about_Comment_Based_Help.md), and Supporting Online Help (/powershell/scripting/developer/module/supporting-online-help), and Writing Help for PowerShell Cmdlets (/powershell/scripting/developer/help/writing-help-for-windows-powershell-cmdlets).
        [Parameter(ParameterSetName='Online', Mandatory=$true)]
        [System.Management.Automation.SwitchParameter] $Online,

        # Displays the help topic in a window for easier reading. The window includes a Find search feature and a Settings box that lets you set options for the display, including options to display only selected sections of a help topic.
        #
        # The ShowWindow parameter supports help topics for commands (cmdlets, functions, CIM commands, workflows, scripts) and conceptual About articles. It does not support provider help.
        #
        # This parameter was introduced in PowerShell 3.0.
        [Parameter(ParameterSetName='ShowWindow', Mandatory=$true)]
        [System.Management.Automation.SwitchParameter] $ShowWindow,

        # Specifies the number of characters in each line of output. If this parameter is not used, the width is determined by the characteristics of the host. The default for the PowerShell console is 80 characters.
        [Parameter(Mandatory=$false)]
        [System.Int32] $Width,

        # Rescan help files. This is required when there are added new or newer help files.
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.SwitchParameter] $Rescan,

        # Hide progress information during help files rescan.
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.SwitchParameter] $HideProgress,

        #
        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable] $Functions
    )

    ###############################
    # function Main - Begin block #
    Begin
    {
        #==========================================================
        # Main function global variables
        #==========================================================

        $HelpInfo = @{
            Items = @();
            ItemIndex = @{};
            }

        $Work = @{
            ProfileDir = (Split-Path -Path $PROFILE -Parent)
            ItemsFile = (Join-Path -Path (Split-Path -Path $PROFILE -Parent) -ChildPath ".PS-pomoc.xml")

            # Used only during building $HelpInfo
            Functions = @{}    # list of all global functions
            Manifests = @{}
            UICultures = @($PSUICulture)

            # Used only during displaying results
            OutputWidth = $Width
            IndentSize = 4
            ListIndent = 0
            WasColon = $false
            CodeIndent = -1 # -1 means that it is unknown
            Colors = @{};
            }

        #----------------------------------------------------------
        # Adding en-US UICulture as a fallback culture
        if ("en-US" -notin $Work.UICultures)
        {
            $Work.UICultures += "en-US"
        }
        #----------------------------------------------------------
        # Configuring output width
        if ($Work.OutputWidth -eq 0)
        {
            try {
                $Work.OutputWidth = [System.Console]::WindowWidth
            }
            catch {
                $Work.OutputWidth = 80
            }
        }
        #----------------------------------------------------------
        # Configuring output colors
        if (-not (Test-Path Variable:psISE))
        {
            $Esc=[char]0x1B;
            $F_Default = "${Esc}[39m";
            # $F_Black = "${Esc}[30m";
            # $F_DarkRed = "${Esc}[31m";
            # $F_DarkGreen = "${Esc}[32m";
            # $F_DarkYellow = "${Esc}[33m";
            # $F_DarkBlue = "${Esc}[34m";
            # $F_DarkMagenta = "${Esc}[35m";
            # $F_DarkCyan = "${Esc}[36m";
            # $F_Gray = "${Esc}[37m";
            # $F_DarkGray = "${Esc}[90m";
            # $F_Red = "${Esc}[91m";
            $F_Green = "${Esc}[92m";
            $F_Yellow = "${Esc}[93m";
            # $F_Blue = "${Esc}[94m";
            $F_Magenta = "${Esc}[95m";
            $F_Cyan = "${Esc}[96m";
            # $F_White = "${Esc}[97m";

            $Work.Colors = @{
                Section = $F_Magenta;
                SubSection = $F_Cyan;
                ExtraSection = $F_Green;
                #Parameter = $F_Yellow;
                Code = $F_Yellow;
                Default = $F_Default;
                }
        }
        else
        {
            $Work.Colors = @{
                Section = '';
                SubSection = '';
                ExtraSection = '';
                #Parameter = '';
                Code = '';
                Default = '';
                }
        }

        #########################################
        # Main function - Begin block functions #
        #########################################

        #==========================================================
        # Universal utility functions
        #==========================================================

        ########################
        # function ToCamelCase #
        ########################
        function ToCamelCase
        {
            param (
                [System.String] $Str
            )

            ########################
            # function ToCamelCase #

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
        }   # function ToCamelCase #
            ########################

        ###################
        # function StrDef #
        ###################
        function StrDef
        {
            param (
                [System.String] $Str,
                [System.String] $Default
            )

            ###################
            # function StrDef #

            if ($null -eq $Str -or '' -eq $Str)
            {
                return $Default
            }
            return $Str
        }   # function StrDef #
            ###################


        #==========================================================
        # Parsing about*.txt HelpFile
        #==========================================================

        ###########################
        # function CleanParagraph #
        ###########################
        function CleanParagraph
        {
            param (
                [System.String] $Paragraph
            )

            ###########################
            # function CleanParagraph #

            # Replace all new lines and tabs with spaces
            $Paragraph = $Paragraph.Replace("`n", ' ')
            $Paragraph = $Paragraph.Replace("`t", ' ')

            # Compress every spaces string into single space
            while ($Paragraph.IndexOf('  ') -ne -1)
            {
                $Paragraph = $Paragraph.Replace('  ', ' ')
            }
            return $Paragraph.Trim()
        }   # function CleanParagraph #
            ###########################


        #############################
        # function ParseTxtHelpFile #
        #############################
        function ParseTxtHelpFile
        {
            param (
                [System.Collections.Hashtable] $Item
            )


            ##############################
            # function AddNavigationLink #
            ##############################
            function AddNavigationLink
            {
                param (
                    [System.Xml.XmlDocument] $XML,
                    [System.Xml.XmlElement] $ParentNode,
                    [System.String] $Line
                )

                ##############################
                # function AddNavigationLink #

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
                if (($Url -ne '') -and ($Text -eq '') -and ($null -ne $ParentNode.LastChild) -and ($ParentNode.LastChild.uri -eq ''))
                {
                    # LinkText and uri are splitted into two lines, so add uri
                    # to the previous navigationLink
                    $ParentNode.LastChild.ChildNodes[1].Set_innerText($Url)
                }
                else
                {
                    # Create navigationLink link elements with two children: linkText and uri.
                    # Set values of both children to $Text and $Url
                    $NavigationLink = $ParentNode.AppendChild($XML.CreateElement('navigationLink'))
                    $LinkText = $NavigationLink.AppendChild($XML.CreateElement('LinkText'))
                    $LinkText.Set_innerText($Text)
                    $Uri = $NavigationLink.AppendChild($XML.CreateElement('uri'))
                    $Uri.Set_innerText($Url)
                }
            }   # function AddNavigationLink #
                ##############################


            ###################################
            # function AddLinesToRelatedLinks #
            ###################################
            function AddLinesToRelatedLinks
            {
                param (
                    [System.Xml.XmlDocument] $XML,
                    [System.Int32] $SkipLines,
                    [System.String] $Paragraph
                )

                ###################################
                # function AddLinesToRelatedLinks #

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
            }   # function AddLinesToRelatedLinks #
                ###################################


            ###############################
            # function AddLinesToNewChild #
            ###############################
            function AddLinesToNewChild
            {
                param (
                    [System.Xml.XmlDocument] $XML,
                    [System.Xml.XmlElement] $ParentNode,
                    [System.String] $ChildName,
                    [System.Int32] $SkipLines,
                    [System.String] $Paragraph
                )

                ###############################
                # function AddLinesToNewChild #

                # Ignore first $SkipLines lines from the $Paragraph
                while (($SkipLines -gt 0) -and ($Paragraph.IndexOf("`n") -gt 0))
                {
                    $Paragraph = $Paragraph.Substring($Paragraph.IndexOf("`n")+1)
                    $SkipLines--
                }
                if (-not ($Paragraph -match ' *(- |-- |[0-9]+ =)'))
                {
                    $Paragraph = CleanParagraph $Paragraph
                }
                $Child = $ParentNode.AppendChild($XML.CreateElement($ChildName))
                $Child.Set_innerText($Paragraph)
            }   # function AddLinesToNewChild #
                ###############################


            ##################################
            # function ParseRegularParagraph #
            ##################################
            function ParseRegularParagraph
            {
                param (
                    [System.Collections.Hashtable] $Item,
                    [System.Xml.XmlDocument] $XML,
                    [System.String] $Paragraph
                )


                ####################################
                # function AddDescriptionParagraph #
                ####################################
                function AddDescriptionParagraph
                {
                    param (
                        [System.Xml.XmlDocument] $XML,
                        [System.String] $Paragraph
                    )

                    ####################################
                    # function AddDescriptionParagraph #

                    $Description = Select-XML -Xml $XML -XPath '/helpItems/command/description'
                    if ($null -eq $Description)
                    {
                        # There was no LONG DESCRIPTION header in file
                        $Description = $Command.AppendChild($XML.CreateElement('description'))
                    }
                    elseif ($null -eq (Get-Member -InputObject $Description -Name Node))
                    {
                        # There was no LONG DESCRIPTION header in file
                        $Description = $Command.AppendChild($XML.CreateElement('description'))
                    }
                    else
                    {
                        $Description = $Description.Node
                    }
                    AddLinesToNewChild $XML $Description 'para' 0 $Paragraph
                }   # function AddDescriptionParagraph #
                    ####################################


                ##################################
                # function StoreRegularParagraph #
                ##################################
                function StoreRegularParagraph
                {
                    param (
                        [System.Collections.Hashtable] $Item,
                        [System.Xml.XmlDocument] $XML,
                        [System.String] $Paragraph
                    )


                    ###########################
                    # function StoreParagraph #
                    ###########################
                    function StoreParagraph
                    {
                        param (
                            [System.Collections.Hashtable] $Item,
                            [System.Xml.XmlDocument] $XML,
                            [System.String] $Paragraph,
                            [System.Boolean] $AddToPrevious
                        )

                        ###########################
                        # function StoreParagraph #
                        if ($Item.CurrentSectionName -eq 'NOTES')
                        {
                            $Alert = (Select-XML -Xml $XML -XPath '/helpItems/command/alertSet/alert').Node
                            if ($AddToPrevious)
                            {
                                # Add to previous !?!?!? !!! ???
                            }
                            else
                            {
                                AddLinesToNewChild $XML $Alert 'para' 0 $Paragraph
                            }
                        }
                        #elseif ($null -ne (Get-Member -InputObject $Item -Name CurrentExtraSectionNode))
                        elseif ($null -ne $Item.CurrentExtraSectionNode)
                        {
                            if ($AddToPrevious)
                            {
                                # Add to previous !?!?!? !!! ???
                            }
                            else
                            {
                                AddLinesToNewChild $XML $Item.CurrentExtraSectionNode 'para' 0 $Paragraph
                            }
                        }
                        elseif ($Item.CurrentSectionName -eq 'EXAMPLES')
                        {
                            $Examples = (Select-XML -Xml $XML -XPath '/helpItems/command/examples').Node
                            $Example = $Examples.AppendChild($XML.CreateElement('example'))
                            $Title = $Example.AppendChild($XML.CreateElement('title'))
                            $Title.Set_innerText('Example')
                            $Item.CurrentExtraSectionNode = $Example
                            $Item.CurrentSectionName = 'EXAMPLE'
                            AddLinesToNewChild $XML $Item.CurrentExtraSectionNode 'para' 0 $Paragraph
                        }
                        else
                        {
                            if ($AddToPrevious)
                            {
                                # Add to previous !?!?!? !!! ???
                            }
                            else
                            {
                                AddDescriptionParagraph $XML $Paragraph
                            }
                        }
                    }   # function StoreParagraph #
                        ###########################


                    ##################################
                    # function StoreRegularParagraph #

                    $AddToPrevious = $false
                    if ($Paragraph -match '^ *([0-9]+ =|-+) *')
                    {
                        # Numbered/bulleted list elemnent
                        $Work.ListIndent = 0
                        while ($Paragraph.Substring($Work.ListIndent,1) -match '[- =0-9]')
                        {
                            $Work.ListIndent++
                        }
                        while ($Paragraph.Substring($Paragraph.Length-1,1) -eq "`n")
                        {
                            $Paragraph = $Paragraph.Substring(0,$Paragraph.Length-1)
                        }
                        $pos = $Paragraph.IndexOf("`n")
                        while ($pos -ne -1)
                        {
                            if ($Paragraph.Substring($pos+1,$Work.ListIndent) -eq (' '*$Work.ListIndent))
                            {
                                $Paragraph = $Paragraph.Substring(0,$pos)+' '+
                                    $Paragraph.Substring($pos+1+$Work.ListIndent)
                                $pos = $Paragraph.IndexOf("`n")
                            }
                            elseif ($Paragraph.Substring($pos+1) -match '^ *([0-9]+ =|-+) *')
                            {
                                StoreParagraph $Item $XML $Paragraph.Substring(0,$pos) $AddToPrevious
                                StoreRegularParagraph $Item $XML $Paragraph.Substring($pos+1)
                                return
                            }
                            else
                            {
                                break
                            }
                        }
                    }
                    elseif (($Work.ListIndent -gt 0) -and
                            ($Paragraph.Substring(0,$Work.ListIndent) -eq (' '*$Work.ListIndent)))
                    {
                        # Numbered/bulleted list element continuation
                        $AddToPrevious = $true
                        Write-Verbose "`$AddToPrevious = `$true"
                    }
                    else
                    {
                        # Regular paragraph
                        $Work.ListIndent = 0
                    }
                    StoreParagraph $Item $XML $Paragraph $AddToPrevious
                }   # function StoreRegularParagraph #
                    ##################################



                ###############################
                # function StoreCodeParagraph #
                ###############################
                function StoreCodeParagraph
                {
                    param (
                        [System.Collections.Hashtable] $Item,
                        [System.Xml.XmlDocument] $XML,
                        [System.String] $Paragraph
                    )

                    ###############################
                    # function StoreCodeParagraph #

                    if (-1 -eq $Work.CodeIndent)
                    {
                        $Work.CodeIndent = 0
                        while (' ' -eq $Paragraph.Substring($Work.CodeIndent,1))
                        {
                            $Work.CodeIndent++
                        }
                    }
                    $Lines = $Paragraph.Split("`n")
                    $Paragraph = ''
                    foreach ($Line in $Lines)
                    {
                        if (($Work.CodeIndent -lt $Line.Length) -and
                            (' '*$Work.CodeIndent -eq $Line.Substring(0,$Work.CodeIndent)))
                        {
                            $Line = $Line.Substring($Work.CodeIndent)
                        }
                        if ('' -eq $Paragraph)
                        {
                            $Paragraph = $Line
                        }
                        else
                        {
                            $Paragraph += "`n"+$Line
                        }
                    }

                    if ($Item.CurrentSectionName -eq 'NOTES')
                    {
                        $ParentNode = (Select-XML -Xml $XML -XPath '/helpItems/command/alertSet/alert').Node
                    }
                    elseif ($Item.CurrentSectionName -eq 'EXAMPLES')
                    {
                        $Examples = (Select-XML -Xml $XML -XPath '/helpItems/command/examples').Node
                        $Example = $Examples.AppendChild($XML.CreateElement('example'))
                        $Title = $Example.AppendChild($XML.CreateElement('title'))
                        $Title.Set_innerText('Example')
                        $Item.CurrentExtraSectionNode = $Example
                        $Item.CurrentSectionName = 'EXAMPLE'
                        $ParentNode = $Item.CurrentExtraSectionNode
                    }
                    #elseif ($null -ne (Get-Member -InputObject $Item -Name CurrentExtraSectionNode))
                    elseif ($null -ne $Item.CurrentExtraSectionNode)
                    {
                        $ParentNode = $Item.CurrentExtraSectionNode
                    }
                    else
                    {
                        # There is formatted paragraph or code ????!!!!
                        $Description = Select-XML -Xml $XML -XPath '/helpItems/command/description'
                        if ($null -ne $Description)
                        {
                            if ($null -ne (Get-Member -InputObject $Description -Name Node))
                            {
                                $Description = $Description.Node
                            }
                        }
                        if ($null -eq $Description)
                        {
                            # There was no LONG DESCRIPTION header in file
                            $Description = $Command.AppendChild($XML.CreateElement('description'))
                        }
                        $ParentNode = $Description
                    }
                    if ($null -ne $ParentNode)
                    {
                        $Child = $ParentNode.AppendChild($XML.CreateElement('code'))
                        $Child.Set_innerText($Paragraph)
                    }
                }   # function StoreCodeParagraph #
                    ###############################


                ################################
                # function ExtraSectionHeading #
                ################################
                function ExtraSectionHeading
                {
                    param (
                        [System.String] $Paragraph
                    )

                    ################################
                    # function ExtraSectionHeading #

                    if ($Paragraph.IndexOf("`n") -ne -1)
                    {
                        $FirstLine = $Paragraph.Substring(0,$Paragraph.IndexOf("`n"))
                    }
                    else
                    {
                        $FirstLine = $Paragraph
                    }
                    if ($FirstLine.Trim().Length -eq 0)
                    {
                        return ''
                    }
                    $FirstLine = $FirstLine.TrimEnd()
                    if ($FirstLine.Substring(0,1) -in @(' ','-',"`t"))
                    {
                        return ''
                    }
                    $FirstLine = $FirstLine.TrimStart("`t ")
                    if ($FirstLine -in @('or', 'and'))
                    {
                        return ''
                    }
                    if ($FirstLine -eq 'Examples:')
                    {
                        return 'Examples'
                    }
                    if ($FirstLine.Substring($FirstLine.Length-1,1) -in @('.', ':'))
                    {
                        return ''
                    }
                    if (($FirstLine.Length -gt 45) -and ($FirstLine.Length -ne $Paragraph.Trim("`t ").Length))
                    {
                        return ''
                    }
                    if ($FirstLine -match ' is ')
                    {
                        return ''
                    }
                    if (-not ($FirstLine -match '^[-\*: a-z0-9$?_\(\)]+$'))
                    {
                        return ''
                    }
                    return $FirstLine
                }   # function ExtraSectionHeading #
                    ################################


                ##################################
                # function ParseRegularParagraph #

                #Write-Verbose $Paragraph
                switch ($Item.CurrentSectionName)
                {
                    {($_ -eq 'NAME') -or
                     ($_ -eq 'SYNOPSIS')}
                        {
                            $Details = Select-XML -Xml $XML -XPath '/helpItems/command/details'
                            if ($null -ne $Details)
                            {
                                $Details = $Details.Node
                            }
                            else
                            {
                                # There was no TOPIC nor item name. Put name into XML.
                                $Details = $Command.AppendChild($XML.CreateElement('details'))
                                $Name = $Details.AppendChild($XML.CreateElement('name'))
                                $Name.Set_innerText($Item.DisplayName)
                            }
                            if ($null -eq (Get-Member -InputObject $XML.helpItems.command.details -Name description))
                            {
                                # The synopsis was separated with empty line from SHORT DESCRIPTION.
                                $Description = $Details.AppendChild($XML.CreateElement('description'))
                                $Para = $Description.AppendChild($XML.CreateElement('para'))
                                if ($UsedIndentation -eq 0)
                                {
                                    $UsedIndentation = $Paragraph.Length-$Paragraph.TrimStart().Length
                                }
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
                     ($_ -eq 'NOTES') -or
                     ($_ -eq 'EXAMPLES') -or
                     ($_ -eq 'EXAMPLE')}
                        {
                            if ($Paragraph -match '^ *EXAMPLE [0-9]+:')
                            {
                                if (($Item.CurrentSectionName -eq 'DESCRIPTION') -and
                                    ($Item.CurrentExtraSectionName -match 'EXAMPLES'))
                                {
                                    $Examples = $Command.AppendChild($XML.CreateElement('examples'))
                                    foreach ($Child in $Item.CurrentExtraSectionNode.ChildNodes)
                                    {
                                        [void]$Examples.AppendChild($Child.Clone())
                                    }
                                    [void]$Description.RemoveChild($Item.CurrentExtraSectionNode)
                                }
                                else
                                {
                                    $Examples = (Select-XML -Xml $XML -XPath '/helpItems/command/examples').Node
                                }
                                $Example = $Examples.AppendChild($XML.CreateElement('example'))
                                $Title = $Example.AppendChild($XML.CreateElement('title'))
                                $Title.Set_innerText($Paragraph.Trim())
                                $Item.CurrentExtraSectionNode = $Example
                                $Item.CurrentSectionName = 'EXAMPLE'
                                break
                            }
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
                                        # Need to check if we have code or fomatted paragraph
                                        if ($Paragraph.Substring(0,$UsedIndentation) -eq (' '*$UsedIndentation))
                                        {
                                            if (($Paragraph.Substring($UsedIndentation,1) -eq ' ') -and
                                                ($Paragraph.TrimStart() -notmatch '\[!(NOTE|WARNING|IMPORTANT)\]'))
                                            {
                                                StoreCodeParagraph $Item $XML $Paragraph
                                                break
                                            }
                                        }
                                        # Actually regular paragraph, add to appropriate section
                                        StoreRegularParagraph $Item $XML $Paragraph
                                        $Work.CodeIndent = -1
                                    }
                                default
                                    {
                                        $Item.CurrentExtraSectionName = $ExtraSection.ToUpper()
                                        #Write-Verbose "ExtraSection: $($Item.CurrentExtraSectionName)"
                                        $Description = Select-XML -Xml $XML -XPath '/helpItems/command/description'
                                        if ($null -eq $Description)
                                        {
                                            return
                                        }
                                        if ($null -eq (Get-Member -InputObject $Description -Name Node))
                                        {
                                            return
                                        }
                                        $Description = $Description.Node
                                        $Item.CurrentExtraSectionNode = $Description.AppendChild($XML.CreateElement('section'))
                                        $Name = $Item.CurrentExtraSectionNode.AppendChild($XML.CreateElement('name'))
                                        $Name.Set_innerText($Item.CurrentExtraSectionName)
                                        $Item.CurrentSectionName = 'DESCRIPTION'
                                        if ($Paragraph.IndexOf("`n") -ne -1)
                                        {
                                            ParseRegularParagraph $Item $XML $Paragraph.Substring($Paragraph.IndexOf("`n")+1)
                                        }
                                    }
                            }
                        }
                }
            }   # function ParseRegularParagraph #
                ##################################


            #############################
            # function ParseTxtHelpFile #

            $Item.CurrentExtraSectionNode = $null
            Write-Verbose "Parsing HelpFile $($Item.File)"
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

            # Convert file name into help item name
            $Item.DisplayName = $Item.Name.Replace('_', ' ')
            $Item.DisplayName = ToCamelCase $Item.DisplayName

            # Gues indentation used in file for regular paragraphs
            $UsedIndentation = 0

            # parse paragraphs converting them into MAML like XML object
            #----------------------------------------------------------=
            $XML = [System.Xml.XmlDocument]'<?xml version="1.0" encoding="utf-8"?><helpItems />'
            $Command = $XML.ChildNodes[1].AppendChild($XML.CreateElement('command'))
            $Item.CurrentSectionName = ''
            foreach ($Paragraph in $Paragraphs)
            {
                if ($Paragraph.IndexOf("`n") -ne -1)
                {
                    $FirstLine = $Paragraph.Substring(0,$Paragraph.IndexOf("`n")).TrimEnd()
                }
                else
                {
                    $FirstLine = $Paragraph.TrimEnd()
                }
                switch -Regex ($FirstLine)
                {
                    ('^(TOPIC|'+$Item.DisplayName+'|'+$Item.Name+')$')
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
                    # In about_Type_Accelerators.help.txt there is a typo: 'SHORT DESRIPTION'
                    '^(SHORT DES(C)?RIPTION|SYNOPSIS)$'
                        {
                            if ($Item.CurrentSectionName -eq 'DESCRIPTION')
                            {
                                ParseRegularParagraph $Item $XML $Paragraph
                                break
                            }
                            if ($null -eq (Get-member -InputObject $XML.helpItems.command -Name details))
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
                                $Paragraph = $Paragraph.Substring($Paragraph.IndexOf("`n")+1)
                                if ($UsedIndentation -eq 0)
                                {
                                    $UsedIndentation = $Paragraph.Length-$Paragraph.TrimStart().Length
                                }
                                AddLinesToNewChild $XML $Description 'para' 0 $Paragraph
                            }
                            $Item.CurrentSectionName = 'SYNOPSIS'
                        }
                    '^((LONG|DETAILED) )?DESCRIPTION$'
                        {
                            if ($Item.CurrentSectionName -eq 'DESCRIPTION')
                            {
                                $Description = (Select-XML -Xml $XML -XPath '/helpItems/command/description').Node
                                if ($null -eq (Get-Member -InputObject $Description.para -Name Count))
                                {
                                    # There was only one extra paragraph before section header,
                                    # So we can move it to Synopsis
                                    $Text = $XML.helpItems.command.details.description.para + ' ' +
                                            $XML.helpItems.command.description.para
                                    $Synopsis = (Select-XML -Xml $XML -XPath '/helpItems/command/details/description/para').Node
                                    $Synopsis.Set_innerText($Text)
                                    $Para = (Select-XML -Xml $XML -XPath '/helpItems/command/description/para').Node
                                    if ($null -eq (Get-Member -InputObject $Para -Name Count))
                                    {
                                        $Description.RemoveChild($Para) | Out-Null
                                    }
                                    else
                                    {
                                        return $XML
                                        Write-Error "Problems with Node in DESCRIPTION in $($Item.File)"
                                        Show-XML -InputObject $XML
                                    }
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
                    <# 
                    '^EXAMPLES$'
                        {
                            if ($Item.Name -ne 'about_Preference_Variables' -and
                                $Item.Name -ne 'about_Comment_Based_Help')
                            {
                                $Examples = $Command.AppendChild($XML.CreateElement('examples'))
                                $Item.CurrentExtraSectionNode = $null
                                $Item.CurrentSectionName = 'EXAMPLES'
                            }
                            else
                            {
                                ParseRegularParagraph $Item $XML $Paragraph
                            }
                        }
                     #>
                    '^(NOTES|REMARKS)$'
                        {
                            #if ($Item.Name -ne 'about_Updatable_Help')
                            if ($Item.Name -ne 'about_Comment_Based_Help')
                            {
                                $AlertSet = $Command.AppendChild($XML.CreateElement('alertSet'))
                                $Alert = $AlertSet.AppendChild($XML.CreateElement('alert'))
                                $Item.CurrentSectionName = 'NOTES'
                                $Item.CurrentExtraSectionNode = $null
                            }
                            else
                            {
                                ParseRegularParagraph $Item $XML $Paragraph
                            }
                        }
                    '^(SEE ALSO|RELATED LINKS)$'
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
        }   # function ParseTxtHelpFile #
            #############################


        #==========================================================
        # File searching stuff
        #==========================================================


        ####################
        # function AddItem #
        ####################
        function AddItem
        {
            param (
                [System.Boolean] $MarkFunc,
                [System.Collections.Hashtable] $Item
            )

            ####################
            # function AddItem #

            #Write-Verbose "Adding Item $($Item.Name)"
            if ($null -eq $HelpInfo.ItemIndex[$Item.Name])
            {
                if ($null -ne $Work.Functions[$Item.Name])
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
            elseif ($null -ne $Item.LastWriteTime)
            {
                if ($Item.LastWriteTime -gt $HelpInfo.Items[$HelpInfo.ItemIndex[$Item.Name]].LastWriteTime)
                {
                    <# Write-Host ("Trying to overwrite " + $Item.Name)
                    Write-Host ("  Old file was: "+$HelpInfo.Items[$HelpInfo.ItemIndex[$Item.Name]].File)
                    Write-Host ("  New file is:  "+$Item.File)
                    !!!!????!?!?!?! #>
                    $HelpInfo.Items[$HelpInfo.ItemIndex[$Item.Name]] = $Item
                }
            }
        }   # function AddItem #
            ####################


        ##########################
        # function FindHelpFiles #
        ##########################
        function FindHelpFiles
        {
            param ()


            ###################################
            # function CheckModuleForManifest #
            ###################################
            function CheckModuleForManifest
            {
                param (
                    [System.String] $ModulePath,
                    [System.String] $Version = ''
                )


                Write-Verbose "Function called CheckModuleForManifest `"$ModulePath`""
                $ModuleName = Split-Path -Path $ModulePath -Leaf
                if ($Version -ne '')
                {
                    $ModulePath = Join-Path -Path $ModulePath -ChildPath $Version
                }
                Write-Verbose "Checking module $ModulePath"
                $ModuleManifest = Join-Path -Path $ModulePath -ChildPath "$ModuleName.psd1"
                if (Test-Path -Path $ModuleManifest)
                {
                    # $Info = $null
                    try
                    {
                        $Info = Import-PowerShellDataFile -Path $ModuleManifest -ErrorAction SilentlyContinue
                    }
                    catch
                    {
                        $Info = @{
                            PowerShellVersion = ""
                            ModuleName = ""
                            ModuleVersion = ""
                            FunctionsToExport = @()
                            CmdletsToExport = @()
                            AliasesToExport = @()
                        }
                        foreach ($Line in (Get-Content -Path $ModuleManifest))
                        {
                            # Write-Host $Line
                            switch ($line.Split("=")[0])
                            {
                                "PowerShellVersion"
                                    {
                                        $Info.PowerShellVersion = $line.Split("=")[1]
                                        break
                                    }
                                "ModuleName"
                                    {
                                        $Info.ModuleName = $line.Split("=")[1]
                                        break
                                    }
                                "ModuleVersion"
                                    {
                                        $Info.ModuleVersion = $line.Split("=")[1]
                                        break
                                    }
                                "FunctionsToExport"
                                    {
                                        $Info.FunctionsToExport = Invoke-Expression $line.Split("=")[1]
                                        break
                                    }
                                "CmdletsToExport"
                                    {
                                        $Info.CmdletsToExport = Invoke-Expression $line.Split("=")[1]
                                        break
                                    }
                                "AliasesToExport"
                                    {
                                        $Info.AliasesToExport = Invoke-Expression $line.Split("=")[1]
                                        break
                                    }
                            }
                        }
                    }
                    if ($null -eq $Info["ModuleName"])
                    {
                        $Info.ModuleName = ""
                    }
                    if ($null -eq $Info["FunctionsToExport"])
                    {
                        $Info.FunctionsToExport = @()
                    }
                    if ($null -eq $Info["CmdletsToExport"])
                    {
                        $Info.CmdletsToExport = @()
                    }
                    if ($null -eq $Info["AliasesToExport"])
                    {
                        $Info.AliasesToExport = @()
                    }
                    if ($null -eq $Info["PowerShellVersion"])
                    {
                        $Info.PowerShellVersion = ""
                    }
                    foreach ($Exp in (@{Type='Function'; Field='FunctionsToExport'},
                                        @{Type='Cmdlet'; Field='CmdletsToExport'},
                                        @{Type='Alias'; Field='AliasesToExport'}))
                    {
                        $Locals = $Info.($Exp.Field)
                        if ($null -ne $Locals)
                        {
                            foreach ($Cmd in $Locals)
                            {
                                if (($null -ne $Cmd) -and ('' -ne $Cmd))
                                {
                                    $Manifest = @{Used = $false
                                                  Category = $Exp.Type
                                                  PowerShellVersion = StrDef $Info.PowerShellVersion "0.0"
                                                  ModuleName = StrDef $Info.ModuleName ""
                                                  ModuleVersion = StrDef $Info.ModuleVersion "0.0"
                                                 }
                                    if ($null -ne $Work.Manifests[$Cmd])
                                    {
                                        if ($Manifest.PowerShellVersion -lt $Work.Manifests[$Cmd].PowerShellVersion)
                                        {
                                            continue
                                        }
                                        if ($Manifest.PowerShellVersion -eq $Work.Manifests[$Cmd].PowerShellVersion -and $Manifest.ModuleVersion -lt $Work.Manifests[$Cmd].ModuleVersion)
                                        {
                                            continue
                                        }
                                    }
                                    $Work.Manifests[$Cmd] = $Manifest
                                }
                            }
                        }
                    }
                }
            }   # function CheckModuleForManifest #
                ###################################


            ###################################
            # function CheckModuleForHelpFile #
            ###################################
            function CheckModuleForHelpFile
            {
                param (
                    [System.String] $ModulePath,
                    [System.String] $Version = ''
                )


                ##############################
                # function CheckTxtHelpFiles #
                ##############################
                function CheckTxtHelpFiles
                {
                    param (
                        [System.String] $Path,
                        [System.String] $ModuleName
                    )


                    ##################################
                    # function GetModuleAndOnlineURI #
                    ##################################
                    function GetModuleAndOnlineURI
                    {
                        param (
                            [System.String] $Name,
                            [System.String] $ModuleName
                        )

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

                        ##################################
                        # function GetModuleAndOnlineURI #

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
                                $version = '{0}.{1}' -f $PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor

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
                    }   # function GetModuleAndOnlineURI #
                        ##################################


                    ##############################
                    # function CheckTxtHelpFiles #

                    if ( -not (Test-Path -Path $Path -PathType Container))
                    {
                        return
                    }
                    $Files = @(Get-ChildItem (Join-Path -Path $Path -ChildPath *.help.txt))
                    if ($Files.Count -gt 0)
                    {
                        Write-Verbose "Checking txt HelpFiles in module $Path"
                        foreach ($File in $Files)
                        {
                            $Name = $File.Name -replace '.help.txt',''
                            $URI = GetModuleAndOnlineURI $Name $ModuleName
                            if ($URI.Module -eq 'drop it !!!')
                            {
                                continue
                            }
                            $Item = @{Name = $Name;
                                      ModuleName = $URI.Module;
                                      File = $File.FullName;
                                      LastWriteTime = $File.LastWriteTime;
                                      OnlineURI = $URI.URI;
                                      Format = 'txt';
                                      Index = -1;
                                      Category = 'HelpFile';
	                                  Component = '';
                                      Functionality = '';
                                      Role = '';
                                      Synopsis = '';
                                      Aliases = @();
                                      CommonParameters = $false}
                            $XML = ParseTxtHelpFile $Item
                            $Item.Synopsis = ""
                            if ($null -ne (Get-member -InputObject $XML -Name helpItems))
                            {
                                if ($null -ne (Get-member -InputObject $XML.helpItems.command -Name details))
                                {
                                    if ($null -ne (Get-member -InputObject $XML.helpItems.command.details -Name description))
                                    {
                                        $Item.Synopsis = CleanParagraph $XML.helpItems.command.details.description.para
                                    }
                                }
                            }
                            AddItem $true $Item
                        }
                    }
                    return
                }   # function CheckTxtHelpFiles #
                    ##############################


                ##############################
                # function CheckXmlHelpFiles #
                ##############################
                function CheckXmlHelpFiles
                {
                    param (
                        [System.String] $Path,
                        [System.String] $ModuleName
                    )


                    #########################
                    # function CheckXMLFile #
                    #########################
                    function CheckXMLFile
                    {
                        param (
                            [System.String] $Path,
                            [System.String] $ModuleName,
                            [System.String] $File
                        )

                        #########################
                        # function ParseCommand #
                        #########################
                        function ParseCommand
                        {
                            param (
                                [System.Xml.XmlElement] $Command
                            )

                            $Category = 'Cmdlet'
                            if ($null -ne $Work.Manifests[$Command.details.name])
                            {
                                $Category = $Work.Manifests[$Command.details.name].Category
                                $Work.Manifests[$Command.details.name].Used = $true
                            }
                            $CmdInfo = Get-Command -Name $Command.details.name -ErrorAction SilentlyContinue
                            if ($null -ne $CmdInfo)
                            {
                                if ($Category -eq 'Function')
                                {
                                    if ($null -eq (Get-Member -InputObject $CmdInfo -Name CmdletBinding))
                                    {
                                        $CommonParameters = $false
                                    }
                                    else
                                    {
                                        $CommonParameters = $CmdInfo.CmdletBinding
                                    }
                                }
                                else
                                {
                                    $CommonParameters = $CmdInfo.Definition.IndexOf('<CommonParameters>') -ne -1
                                }
                            }
                            else
                            {
                                $CommonParameters = $false
                            }
                            AddItem $true @{Name = $Command.details.name;
                                            ModuleName = $ModuleName;
                                            File = $FilePath;
                                            LastWriteTime = $null;
                                            OnlineURI = '';
                                            Format = 'xml';
                                            Index = $Index;
                                            Category = $Category;
                                            Component = '';
                                            Functionality = '';
                                            Role = '';
                                            Synopsis = $Command.details.description.para;
                                            Aliases = @();
                                            CommonParameters = $CommonParameters}
                        } # function ParseCommand #
                          #########################

                        #########################
                        # function CheckXMLFile #

                        $FilePath = Join-Path -Path $Path -ChildPath $File
                        Write-Verbose "Checking file $FilePath"
                        $XML = [System.Xml.XmlDocument](Get-Content $FilePath)
                        if ($null -ne $XML.helpItems)
                        {
                            if ($null -ne $XML.helpItems.command)
                            {
                                if (($XML.helpItems.command).GetType().Name -eq 'Object[]')
                                {
                                    for ($Index = 0; $Index -lt ($XML.helpItems.command).Count; $Index++)
                                    {
                                        $command = ($XML.helpItems.command)[$Index]
                                        ParseCommand $command
                                    }
                                }
                                else
                                {
                                    $Index = $null
                                    ParseCommand $XML.helpItems.command
                                }
                            }
                        }
                    }   # function CheckXMLFile #
                        #########################


                    ##############################
                    # function CheckXmlHelpFiles #

                    if ( -not (Test-Path -Path $Path -PathType Container))
                    {
                        return
                    }
                    $Pattern = '-help.xml'
                    Write-Verbose "Checking xml HelpFiles in module $Path"
                    $Files = @(Get-ChildItem (Join-Path -Path $Path -ChildPath *$Pattern))
                    if ($null -eq $Files)
                    {
                        return
                    }
                    if ($Files.Count -eq 0)
                    {
                        return
                    }
                    foreach ($File in $Files.Name)
                    {
                        if ($ModuleName -eq '')
                        {
                            $MN = $File.Remove($File.Length - $Pattern.Length)
                            $MN = $MN -replace '.dll',''
                            $MN = $MN -replace 'Microsoft.PowerShell.Commands.','Microsoft.PowerShell.'
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
                            }
                            CheckXMLFile $Path $MN $File
                        }
                        else
                        {
                            CheckXMLFile $Path $ModuleName $File
                        }
                    }
                }   # function CheckXmlHelpFiles #
                    ##############################


                ###################################
                # function CheckModuleForHelpFile #

                Write-Verbose "Function called CheckModuleForHelpFile `"$For`" `"$ModulePath`""
                $ModuleName = Split-Path -Path $ModulePath -Leaf
                foreach ($UICulture in $Work.UICultures)
                {
                    if ($ModuleName -eq $UICulture)
                    {
                        CheckTxtHelpFiles $ModulePath ""
                        CheckXmlHelpFiles $ModulePath ""
                        return
                    }
                }
                if ($ModuleName -eq "v1.0" -or $ModuleName -eq "7")
                {
                    $ModuleName = ""
                }
                if ($Version -ne '')
                {
                    $ModulePath = Join-Path -Path $ModulePath -ChildPath $Version
                }
                Write-Verbose "Checking module $ModulePath"
                foreach ($UICulture in $Work.UICultures)
                {
                    $UICulturePath = Join-Path -Path $ModulePath -ChildPath $UICulture
                    CheckTxtHelpFiles $UICulturePath $ModuleName
                    CheckXmlHelpFiles $UICulturePath $ModuleName
                }
            }   # function CheckModuleForHelpFile #
                ###################################


            ########################
            # function CheckModule #
            ########################
            function CheckModule
            {
                param (
                    [System.String] $For,
                    [System.String] $ModulePath,
                    [System.String] $Version = ''
                )

                if ($For -eq "Manifest")
                {
                    CheckModuleForManifest $ModulePath $Version
                }
                else
                {
                    CheckModuleForHelpFile $ModulePath $Version
                }
            }   # function CheckModule #
            ########################

            ##########################
            # function FindHelpFiles #

            #----------------------------------------------------------
            # Create a list of directories to search for manifest and 
            # help files.
            $DirList = @($PSHOME)
            foreach ($ModulePath in (($env:PSModulePath).Split(';')))
            {
                # If last item in path is Modules then start from its parent
                if ((Split-Path -Path $ModulePath -Leaf) -eq "Modules")
                {
                    $ModulePath = Split-Path -Path $ModulePath -Parent
                }
                if (-not (Test-Path -Path $ModulePath -PathType Container))
                {
                    # $ModulePath do not exists or is not a directory
                    continue
                }
                if ($ModulePath -notin $DirList)
                {
                    $DirList += $ModulePath
                    Write-Verbose "Path to scan for help files: $ModulePath"
                }
            }
            #----------------------------------------------------------
            # First seach for manifests and create list of command 
            # categories.
            foreach ($SearchFor in ("Manifest", "HelpFile"))
            {
                if ($showProgress)
                {
                    $DirNo = 0
                    $DirPercentComplete = 0
                }
                #----------------------------------------------------------
                # Search for manifests or help files in module directories
                foreach ($Directory in $DirList)
                {
                    # $SubDirList will contain list of directories which can contain manifest of help files.
                    $SubDirList = @()
                    # First check subdirectory "Help" of module path.
                    $Path = Join-Path -Path $Directory -ChildPath "Help"
                    if (Test-Path -Path $Path -PathType Container)
                    {
                        $SubDirs = (Get-ChildItem -Path $Path -Directory).FullName
                        if ($null -ne $SubDirs)
                        {
                            $SubDirList += $SubDirs
                        }
                    }
                    # Next check subdirectory "Modules" of module path.
                    $Path = Join-Path -Path $Directory -ChildPath "Modules"
                    if (Test-Path -Path $Path -PathType Container)
                    {
                        $SubDirs = (Get-ChildItem -Path $Path -Directory).FullName
                        if ($null -ne $SubDirs)
                        {
                            $SubDirList += $SubDirs
                        }
                    }
                    if ($SubDirList.Count -eq 0)
                    {
                        continue
                    }
                    if ($showProgress)
                    {
                        Write-Progress -Activity "Analyzing directories" -Status "Progress:" -Id 0 `
                            -CurrentOperation $Directory -PercentComplete $DirPercentComplete
                        $moduleCount = 1 + $SubDirList.Count
                        $moduleNo = 0
                        $modulePercentComplete = 0
                    }
                    #----------------------------------------------------------
                    # Checking $Directory directory
                    CheckModule $SearchFor $Directory
                    #----------------------------------------------------------
                    # Checking Help and Modules directories
                    foreach ($ModulePath in $SubDirList)
                    {
                        if ($showProgress)
                        {
                            Write-Progress -Activity ("Checking module " + $ModulePath) -Status "Progress:" -Id 1 -ParentId 0 -PercentComplete $modulePercentComplete
                        }
                        CheckModule $SearchFor $ModulePath
                        $Version = ''
                        foreach ($SubDir in (Get-Childitem $ModulePath -Directory))
                        {
                            if ($SubDir.Name -match '^([0-9]\.)+[0-9]+$')
                            {
                                if ([Version](StrDef $Version "0.0") -lt [Version]($SubDir.Name))
                                {
                                    $Version = $SubDir.Name
                                }
                            }
                        }
                        if ($Version -ne '')
                        {
                            Write-Verbose "Checking module: $ModulePath"
                            CheckModule $SearchFor $ModulePath $Version
                        }
                        if ($showProgress)
                        {
                            $moduleNo += 1
                            $modulePercentComplete = ($moduleNo*100/$moduleCount)
                        }
                    }
                    if ($showProgress)
                    {
                        Write-Progress -Activity "Checking modules" -Status "Progress:" -Id 1 -ParentId 0 -Completed
                        $DirNo += 1
                        $DirPercentComplete = ($DirNo*100/$DirList.Count)
                    }
                }
            }
            if ($showProgress)
            {
                Write-Progress -Activity "Analyzing help files" -Completed
            }
            
            #----------------------------------------------------------
            # Added help items for global functions for which we do not
            # found yet help.
            foreach ($Function in $Work.Functions.keys)
            {
                if ($null -ne $Work.Functions[$Function])
                {
                    $CommonParameters = $false
                    $CmdInfo = Get-Command -Name $Function -ErrorAction SilentlyContinue
                    if ($null -ne $CmdInfo)
                    {
                        $CommonParameters = $CmdInfo.Definition.IndexOf('<CommonParameters>') -ne -1
                    }
                    AddItem $false @{Name = $Function;
                                     ModuleName = '';
                                     File = '';
                                     LastWriteTime = $null;
                                     OnlineURI = '';
                                     Format = '';
                                     Index = -1;
                                     Category = 'Function';
                                     Component = '';
                                     Functionality = '';
                                     Role = '';
                                     Synopsis = '...';
                                     Aliases = @();
                                     CommonParameters = $CommonParameters
                                    }
                }
            }
            $Work.Functions = @{}
            
            #----------------------------------------------------------
            # Extract comment-based help from function or cmdlet definition.
            foreach ($CommandName in $Work.Manifests.keys)
            {
                if (-not $Work.Manifests[$CommandName].Used -and
                    $Work.Manifests[$CommandName].Category -ne 'Alias')
                {
                    $ModuleName = $Work.Manifests[$CommandName].ModuleName
                    $File = ''
                    $LastWriteTime = $null
                    $OnlineURI = ''
                    $Category = $Work.Manifests[$CommandName].Category
                    $Synopsis = '...'
                    $CommonParameters = $false
                    $CmdInfo = Get-Command -Name $CommandName -ErrorAction SilentlyContinue
                    if ($null -ne $CmdInfo)
                    {
                        $OnlineURI = $CmdInfo.HelpUri
                        if ($null -ne (Get-Member -InputObject $CmdInfo -Name ResolvedCommandName))
                        {
                            $Synopsis = $CmdInfo.ResolvedCommandName
                        }
                        else
                        {
                            $Synopsis = ""
                        }
                        $CommonParameters = $CmdInfo.Definition.IndexOf('<CommonParameters>') -ne -1
                        if ($ModuleName -eq '')
                        {
                            $ModuleName = $CmdInfo.ModuleName
                        }
                    }
                    AddItem $false @{Name = $CommandName;
                                     ModuleName = $ModuleName;
                                     File = $File;
                                     LastWriteTime = $LastWriteTime;
                                     OnlineURI = $OnlineURI;
                                     Format = '';
                                     Index = -1;
                                     Category = $Category;
                                     Component = '';
                                     Functionality = '';
                                     Role = '';
                                     Synopsis = $Synopsis;
                                     Aliases = @();
                                     CommonParameters = $CommonParameters}
                }
            }

            #----------------------------------------------------------
            # Add help items for aliases

            Get-ChildItem alias: |
                ForEach-Object `
                {
                    # Alias  $_.Name  ->  $_.Definition
                    #
                    if ($null -ne $HelpInfo.ItemIndex[$_.Definition])
                    {
                        # Aliases for which we have already regular items
                        $Item = $HelpInfo.Items[$HelpInfo.ItemIndex[$_.Definition]]
                        $Item.Aliases += $_.Name;
                        AddItem $true @{Name = $_.Name;
                                        ModuleName = $Item.ModuleName;
                                        File = $Item.File;
                                        LastWriteTime = $Item.LastWriteTime;
                                        OnlineURI = $Item.OnlineURI;
                                        Format = $Item.Format;
                                        Index = $Item.Index;
                                        Category = 'Alias';
                                        Component = $Item.Component;
                                        Functionality = $Item.Functionality;
                                        Role = $Item.Role;
                                        Synopsis = $_.Definition;
                                        Aliases = $Item.Aliases;
                                        CommonParameters = $Item.CommonParameters}
                    }
                    else
                    {
                        # Other aliases
                        AddItem $true @{Name = $_.Name;
                                        ModuleName = '';
                                        File = '';
                                        LastWriteTime = $null;
                                        OnlineURI = '';
                                        Format = '';
                                        Index = -1;
                                        Category = 'Alias';
                                        Component = '';
                                        Functionality = '';
                                        Role = '';
                                        Synopsis = $_.Definition;
                                        Aliases = @();
                                        CommonParameters = $false}
                    }
                }
            
            #----------------------------------------------------------
            # Extract comment-based help for function definition
            foreach ($CommandName in $Work.Manifests.keys)
            {
                if (-not $Work.Manifests[$CommandName].Used -and
                    $Work.Manifests[$CommandName].Category -eq 'Alias')
                {
                    $ModuleName = $Work.Manifests[$CommandName].ModuleName
                    $File = ''
                    $LastWriteTime = $null
                    $OnlineURI = ''
                    $Category = $Work.Manifests[$CommandName].Category
                    $Synopsis = '...'
                    $CommonParameters = $false
                    $CmdInfo = Get-Command -Name $CommandName -ErrorAction SilentlyContinue
                    if ($null -ne $CmdInfo)
                    {
                        $OnlineURI = $CmdInfo.HelpUri
                        if ($null -ne (Get-Member -InputObject $CmdInfo -Name ResolvedCommandName))
                        {
                            $Synopsis = $CmdInfo.ResolvedCommandName
                        }
                        else
                        {
                            $Synopsis = ""
                        }
                        $CommonParameters = $CmdInfo.Definition.IndexOf('<CommonParameters>') -ne -1
                        # Aliases for which we have already regular items
                        $Item = $HelpInfo.Items[$HelpInfo.ItemIndex[$Synopsis]]
                        if (-not $CommandName -in $Item.Aliases)
                        {
                            $Item.Aliases += $CommandName
                        }
                        if ($ModuleName -eq '')
                        {
                            $ModuleName = $Item.ModuleName
                        }
                        $File = $Item.File
                        $LastWriteTime = $Item.LastWriteTime
                    }
                    AddItem $false @{Name = $CommandName;
                                     ModuleName = $ModuleName;
                                     File = $File;
                                     LastWriteTime = $LastWriteTime;
                                     OnlineURI = $OnlineURI;
                                     Format = '';
                                     Index = -1;
                                     Category = $Category;
                                     Component = '';
                                     Functionality = '';
                                     Role = '';
                                     Synopsis = $Synopsis;
                                     Aliases = @();
                                     CommonParameters = $CommonParameters}
                }
            }
        }   # function FindHelpFiles #
            ##########################


        #==========================================================
        # Displaying stuff
        #==========================================================

        ###############################
        # function DisplayXmlHelpFile #
        ###############################
        function DisplayXmlHelpFile
        {
            param (
                [System.Collections.Hashtable] $Item,
                [System.Xml.XmlElement] $CommandNode
            )


            ###########################
            # function BuildLinkValue #
            ###########################
            function BuildLinkValue
            {
                param (
                    [System.String] $LinkText,
                    [System.String] $URI
                )

                ###########################
                # function BuildLinkValue #

                if ($URI -ne '')
                {
                    if (($URI.Substring(0,8) -ne 'https://') -and ($URI.Substring(0,7) -ne 'http://'))
                    {
                        Write-Error "Unknown link for ${LinkText}: $URI"
                    }
                }
                else
                {
                    $version = '{0}.{1}' -f $PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor
                    if ($null -ne $HelpInfo.ItemIndex[$LinkText])
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
                    if (($LinkText.Length -gt 0) -and ($LinkText.Substring($LinkText.Length-1,1) -eq ':'))
                    {
                        $LinkText = $LinkText.Substring(0,$LinkText.Length-1)
                    }
                    $LinkValue = $LinkText+': '+$URI
                }
                return $LinkValue
            }   # function BuildLinkValue #
                ###########################


            #############################
            # function DisplayParagraph #
            #############################
            function DisplayParagraph
            {
                param (
                    [System.Int32] $IndentLevel,

                    # All regular paragraphs displays empty line after it.
                    # All compact paragraphs do not displays empty line after it.
                    #
                    # Allowed values for $ParagraphType:
                    #
                    ### Regular paragraphs:
                    # 'regular' - regular paragraph
                    # 'hanging' - regular paragraph with all except first line indented one more level
                    # 'code' - multiline paragraph containing code (formatting supressed) except $Work.Colors.Code
                    #
                    ### Compact paragraphs:
                    # 'empty' - empty line
                    # 'compact' - compact paragraph, like regular but without empty line after it
                    # 'bulletedlist' - compact paragraph being element of bulleted list, if wrapped then 
                    #                  next lines indented 2 characters
                    # 'numberedlist'
                    # 'formatted' - multiline compact paragraph containing already formatted text
                    # 'section' - compact paragraph coloured with $Work.Colors.Section
                    # 'subsection' - compact paragraph coloured with $Work.Colors.SubSection
                    # 'extrasection' - compact paragraph coloured with $Work.Colors.ExtraSection
                    [System.String] $ParagraphType,
                    [System.String] $Text = '',
                    [System.String] $DisplayedLinesVar = ''
                )

                #############################
                # function DisplayParagraph #

                if ($Text.Trim() -eq '')
                {
                    $ParagraphType = 'empty'
                }
                if (($ParagraphType -ne 'regular') -and
                    ($ParagraphType -ne 'hanging') -and
                    ($ParagraphType -ne 'code') -and
                    ($ParagraphType -ne 'empty') -and
                    ($ParagraphType -ne 'compact') -and
                    ($ParagraphType -ne 'bulletedlist') -and
                    ($ParagraphType -ne 'numberedlist') -and
                    ($ParagraphType -ne 'formatted') -and
                    ($ParagraphType -ne 'section') -and
                    ($ParagraphType -ne 'subsection') -and
                    ($ParagraphType -ne 'extrasection'))
                {
                    Write-Error "Unknown paragraph type: $_, fallback into regular"
                    $ParagraphType = 'regular'
                }
                $IndentWidth = $Work.IndentSize * $IndentLevel
                $TextWidth = $Work.OutputWidth-$IndentWidth
                $DisplayedLines = 0
                switch -wildcard ($ParagraphType)
                {
                    'empty'
                        {
                            Write-Output ''
                            $DisplayedLines++
                            break
                        }
                    'code'
                        {
                            $LastLineWasEmpty = $false
                            $Lines = $Text.Split("`n")
                            foreach ($Line in $Lines)
                            {
                                # Drop empty line if previous was also ampty
                                if ($LastLineWasEmpty -and ($Line.Trim().Length -eq 0))
                                {
                                    continue
                                }
                                # Line containg sentence with one dot char at the end is regular paragraph
                                if (($Item.Format -eq 'txt') -or
                                    ($Line.Length -lt 20) -or
                                    ($Line.Substring($Line.Length-1,1) -ne '.') -or
                                    ($Line.Substring($Line.Length-3,3) -eq '...') -or
                                    ($Line.Substring(0,7) -eq 'PS C:\>') -or
                                    ($Line.Substring(0,1) -eq '#') -or
                                    ($CommandNode.details.name -eq 'Get-Help'))
                                {
                                    if ($Line.Length -gt $TextWidth)
                                    {
                                        $Line = $Line.Substring(0, $TextWidth-3)+'...'
                                    }
                                    Write-Output ((' ' * $IndentWidth)+$Work.Colors.Code+$Line+$Work.Colors.Default)
                                    $DisplayedLines++
                                    $LastLineWasEmpty = ($Line.Trim().Length -eq 0)
                                }
                                else
                                {
                                    DisplayParagraph $IndentLevel 'regular' $Line 'D'
                                    $DisplayedLines += $D
                                    $LastLineWasEmpty = $true
                                }
                            }
                            if (-not $LastLineWasEmpty)
                            {
                                Write-Output ''
                                $DisplayedLines++
                            }
                            break
                        }
                    'formatted'
                        {
                            $Lines = $Text.Split("`n")
                            foreach ($Line in $Lines)
                            {
                                if ($Line.Length -gt $TextWidth)
                                {
                                    $Line = $Line.Substring(0, $TextWidth-3)+'...'
                                }
                                Write-Output ((' ' * $IndentWidth)+$Line)
                                $DisplayedLines++
                            }
                            break
                        }
                    'regular'
                        {
                            $Compact = $false
                            $HangSize = 0
                            $StartColor = ''
                            $EndColor = ''
                        }
                    'hanging'
                        {
                            $Compact = $false
                            $HangSize = $Work.IndentSize
                            $StartColor = ''
                            $EndColor = ''
                        }
                    'compact'
                        {
                            $Compact = $true
                            $HangSize = 0
                            $StartColor = ''
                            $EndColor = ''
                        }
                    'bulletedlist'
                        {
                            $Compact = $true
                            $HangSize = 2
                            $StartColor = ''
                            $EndColor = ''
                        }
                    'numberedlist'
                        {
                            $Compact = $true
                            $HangSize = $Work.IndentSize
                            $StartColor = ''
                            $EndColor = ''
                        }
                    'section'
                        {
                            $Compact = $true
                            $HangSize = 0
                            $StartColor = $Work.Colors.Section
                            $EndColor = $Work.Colors.Default
                        }
                    'subsection'
                        {
                            $Compact = $true
                            $HangSize = 0
                            $StartColor = $Work.Colors.SubSection
                            $EndColor = $Work.Colors.Default
                        }
                    'extrasection'
                        {
                            $Compact = $true
                            $HangSize = 0
                            $StartColor = $Work.Colors.ExtraSection
                            $EndColor = $Work.Colors.Default
                        }
                    '*'
                        {
                            $Text = $Text.Trim()
                            # Replace inside text all sequences of more spaces into one space
                            while ($Text.IndexOf('  ') -ne -1)
                            {
                                $Text = $Text.Replace('  ', ' ')
                            }
                            # Remove all spaces prepending dot character
                            # Should not remove space before '.NET' !!!???
                            while ($Text.IndexOf(' .') -ne -1)
                            {
                                $Text = $Text.Replace(' .', '.')
                            }
                            #----------------------------------------------------------
                            # Break paragraph text into lines not longer than $TextWidth
                            $Lines = @()
                            while ($Text.Length -gt $TextWidth)
                            {
                                $Line = $Text.Substring(0, $TextWidth)
                                # Move last broken (or glued to the end of $Line) word to $NextLine
                                $NextLine = ''
                                if ($Line.IndexOf(' ') -ne -1)
                                {
                                    while (($Line.Length -gt 0) -and ($Line.Substring($Line.Length-1, 1) -ne ' '))
                                    {
                                        $NextLine = $NextLine.Insert(0, $Line.Substring($Line.Length-1, 1))
                                        $Line = $Line.Substring(0, $Line.Length-1)
                                    }
                                }
                                # Store current line in $Lines
                                $Lines += @($Line.TrimEnd())
                                # Use $NextLine and rest of the $Text as a new $Text
                                $Text = $NextLine+$Text.Substring($TextWidth)
                                #
                                if ($Lines.Count -eq 1)
                                {
                                    $TextWidth -= $HangSize
                                }
                            }
                            $Lines += @($Text)
                            #----------------------------------------------------------
                            # Display all paragraph lines
                            foreach ($Line in $Lines)
                            {
                                Write-Output ((' ' * $IndentWidth)+$StartColor+$Line+$EndColor)
                                $DisplayedLines++
                                if ($DisplayedLines -eq 1)
                                {
                                    $IndentWidth += $HangSize
                                }
                            }
                            #----------------------------------------------------------
                            # Display extra empty line for regular paragraphs
                            if (-not $Compact)
                            {
                                Write-Output ''
                                $DisplayedLines++
                            }
                        }
                }
                if ($DisplayedLinesVar -ne '')
                {
                    Set-Variable -Scope 1 -Name $DisplayedLinesVar -Value $DisplayedLines
                }
            }   # function DisplayParagraph #
                #############################


            ##########################################
            # function DisplayCollectionOfParagraphs #
            ##########################################
            function DisplayCollectionOfParagraphs
            {
                param (
                    [System.Int32] $IndentLevel,
                    $Collection,
                    [System.String] $DisplayedLinesVar = ''
                )


                ###########################################
                # function DisplayParagraphFromCollection #
                ###########################################
                function DisplayParagraphFromCollection
                {
                    param (
                        [System.Int32] $IndentLevel,
                        [System.String] $Paragraph,
                        [System.String] $DisplayedLinesVar = ''
                    )

                    ###########################################
                    # function DisplayParagraphFromCollection #
                    
                    $DisplayedLines = 0
                    $Paragraph = $Paragraph.TrimEnd(" `n")
                    $Paragraph = $Paragraph.TrimStart()
                    if ($Paragraph.Length -eq 0)
                    {
                        Set-Variable -Scope 1 -Name $DisplayedLinesVar -Value $DisplayedLines
                        return
                    }
                    if ($Paragraph.IndexOf("`n") -ne -1)
                    {
                        # Instead of $Work.WasColon there should be used info about colon in last paragraph
                        DisplayCollectionOfParagraphs $IndentLevel ($Paragraph.Split("`n")) 'Displayed'
                        $DisplayedLines += $Displayed
                    }
                    else
                    {
                        if (-not $Work.WasColon)
                        {
                            #----------------------------------------------------------
                            # Previous paragraph/line was not the list heading paragraph
                            if ($Paragraph.Substring($Paragraph.Length-1, 1) -eq ':')
                            {
                                # List heading paragraph
                                $Work.WasColon = $true
                                DisplayParagraph $IndentLevel 'compact' $Paragraph 'Displayed'
                                $DisplayedLines += $Displayed
                            }
                            else
                            {
                                # Regular paragraph
                                DisplayParagraph $IndentLevel 'regular' $Paragraph 'Displayed'
                                $DisplayedLines += $Displayed
                            }
                        }
                        else
                        {
                            #----------------------------------------------------------
                            # Previous paragraph/line was the list heading paragraph or the list item
                            if ($Paragraph -match '^[0-9]+ =')
                            {
                                # Numbered list item
                                DisplayParagraph $IndentLevel 'numberedlist' $Paragraph 'Displayed'
                                $DisplayedLines += $Displayed
                            }
                            elseif (($Paragraph.Length -gt 2) -and
                                    ($Paragraph.Substring(0, 2) -eq '- '))
                            {
                                # Bulleted list item
                                DisplayParagraph $IndentLevel 'bulletedlist' $Paragraph 'Displayed'
                                $DisplayedLines += $Displayed
                            }
                            elseif (($Paragraph.Length -gt 3) -and
                                    ($Paragraph.Substring(0, 3) -eq '-- '))
                            {
                                # Bulleted list item
                                DisplayParagraph $IndentLevel 'bulletedlist' ('- '+$Paragraph.Substring(3)) 'Displayed'
                                $DisplayedLines += $Displayed
                            }
                            elseif (($Paragraph.Length -gt 2) -and
                                    ($Paragraph.Substring(0, 2) -eq '--'))
                            {
                                # Bulleted list item
                                DisplayParagraph $IndentLevel 'bulletedlist' ('- '+$Paragraph.Substring(2)) 'Displayed'
                                $DisplayedLines += $Displayed
                            }
                            else
                            {
                                # End of the list
                                DisplayParagraph $IndentLevel 'empty'
                                $DisplayedLines += 1
                                $Work.WasColon = $false
                                # But it could be begin of the next list
                                if ($Paragraph.Substring($Paragraph.Length-1, 1) -eq ':')
                                {
                                    # List heading paragraph
                                    $Work.WasColon = $true
                                    DisplayParagraph $IndentLevel 'compact' $Paragraph 'Displayed'
                                    $DisplayedLines += $Displayed
                                }
                                else
                                {
                                    # Regular paragraph
                                    DisplayParagraph $IndentLevel 'regular' $Paragraph 'Displayed'
                                    $DisplayedLines += $Displayed
                                }
                            }
                        }
                    }
                    if ($DisplayedLinesVar -ne '')
                    {
                        Set-Variable -Scope 1 -Name $DisplayedLinesVar -Value $DisplayedLines
                    }
                }   # function DisplayParagraphFromCollection #
                    ###########################################

                ##########################################
                # function DisplayCollectionOfParagraphs #

                $DisplayedLines = 0
                switch ($Collection.GetType().FullName)
                {
                    'System.Xml.XmlElement'
                        {
                            foreach ($Child in $Collection.ChildNodes)
                            {
                                switch ($Child.LocalName)
                                {
                                    'para'
                                        {
                                            $Paragraph = $Child.InnerText
                                            DisplayParagraphFromCollection $IndentLevel $Paragraph 'Displayed'
                                            $DisplayedLines += $Displayed
                                        }
                                    'code'
                                        {
                                            if ($Work.WasColon)
                                            {
                                                DisplayParagraph $IndentLevel 'empty'
                                                $DisplayedLines += 1
                                                $Work.WasColon = $false
                                            }
                                            DisplayParagraph $IndentLevel 'code' $Child.InnerText 'Displayed'
                                            $DisplayedLines += $Displayed
                                        }
                                }
                            }
                        }
                    {($_ -eq 'System.Object[]') -or
                     ($_ -eq 'System.String[]')}
                        {
                            foreach ($Paragraph in $Collection)
                            {
                                DisplayParagraphFromCollection $IndentLevel $Paragraph 'Displayed'
                                $DisplayedLines += $Displayed
                            }
                        }
                    default
                        {
                        }
                }
                if ($Work.WasColon)
                {
                    DisplayParagraph $IndentLevel 'empty'
                    $DisplayedLines += 1
                }
                if ($DisplayedLinesVar -ne '')
                {
                    Set-Variable -Scope 1 -Name $DisplayedLinesVar -Value $DisplayedLines
                }
            }   # function DisplayCollectionOfParagraphs #
                ##########################################


            ################################
            # function DisplaySingleSyntax #
            ################################
            function DisplaySingleSyntax
            {
                param (
                    [System.Xml.XmlElement] $syntaxItem,
                    [System.Boolean] $CommonParameters
                )

                ################################
                # function DisplaySingleSyntax #

                $Paragraph = $syntaxItem.name
                #----------------------------------------------------------
                # Regular parameters
                if ($null -ne (Get-Member -InputObject $syntaxItem -Name parameter))
                {
                    foreach ($ParamNode in $syntaxItem.parameter)
                    {
                        $Required = $ParamNode.required -eq 'true'
                        $Position = $ParamNode.position -ne 'named'
                        $Paragraph += ' '
                        if (-not $Required)
                        {
                            $Paragraph += '['
                        }
                        if ($Position)
                        {
                            $Paragraph += '['
                        }
                        $Paragraph += '-'+$ParamNode.name
                        if ($Position)
                        {
                            $Paragraph += ']'
                        }
                        $TypeName = ''
                        if ($null -ne (Get-member -InputObject $ParamNode -Name parameterValueGroup))
                        {
                            if ($ParamNode.parameterValueGroup.parameterValue.Count -gt 0)
                            {
                                foreach ($Value in $ParamNode.parameterValueGroup.parameterValue)
                                {
                                    if ($TypeName -eq '')
                                    {
                                        $TypeName = '{'+$Value.InnerText
                                    }
                                    else
                                    {
                                        $TypeName += ' | '+$Value.InnerText
                                    }
                                }
                                $TypeName += '}'
                            }
                        }
                        if ($null -ne (Get-member -InputObject $ParamNode -Name parameterValue))
                        {
                            if (($TypeName -eq '') -and ($null -ne $ParamNode.parameterValue.FirstChild.InnerText))
                            {
                                $TypeName = '<'+$ParamNode.parameterValue.FirstChild.InnerText+'>'
                            }
                        }
                        if (($TypeName -eq '') -and ($null -ne (Get-Member -InputObject $ParamNode -Name type)))
                        {
                            $TypeName = '<'+$ParamNode.type.name+'>'
                        }
                        if (($TypeName -ne '<System.Management.Automation.SwitchParameter>') -and
                            ($TypeName -ne '<SwitchParameter>') -and
                            ($TypeName -ne '') -and ($null -ne $TypeName))
                        {
                            $Paragraph += ' '+$TypeName
                        }
                        if (-not $Required)
                        {
                            $Paragraph += ']'
                        }
                    }
                }
                #----------------------------------------------------------
                # Common parameters
                if ($CommonParameters)
                {
                    $Paragraph += ' [<CommonParameters>]'
                }
                DisplayParagraph 1 'hanging' $Paragraph
            }   # function DisplaySingleSyntax #
                ################################


            ###################################
            # function DisplaySingleParameter #
            ###################################
            function DisplaySingleParameter
            {
                param (
                    [System.Xml.XmlElement] $ParamNode
                )


                ####################################
                # function NormalizeAttributeValue #
                ####################################
                function NormalizeAttributeValue
                {
                    param (
                        [System.String] $Value
                    )
                
                    ####################################
                    # function NormalizeAttributeValue #
                    switch -regex ($Value.Trim()){
                        '^True$'
                            {
                                'True'
                            }
                        '^False$'
                            {
                                'False'
                            }
                        '^None$'
                            {
                                'None'
                            }
                        '^Named$'
                            {
                                'Named'
                            }
                        '^True *\(?ByPropertyName\)?$'
                            {
                                'True (ByPropertyName)'
                            }
                        '^True *\(?ByValue\)?$'
                            {
                                'True (ByValue)'
                            }
                        '^True *\(?ByValue, *ByPropertyName\)?$'
                            {
                                'True (ByValue, ByPropertyName)'
                            }
                        '^True *\(?ByPropertyName, *ByValue\)?$'
                            {
                                'True (ByValue, ByPropertyName)'
                            }
                        default
                            {
                                $Value
                            }
                    }
                }   # function NormalizeAttributeValue #
                    ####################################


                ###################################
                # function DisplaySingleParameter #

                $Required = NormalizeAttributeValue $ParamNode.required
                $Position = NormalizeAttributeValue $ParamNode.position
                $DefVal = NormalizeAttributeValue $ParamNode.defaultValue
                $PipelineInput = NormalizeAttributeValue $ParamNode.pipelineInput
                #Parameter set name           (All)
                $Globbing = NormalizeAttributeValue $ParamNode.globbing
                if ($null -ne (Get-Member -InputObject $ParamNode -Name aliases))
                {
                    $Aliases = $ParamNode.aliases
                }
                else
                {
                    $Aliases = ''
                }
                if (($Aliases -eq '') -or ($Aliases -eq 'none'))
                {
                    $Aliases = 'None'
                }
                #Dynamic?                     false

                DisplayParagraph 1 'subsection' ('-'+$ParamNode.name+' <'+$ParamNode.type.name+'>')

                $Work.WasColon = $false
                DisplayCollectionOfParagraphs 2 $ParamNode.Description

                DisplayParagraph 2 'formatted' "Required?                    $Required"
                DisplayParagraph 2 'formatted' "Position?                    $Position"
                DisplayParagraph 2 'formatted' "Default value                $DefVal"
                DisplayParagraph 2 'formatted' "Accept pipeline input?       $PipelineInput"
                DisplayParagraph 2 'formatted' "Accept wildcard characters?  $Globbing"
                DisplayParagraph 2 'formatted' "Aliases                      $Aliases"
                DisplayParagraph 0 'empty'
            }   # function DisplaySingleParameter #
                ###################################


            #################################
            # function DisplaySingleExample #
            #################################
            function DisplaySingleExample
            {
                param (
                    [System.Xml.XmlElement] $Example
                )

                #################################
                # function DisplaySingleExample #

                DisplayParagraph 1 'subsection' $Example.title.Trim('- ')
                #DisplayParagraph 2 'code' $Example.code
                $Work.WasColon = $false
                DisplayCollectionOfParagraphs 2 $Example
                if ($null -ne (Get-Member -InputObject $Example -Name remarks))
                {
                    $Work.WasColon = $false
                    DisplayCollectionOfParagraphs 2 $Example.remarks
                }
            }   # function DisplaySingleExample #
                #################################


            ###############################
            # function DisplayXmlHelpFile #

            # DisplayParagraph 0 'empty'

            $Item.CurrentExtraSectionNode = $null

            #----------------------------------------------------------
            # Section NAME
            if ($null -ne (Get-Member -InputObject $CommandNode -Name details))
            {
                DisplayParagraph 0 'section' 'NAME'
                DisplayParagraph 1 'regular' $CommandNode.details.name
            }
            else
            {
                Write-Error ("Missing details node for help for command " + $Item.Name)
            }

            #----------------------------------------------------------
            # Section SYNOPSIS
            if ($null -ne (Get-Member -InputObject $CommandNode -Name details))
            {
                DisplayParagraph 0 'section' 'SYNOPSIS'
                DisplayParagraph 1 'regular' $CommandNode.details.description.para
            }
            else
            {
                Write-Error ("Missing details node for help for command " + $Item.Name)
            }

            #----------------------------------------------------------
            # Section SYNTAX
            $Item.CommonParameters = [System.Boolean]$Item.CommonParameters # Should be set correctly during scan !!!
            if (($null -ne (Get-Member -InputObject $CommandNode -Name syntax)) -and
                ($null -ne (Get-Member -InputObject $CommandNode.syntax -Name syntaxItem)))
            {
                DisplayParagraph 0 'section' 'SYNTAX'

                if ($CommandNode.syntax.syntaxItem.Count -ne 0)
                {
                    foreach ($syntaxItem in $CommandNode.syntax.syntaxItem)
                    {
                        DisplaySingleSyntax $syntaxItem $Item.CommonParameters
                    }
                }
                else
                {
                    DisplaySingleSyntax $CommandNode.syntax.syntaxItem $Item.CommonParameters
                }
            }

            #----------------------------------------------------------
            # Section ALIASES
            if (($null -ne $Item.Aliases) -and ($Item.Aliases.Count -gt 0))
            {
                DisplayParagraph 0 'section' 'ALIASES'
                $Paragraph = ''
                foreach ($Alias in $Item.Aliases)
                {
                    if ($Paragraph -eq '')
                    {
                        $Paragraph = $Alias
                    }
                    else
                    {
                        $Paragraph += ", $Alias"
                    }
                }
                DisplayParagraph 1 'regular' $Paragraph
            }

            #----------------------------------------------------------
            # Section DESCRIPTION
            if ($null -ne (Get-Member -InputObject $CommandNode -Name description))
            {
                DisplayParagraph 0 'section' 'DESCRIPTION'
                $Work.WasColon = $false
                DisplayCollectionOfParagraphs 1 $CommandNode.description

                #----------------------------------------------------------
                # Display extra sections
                foreach ($Child in $CommandNode.description.ChildNodes)
                {
                    if ($Child.LocalName -eq 'section')
                    {
                        DisplayParagraph 0 'extrasection' $Child.FirstChild.InnerText
                        $Work.WasColon = $false
                        DisplayCollectionOfParagraphs 1 $Child
                    }
                }
            }

            #----------------------------------------------------------
            # Section PARAMETERS
            $HeaderDisplayed = $false
            if (($null -ne (Get-Member -InputObject $CommandNode -Name parameters)) -and
                ($null -ne (Get-Member -InputObject $CommandNode.parameters -Name parameter)))
            {
                DisplayParagraph 0 'section' 'PARAMETERS'
                $HeaderDisplayed = $true
                if ($CommandNode.parameters.parameter.Count -ne 0)
                {
                    foreach ($ParamNode in $CommandNode.parameters.parameter)
                    {
                        DisplaySingleParameter $ParamNode
                    }
                }
                else
                {
                    DisplaySingleParameter $CommandNode.parameters.parameter
                }
            }
            if ($Item.CommonParameters)
            {
                if (-not $HeaderDisplayed)
                {
                    DisplayParagraph 0 'section' 'PARAMETERS'
                }
                DisplayParagraph 1 'subsection' ('<CommonParameters>')

                DisplayParagraph 2 'regular' ('This cmdlet supports the common parameters: '+
                    'Verbose, Debug, ErrorAction, ErrorVariable, WarningAction, '+
                    'WarningVariable, OutBuffer, PipelineVariable, and OutVariable. '+
                    'For more information, see about_CommonParameters '+
                    '(https://go.microsoft.com/fwlink/?LinkID=113216).')
            }

            #----------------------------------------------------------
            # Section INPUTS
            if (($null -ne (Get-Member -InputObject $CommandNode -Name InputTypes)) -and
                ($null -ne (Get-Member -InputObject $CommandNode.InputTypes -Name InputType)))
            {
                $HeaderDisplayed = $false
                foreach ($InputType in $CommandNode.InputTypes.InputType)
                {
                    $DisplayedLines = 0
                    if ('' -ne $InputType.type.name)
                    {
                        if (-not $HeaderDisplayed)
                        {
                            DisplayParagraph 0 'section' 'INPUTS'
                            $HeaderDisplayed = $true
                        }
                        DisplayParagraph 1 'compact' $InputType.type.name
                    }
                    # if (($null -ne $InputType.description.para) -and
                    #     ($null -ne $InputType.description.para[0]))
                    if ($null -ne (Get-Member -InputObject $InputType.description -Name para))
                    {
                        if (-not $HeaderDisplayed)
                        {
                            DisplayParagraph 0 'section' 'INPUTS'
                            $HeaderDisplayed = $true
                        }
                        $Work.WasColon = $false
                        DisplayCollectionOfParagraphs 2 $InputType.description 'DisplayedLines'
                    }
                    if ($HeaderDisplayed -and ($DisplayedLines -eq 0))
                    {
                        DisplayParagraph 2 'empty'
                    }
                }
            }

            #----------------------------------------------------------
            # Section OUTPUTS
            if (($null -ne (Get-Member -InputObject $CommandNode -Name returnValues)) -and
                ($null -ne (Get-Member -InputObject $CommandNode.returnValues -Name returnValue)))
            {
                $HeaderDisplayed = $false
                foreach ($returnValue in $CommandNode.returnValues.returnValue)
                {
                    $DisplayedLines = 0
                    if ($returnValue.type.name -ne '')
                    {
                        if (-not $HeaderDisplayed)
                        {
                            DisplayParagraph 0 'section' 'OUTPUTS'
                            $HeaderDisplayed = $true
                        }
                        DisplayParagraph 1 'compact' $returnValue.type.name
                    }
                    # if (($null -ne $returnValue.description.para) -and
                    #     ($null -ne $returnValue.description.para[0]))
                    if ($null -ne (Get-Member -InputObject $returnValue.description -Name para))
                    {
                        if (-not $HeaderDisplayed)
                        {
                            DisplayParagraph 0 'section' 'OUTPUTS'
                            $HeaderDisplayed = $true
                        }
                        $Work.WasColon = $false
                        DisplayCollectionOfParagraphs 2 $returnValue.description 'DisplayedLines'
                    }
                    if ($HeaderDisplayed -and ($DisplayedLines -eq 0))
                    {
                        DisplayParagraph 2 'empty'
                    }
                }
            }

            #----------------------------------------------------------
            # Section NOTES
            if (($null -ne (Get-Member -InputObject $CommandNode -Name alertSet)) -and
                ($null -ne (Get-Member -InputObject $CommandNode.alertSet -Name alert)) -and
                ($null -ne (Get-Member -InputObject $CommandNode.alertSet.alert -Name para)))
            {
                DisplayParagraph 0 'section' 'NOTES'
                foreach ($alert in $CommandNode.alertSet.alert)
                {
                    $Work.WasColon = $false
                    DisplayCollectionOfParagraphs 1 $alert
                }
            }

            #----------------------------------------------------------
            # Section EXAMPLES
            if (($null -ne (Get-member -InputObject $CommandNode -Name examples)) -and
                ($null -ne (Get-member -InputObject $CommandNode.examples -Name example)))
            {
                if ('name' -eq $CommandNode.examples.FirstChild.name)
                {
                    DisplayParagraph 0 'section' $CommandNode.examples.FirstChild.InnerText
                }
                else
                {
                    DisplayParagraph 0 'section' 'EXAMPLES'
                }
                if ($CommandNode.examples.example.Count -ne 0)
                {
                    foreach ($Example in $CommandNode.examples.example)
                    {
                        DisplaySingleExample $Example
                    }
                }
                else
                {
                    DisplaySingleExample $CommandNode.examples.example
                }
            }

            #----------------------------------------------------------
            # Section RELATED LINKS
            if (($null -ne (Get-Member -InputObject $CommandNode -Name relatedLinks)) -and
                ($null -ne (Get-Member -InputObject $CommandNode.relatedLinks -Name navigationLink)))
            {
                DisplayParagraph 0 'section' 'RELATED LINKS'
                $Links = @()
                if (($CommandNode.relatedLinks.navigationLink -is [System.Object[]]) -and
                    ($CommandNode.relatedLinks.navigationLink.Count -gt 0))
                {
                    foreach ($NavigationLink in $CommandNode.relatedLinks.navigationLink)
                    {
                        $LinkValue = BuildLinkValue $NavigationLink.linkText $NavigationLink.uri
                        $Links += @('* '+$LinkValue)
                    }
                }
                else
                {
                    $LinkValue = BuildLinkValue $CommandNode.relatedLinks.navigationLink.linkText $CommandNode.relatedLinks.navigationLink.uri
                    $Links += @('* '+$LinkValue)
                }
                $Work.WasColon = $false
                DisplayCollectionOfParagraphs 1 $Links
            }

            #----------------------------------------------------------
            # Section REMARKS
            if ($false)
            {
                DisplayParagraph 0 'section' 'REMARKS'
                DisplayParagraph 1 'regular' 'Remarks will be described later !!!'
            }

        }   # function DisplayXmlHelpFile #
            ###############################


        ############################
        # function DisplayHelpItem #
        ############################
        function DisplayHelpItem
        {
            param (
                [System.Collections.Hashtable] $Item
            )

            ############################
            # function DisplayHelpItem #

            switch ($Item.Format)
            {
                'txt'
                    {
                        $XML = ParseTxtHelpFile $Item
                        Write-Verbose ("Displaying help file " + $HelpInfo.Items[$found[0]].File)
                        Write-Verbose (show-XML.ps1 $XML -Tree ascii -Width ($Work.OutputWidth-5)| Out-String)
                        Write-Output $Work.Colors.Default
                        DisplayXmlHelpFile $Item $XML.ChildNodes[1].ChildNodes[0]
                    }
                'xml'
                    {
                        Write-Verbose "Displaying file: $($Item.File) Item no: $($Item.Index)"
                        $XML = [System.Xml.XmlDocument](Get-Content $Item.File)
                        Write-Output $Work.Colors.Default
                        if ($null -eq $Item.Index)
                        {
                            DisplayXmlHelpFile $Item $XML.helpItems.command
                        }
                        else
                        {
                            DisplayXmlHelpFile $Item ($XML.helpItems.command)[$Item.Index]
                        }
                    }
                default
                    {
                        Write-Verbose "Unknown format in $($Work.ItemsFile), fallback to Get-Help"
                        try
                        {
                            Microsoft.PowerShell.Core\Get-Help -Name $Item.Name -Full
                        }
                        catch
                        {
                            Write-Verbose "Error during: Get-Help -Name $Item.Name -Full"
                        }
                    }
            }
        }   # function DisplayHelpItem #
            ############################

        #==========================================================
        # Main function - Begin block code
        #==========================================================

        # Extract list of all global functions into $Work.Functions
        $Work.Functions = $Functions
        $PSBoundParameters.Remove('Functions') | Out-Null
        #----------------------------------------------------------
        # Ignore HideProgress parameter
        $showProgress = -not $HideProgress
        if ($HideProgress)
        {
            $PSBoundParameters.Remove('HideProgress') | Out-Null
        }
        #----------------------------------------------------------
        # Create profile directory if not yet exits.
        if (-not (Test-Path -Path $Work.ProfileDir))
        {
            New-Item -Path $Work.ProfileDir
        }
        #----------------------------------------------------------
        # If items file exists and we are not required to recreate it,
        # loads its contents.
        if ((Test-Path -Path $Work.ItemsFile) -and -not $Rescan)
        {
            # Import description of all previously found help items
            Write-Verbose ("Importing stored info about help files from $($Work.ItemsFile)")
            $HelpInfo = Import-Clixml -Path $Work.ItemsFile
        }
        else
        {
            # Find all *.help.txt and  *help.xml HelpFiles and parse them
            Write-Verbose 'Find all *.help.txt and *help.xml HelpFiles'
            FindHelpFiles
            # Export description of all found help items
            Write-Verbose "Storing info about help files to $($Work.ItemsFile)"
            Export-Clixml -Path $Work.ItemsFile -Encoding UTF8 -InputObject $HelpInfo
            # Ignore Rescan parameter
            if ($Rescan)
            {
                $PSBoundParameters.Remove('Rescan')
            }
        }

    } # function Main - Begin block #

    #################################
    # function Main - Process block #
    #################################
    Process
    {
        #==========================================================
        # Main function - Process block code
        #==========================================================

        #----------------------------------------------------------
        if ($Online -or $ShowWindow)
        {
            Write-Verbose 'Parameter Online or ShowWindow used, fallback to Get-Help'
            Microsoft.PowerShell.Core\Get-Help @PSBoundParameters
            return
        }
        #----------------------------------------------------------
        # Searching help items to display
        $found = @()
        # Write-Verbose "There are $($HelpInfo.Items.Count) items in our Items File"
        if ($null -eq $Category)
        {
            $Category = @()
        }
        if ($null -eq $Component)
        {
            $Component = @()
        }
        if ($null -eq $Functionality)
        {
            $Functionality = @()
        }
        if ($null -eq $Role)
        {
            $Role = @()
        }
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
        #----------------------------------------------------------
        # Displaying found results
        switch ($found.Count)
        {
            0
                {
                    Write-Verbose "Not found in $($Work.ItemsFile), fallback to Get-Help"
                    # Ignore Detailed parameter
                    if ($Detailed)
                    {
                        $PSBoundParameters.Remove('Detailed')
                    }
                    # Ignore Examples parameter
                    if ($Examples)
                    {
                        $PSBoundParameters.Remove('Examples')
                    }
                    # Ignore Parameter parameter
                    if ($Parameter)
                    {
                        $PSBoundParameters.Remove('Parameter')
                    }
                    # Force Full parameter
                    if (-not $Full)
                    {
                        $PSBoundParameters['Full'] = $true
                    }
                    try
                    {
                        Microsoft.PowerShell.Core\Get-Help @PSBoundParameters
                    }
                    catch
                    {
                        Write-Verbose "Error during: Get-Help @PSBoundParameters"
                    }
                }
            1
                {
                    DisplayHelpItem $HelpInfo.Items[$found[0]]
                }
            default
                {
                    # $Result = @()
                    for ($i = 0; $i -lt $found.Count; $i++)
                    {
                        Write-Output ([pscustomobject]@{Name = $HelpInfo.Items[$found[$i]].Name;
                                                       Category = $HelpInfo.Items[$found[$i]].Category;
                                                       Module = $HelpInfo.Items[$found[$i]].ModuleName;
                                                       Synopsis = $HelpInfo.Items[$found[$i]].Synopsis;
                                                       #File = $HelpInfo.Items[$found[$i]].File;
                                                       #Index = $HelpInfo.Items[$found[$i]].Index
                                                       })
                        # $Result += [pscustomobject]@{Name = $HelpInfo.Items[$found[$i]].Name;
                        #                              Category = $HelpInfo.Items[$found[$i]].Category;
                        #                              Module = $HelpInfo.Items[$found[$i]].ModuleName;
                        #                              Synopsis = $HelpInfo.Items[$found[$i]].Synopsis;
                        #                              #File = $HelpInfo.Items[$found[$i]].File;
                        #                              #Index = $HelpInfo.Items[$found[$i]].Index
                        #                              }
                    }
                    # $Result | Format-Table | Out-String
                }
        }
    } # function Main - Process block #

    #############################
    # function Main - End block #
    #############################
    End
    {
    }
}   # function Main #
    #################

$Functions = @{}
Get-ChildItem -Path function: |
ForEach-Object -Process `
{
    if ($_.Name -ne 'Main')
    {
        $Functions[$_.Name] = 'Function'
    }
}
Main @PSBoundParameters -Functions $Functions
