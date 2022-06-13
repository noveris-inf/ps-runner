<#
#>

########
# Global settings
$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"
Set-StrictMode -Version 2

########
# Add types
Add-Type -AssemblyName 'System.Web'

# List of library items that can be referenced in Add-ReportRunnerSection
$Script:Definitions = New-Object 'System.Collections.Generic.Dictionary[string, ReportRunnerBlock]'

Class ReportRunnerContext
{
    [string]$Title
    [System.Collections.Generic.List[ReportRunnerSection]]$Sections
    [HashTable]$Data

    ReportRunnerContext([string]$title, [HashTable]$data)
    {
        $this.Title = $title
        $this.Data = $data.Clone()
        $this.Sections = New-Object 'System.Collections.Generic.LinkedList[ReportRunnerSection]'
    }
}

Class ReportRunnerSection
{
    [string]$Name
    [string]$Description
    [HashTable]$Data
    [System.Collections.Generic.Dictionary[string, ReportRunnerBlock]]$Blocks
    [Guid]$Guid

    ReportRunnerSection([string]$name, [string]$description, [HashTable]$data)
    {
        $this.Name = $name
        $this.Description = $description
        $this.Data = $data.Clone()
        $this.Blocks = New-Object 'System.Collections.Generic.Dictionary[string, ReportRunnerBlock]'
        $this.Guid = [Guid]::NewGuid()
    }
}

class ReportRunnerBlock
{
    [string]$Id
    [string]$Name
    [string]$Description
    [HashTable]$Data
    [ScriptBlock]$Script
    [System.Collections.Generic.LinkedList[PSObject]]$Content
    [Guid]$Guid

    ReportRunnerBlock([string]$id, [string]$name, [string]$description, [HashTable]$data, [ScriptBlock]$script)
    {
        $this.Id = $id
        $this.Name = $name
        $this.Description = $description
        $this.Content = New-Object 'System.Collections.Generic.LinkedList[PSObject]'
        $this.Data = $data.Clone()
        $this.Script = $script
        $this.Guid = [Guid]::NewGuid()
    }
}

Class ReportRunnerFormatTable
{
    [System.Collections.ArrayList]$Content

    ReportRunnerFormatTable([System.Collections.ArrayList]$content)
    {
        $this.Content = $content
    }
}

enum ReportRunnerEncodeStatus
{
    Ignore = 0
    Encode
    Decode
}

Class ReportRunnerBlockSettings
{
    [ReportRunnerEncodeStatus]$EncodeStatus

    ReportRunnerBlockSettings()
    {
    }

    ReportRunnerBlockSettings([ReportRunnerBlockSettings] $otherSettings)
    {
        $this.Copy($otherSettings)
    }

    Copy([ReportRunnerBlockSettings] $otherSettings)
    {
        if ($null -ne $otherSettings.EncodeStatus)
        {
            $this.EncodeStatus = $otherSettings.EncodeStatus
        }
    }
}


Class ReportRunnerBlockSettingsPop
{
    ReportRunnerBlockSettingsPop()
    {
    }
}

enum ReportRunnerStatus
{
    None = 0
    Info
    Warning
    Error
    InternalError
}

<#
#>
Class ReportRunnerNotice
{
    [ReportRunnerStatus]$Status
    [string]$Description
    [string]$SourceBlock

    ReportRunnerNotice([ReportRunnerStatus]$status, [string]$description)
    {
        $this.Status = $status
        $this.Description = $description
        $this.SourceBlock = $null
    }

    [string] ToString()
    {
        return ("{0}: {1}" -f $this.Status.ToString().ToUpper(), $this.Description)
    }
}

<#
#>
Function New-ReportRunnerContext
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    [OutputType('ReportRunnerContext')]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [HashTable]$Data = @{}
    )

    process
    {
        $obj = New-Object ReportRunnerContext -ArgumentList $Title, $Data

        $obj
    }
}

<#
#>
Function New-ReportRunnerSection
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    [OutputType('ReportRunnerSection')]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ReportRunnerContext]$Context,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Description,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [HashTable]$Data = @{}
    )

    process
    {
        $obj = New-Object ReportRunnerSection -ArgumentList $Name, $Description, $Data

        # Add this new section to the list of sections in the current context
        $Context.Sections.Add($obj)

        # Pass the section on to allow the caller access to the section
        $obj
    }
}

<#
#>
Function New-ReportRunnerBlock
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding(DefaultParameterSetName="NewBlock")]
    [OutputType('ReportRunnerBlock')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName="NewBlock")]
        [Parameter(Mandatory=$true, ParameterSetName="Library")]
        [ValidateNotNullOrEmpty()]
        [ReportRunnerSection]$Section,

        [Parameter(Mandatory=$true, ParameterSetName="Library")]
        [ValidateNotNullOrEmpty()]
        [string]$LibraryFilter,

        [Parameter(Mandatory=$true, ParameterSetName="NewBlock")]
        [ValidatePattern("^[a-zA-Z_-]*\.[a-zA-Z_-]*\.[a-zA-Z_-]*$")]
        [string]$Id,

        [Parameter(Mandatory=$true, ParameterSetName="NewBlock")]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$true, ParameterSetName="NewBlock")]
        [ValidateNotNullOrEmpty()]
        [string]$Description,

        [Parameter(Mandatory=$false, ParameterSetName="Library")]
        [Parameter(Mandatory=$false, ParameterSetName="NewBlock")]
        [ValidateNotNull()]
        [HashTable]$Data = @{},

        [Parameter(Mandatory=$true, ParameterSetName="NewBlock")]
        [ValidateNotNull()]
        [ScriptBlock]$Script
    )

    process
    {
        # Check if we're matching on a LibraryFilter, rather than a new block
        if (![string]::IsNullOrEmpty($LibraryFilter))
        {
            # Find all matches and add each block to the section with the supplied data
            $script:Definitions.Keys | Where-Object { $_ -match $LibraryFilter} | ForEach-Object {
                $lib = $Script:Definitions[$_]

                $obj = New-Object ReportRunnerBlock -ArgumentList $lib.Id, $lib.Name, $lib.Description, $Data, $lib.Script

                $Section.Blocks[$lib.Id] = $obj
            }

            return
        }

        # Create a new block that will be added to the section
        $obj = New-Object ReportRunnerBlock -ArgumentList $Id, $Name, $Description, $Data, $Script

        # Add this new block to the list of blocks in the current section
        $Section.Blocks[$Id] = $obj
    }
}

<#
#>
Function New-ReportRunnerNotice
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Description,

        [Parameter(mandatory=$false)]
        [ValidateNotNull()]
        [ReportRunnerStatus]$Status = [ReportRunnerStatus]::None
    )

    process
    {
        $notice = New-Object ReportRunnerNotice -ArgumentList $Status, $Description

        $notice
    }
}

<#
#>
Function ConvertTo-ReportRunnerFormatTable
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true,ValueFromPipeline)]
        [ValidateNotNull()]
        $Content
    )

    begin
    {
        $objs = New-Object 'System.Collections.ArrayList'
    }

    process
    {
        $objs.Add($Content) | Out-Null
    }

    end
    {
        $format = [ReportRunnerFormatTable]::New($objs)

        $format
    }
}

<#
#>
Function Add-ReportRunnerLibraryBlock
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, HelpMessage = "Must be in module.group.id format")]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern("^[a-zA-Z_-]*\.[a-zA-Z_-]*\.[a-zA-Z_-]*$")]
        [string]$Id,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Description,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ScriptBlock]$Script
    )

    process
    {
        $script:Definitions[$Id] = New-Object ReportRunnerBlock -ArgumentList $Id, $Name, $Description, @{}, $Script
    }
}

<#
#>
Function Invoke-ReportRunnerContext
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ReportRunnerContext]$Context
    )

    process
    {
        $Context.Sections | ForEach-Object {
            $section = $_

            # Flatten the Context and Section data in to a new HashTable
            $sectionData = $Context.Data.Clone()
            $section.Data.Keys | ForEach-Object { $sectionData[$_] = $section.Data[$_] }

            $section.Blocks.Keys | ForEach-Object {
                $blockId = $_
                $block = $section.Blocks[$blockId]

                # Flatten the Section and Block data in to a new HashTable
                $blockData = $sectionData.Clone()
                $block.Data.Keys | ForEach-Object { $blockData[$_] = $block.Data[$_] }

                # Invoke the block script with the relevant data and store content
                $content = New-Object 'System.Collections.Generic.LinkedList[PSObject]'
                Invoke-Command -NoNewScope {
                    # Run the script block
                    try {
                        ForEach-Object -InputObject $blockData -Process $block.Script
                    } catch {
                        New-ReportRunnerNotice -Status InternalError -Description "Error running script: $_"
                    }
                } *>&1 | ForEach-Object {

                    # Add the source block guid, if it is a notice
                    if ([ReportRunnerNotice].IsAssignableFrom($_.GetType()))
                    {
                        [ReportRunnerNotice]$notice = $_
                        $notice.SourceBlock = [string]($block.Guid)
                    }

                    # Add the content to the list for this block
                    $content.Add($_)
                }

                # Save the content back to the block
                $block.Content = $content
            }
        }
    }
}

<#
#>
Function Format-ReportRunnerContextAsHtml
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '')]
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ReportRunnerContext]$Context,

        [Parameter(Mandatory=$false)]
        [switch]$SummaryOnly = $false
    )

    process
    {
        # Collection of all notices across all sections
        $allNotices = [ordered]@{}

        # Html preamble
        $title = $Context.Title
        "<!DOCTYPE html PUBLIC `"-//W3C//DTD XHTML 1.0 Strict//EN`"  `"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd`">"
        "<html xmlns=`"http://www.w3.org/1999/xhtml`">"
        "<head>"
        "<title>$title</title>"
        "<style>"
        "table {"
        "  font-family: Arial, Helvetica, sans-serif;"
        "  border-collapse: collapse;"
        "  width: 100%;"
        "}"
        "td, th {"
        "  border: 1px solid #ddd;"
        "  padding: 6px;"
        "}"
        "tr:nth-child(even){background-color: #f2f2f2;}"
        "tr:hover {background-color: #ddd;}"
        ".warningCell {background-color: #ffeb9c;}"
        ".errorCell {background-color: #ffc7ce;}"
        ".internalErrorCell {background-color: #ffc7ce;}"
        "th {"
        "  padding-top: 12px;"
        "  padding-bottom: 12px;"
        "  text-align: left;"
        "  background-color: #04AA6D;"
        "  color: white;"
        "}"
        "div.section {"
        "  padding: 10px;"
        "  padding-bottom: 20px;"
        "  border: 1px solid gray;"
        "  margin-bottom: 10px;"
        "  box-shadow: 4px 3px 8px 1px #969696"
        "}"
        "div.block {"
        "  border-top: 1px solid gray;"
        "  margin-top: 20px;"
        "}"
        "div.blockContent {"
        "  font-family: Courier New, monospace;"
        "  white-space: pre"
        "}"
        ".row {"
        "  display: flex;"
        "}"
        ".column {"
        "  flex: 50%;"
        "}"
        ".rrformattable {"
        "  white-space: normal;"
        "}"
        "</style>"
        "</head><body>"
        "<div id=`"top`"></div>"
        "<h2>$title</h2>"

        $sectionList = New-Object 'System.Collections.Generic.LinkedList[PSCustomObject]'
        $sectionContent = $Context.Sections | ForEach-Object {
            $section = $_
            $notices = New-Object 'System.Collections.Generic.LinkedList[ReportRunnerNotice]'
            $sectionGuid = $section.Guid

            # Add the section to the section list
            $sectionList.Add([PSCustomObject]@{
                Name = $section.Name
                Id = [string]$sectionGuid
            })

            # Format section start
            "<div class=`"section`" id=`"$sectionGuid`">"
            ("<div class=`"row`"><div class=`"column`"><h3>Section: {0}</h3></div><div class=`"column`" align=`"right`"><a href=`"#top`">Back to top</a></div></div>" -f $section.Name)
            ("<i>{0}</i><br><br>" -f $section.Description)

            # Iterate through block content
            $content = $section.Blocks.Keys | ForEach-Object {
                $blockId = $_
                $block = $section.Blocks[$blockId]
                $blockGuid = $block.Guid

                # Format block start
                "<div class=`"block`" id=`"$blockGuid`">"
                ("<div class=`"row`"><div class=`"column`"><h4>{0} ({1})</h4></div>" -f $block.Name, $block.Id)
                "<div class=`"column`" align=`"right`"><a href=`"#$sectionGuid`">Back to section</a> | <a href=`"#top`">Back to top</a></div></div>"
                ("<i>{0}</i><br><br>" -f $block.Description)

                # Default block settings
                $blockSettings = New-Object 'System.Collections.Generic.List[ReportRunnerBlockSettings]'
                $defaultSetting = [ReportRunnerBlockSettings]::New()
                $defaultSetting.EncodeStatus = [ReportRunnerEncodeStatus]::Ignore
                $blockSettings.Add($defaultSetting)

                # Format block content
                $blockContent = $block.Content | ForEach-Object {
                    $msg = $_

                    # Check if it is a string or status object
                    switch ($msg.GetType().FullName)
                    {
                        "ReportRunnerBlockSettings"
                        {
                            # Create a new settings object based on the current one
                            $newSettings = [ReportRunnerBlockSettings]::New($blockSettings[0])
                            $newSettings.Copy([ReportRunnerBlockSettings]$msg)
                            $blockSettings.Insert(0, $newSettings)
                        }

                        "ReportRunnerBlockSettingsPop"
                        {
                            # Remove settings, but don't remove the last one
                            if ($blockSettings.Count -gt 1)
                            {
                                $blockSettings.RemoveAt(0)
                            } else {
                                Write-Warning "Settings pop, but no settings to pop"
                            }
                        }

                        "ReportRunnerNotice" {
                            [ReportRunnerNotice]$notice = $msg
                            $notices.Add($notice) | Out-Null

                            if ($allNotices.Keys -notcontains $section.Name)
                            {
                                $allNotices[$section.Name] = New-Object 'System.Collections.Generic.LinkedList[ReportRunnerNotice]'
                            }

                            $allNotices[$section.Name].Add($notice) | Out-Null

                            $notice.ToString()
                        }

                        "System.Management.Automation.InformationRecord" {
                            ("INFO: {0}" -f $msg.ToString())
                        }

                        "System.Management.Automation.VerboseRecord" {
                            ("VERBOSE: {0}" -f $msg.ToString())
                        }

                        "System.Management.Automation.ErrorRecord" {
                            ("ERROR: {0}" -f $msg.ToString())
                        }

                        "System.Management.Automation.DebugRecord" {
                            ("DEBUG: {0}" -f $msg.ToString())
                        }

                        "System.Management.Automation.WarningRecord" {
                            ("WARNING: {0}" -f $msg.ToString())
                        }

                        "ReportRunnerFormatTable" {
                            $content = "<div class=`"rrformattable`">"
                            $content += $msg.Content | ConvertTo-Html -As Table -Fragment | ForEach-Object {
                                $_.Replace([Environment]::Newline, "<br>")
                            } | Out-String
                            $content += "</div>"
                            $content = $content.Replace([Environment]::Newline, "")

                            $content
                        }

                        "System.String" {
                            switch ($blockSettings[0].EncodeStatus)
                            {
                                "Decode" {
                                    $msg = [System.Web.HttpUtility]::HtmlDecode($msg)
                                }

                                "Encode" {
                                    $msg = [System.Web.HttpUtility]::HtmlEncode($msg)
                                }
                            }

                            $msg
                        }

                        default {
                            $msg
                        }
                    }
                } | Out-String

                # Replace newlines with breaks and output
                $blockContent = $blockContent.Replace([Environment]::Newline, "<br>")
                "<div class=`"blockContent`">"
                $blockContent
                "</div>"

                # Format block end
                "</div>"
            }

            # Display notices for this section
            if (($notices | Measure-Object).Count -gt 0)
            {
                "<div class=`"notice`">"
                "<h4>Notices</h4>"

                $notices |
                    Sort-Object -Property Status -Descending |
                    Format-ReportRunnerNotice -IncludeLinks |
                    ConvertTo-Html -As Table -Fragment |
                    Update-ReportRunnerNoticeCellClass |
                    Format-ReportRunnerDecodeHtml

                "</div>"
            }

            # Display block content
            $content | Out-String

            # Format section end
            "</div>"
        } | Out-String

        # Display all notices here
        "<div class=`"section`"><div class=`"notice`">"
        "<h3>All Notices</h3>"
        "<i>Notices generated by any section</i>"

        $allNotices.Keys | ForEach-Object {
            $key = $_
            $allNotices[$key] | ForEach-Object {
                [PSCustomObject]@{
                    Section = $key
                    Status = [int]($_.Status)
                    Notice = $_
                }
            }
        } | Sort-Object -Property Status,Section -Descending |
            ForEach-Object {
                $_.Notice | Format-ReportRunnerNotice -SectionName $_.Section -IncludeLinks:(!$SummaryOnly)
            } |
            ConvertTo-Html -As Table -Fragment |
            Update-ReportRunnerNoticeCellClass |
            Format-ReportRunnerDecodeHtml

        "</div></div>"

        # Display body for the report, if required
        if (!$SummaryOnly)
        {
            # Display section table of contents
            "<div class=`"section`">"
            "<h3>Index</h3>"

            $Context.Sections | ForEach-Object {
                $section = $_
                $sectionGuid = $section.Guid

                # Display section heading
                ("<a href=`"#{0}`">{1}</a><br>" -f $sectionGuid, $section.Name)
                "<ul>"

                $section.Blocks.Keys | ForEach-Object {
                    $blockId = $_
                    $block = $section.Blocks[$blockId]
                    $blockGuid = $block.Guid

                    # Display block link
                    ("<a href=`"#{0}`">{1}</a><br>" -f $blockGuid, $block.Name)
                }

                "</ul>"
            }

            "</div>"

            # Display all section content
            $sectionContent | ForEach-Object { $_ }
        }

        # Wrap up HTML
        "</body></html>"
    }
}

Function Update-ReportRunnerNoticeCellClass
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    [OutputType('System.String')]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Content
    )

    process
    {
        $val = $Content

        if ($null -eq $val)
        {
            $val = ""
        }

        $val = $val.Replace("<td>Warning</td>", "<td class=`"warningCell`">Warning</td>")
        $val = $val.Replace("<td>Error</td>", "<td class=`"errorCell`">Error</td>")
        $val = $val.Replace("<td>InternalError</td>", "<td class=`"internalErrorCell`">InternalError</td>")

        $val
    }
}

Function Format-ReportRunnerNotice
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [ValidateNotNull()]
        [ReportRunnerNotice]$Notice,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$SectionName,

        [Parameter(Mandatory=$false)]
        [switch]$IncludeLinks = $false
    )

    process
    {
        # Format the description as Html ID reference, if SourceBlock has been defined
        $description = $_.Description
        if ($IncludeLinks -and ![string]::IsNullOrEmpty($Notice.SourceBlock))
        {
            $description = ("<a href=`"#{0}`">{1}</a>" -f $_.SourceBlock, $description)
        }

        # Don't add the properties in here just yet. Want Section to be first, if specified
        $obj = [ordered]@{}

        # Add the section, if it has been defined
        if (![string]::IsNullOrEmpty($SectionName))
        {
            $obj["Section"] = $SectionName
        }

        # Add status and description properties
        $obj["Status"] = $_.Status
        $obj["Description"] = $description

        [PSCustomObject]$obj
    }
}

Function Format-ReportRunnerDecodeHtml
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Content
    )

    process
    {
        $output = $Content

        if (![string]::IsNullOrEmpty($output))
        {
            $output = [System.Web.HttpUtility]::HtmlDecode($output)
        }

        $output
    }
}

<#
#>
Function Update-ReportRunnerBlockData
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ReportRunnerSection]$Section,

        [Parameter(Mandatory=$true)]
        [ValidatePattern("^[a-zA-Z_-]*\.[a-zA-Z_-]*\.[a-zA-Z_-]*$")]
        [string]$Id,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [HashTable]$Data,

        [Parameter(Mandatory=$false)]
        [switch]$Replace = $false
    )

    process
    {
        if ($Section.Blocks.Keys -notcontains $Id)
        {
            Write-Error "Block with id ($Id) does not exist in section"
        }

        $block = $Section.Blocks[$Id]
        if ($Replace)
        {
            $block.Data = $Data.Clone()
        } else {
            $block.Data = $block.Data.Clone()
            $Data.Keys | ForEach-Object {
                $block.Data[$_] = $Data[$_]
            }
        }
    }
}

<#
#>
Function Set-ReportRunnerBlockSetting
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [ReportRunnerEncodeStatus]$EncodeStatus
    )

    process
    {
        $setting = [ReportRunnerBlockSettings]::New()

        if ($PSBoundParameters.Keys -contains "EncodeStatus")
        {
            $setting.EncodeStatus = $EncodeStatus
        }

        $setting
    }
}

Function Get-ReportRunnerDataProperty
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [HashTable]$Data,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Property,

        [Parameter(Mandatory=$false)]
        [AllowNull()]
        [AllowEmptyCollection()]
        $DefaultValue
    )

    process
    {
        $source = ""
        $value = $null

        # Determine the source of the property value
        if ($Data.Keys -contains $Property)
        {
            $value = $Data[$Property]
            $source = "supplied"
        } elseif ($PSBoundParameters.Keys -contains "DefaultValue")
        {
            $value = $DefaultValue
            $source = "default"
        } else {
            Write-Error "Missing property $Property in HashTable data and no default value"
        }

        $valueStr = $value

        # Make a null value readable
        if ($null -eq $valueStr)
        {
            $valueStr = "(null)"
        } else {
            $valueStr = [string]$value
        }

        # Put the string on a new line if it's >50 chars
        if ($valueStr.Length -gt 50)
        {
            # Truncate the string if it's greater than 80 chars
            if ($valueStr.Length -gt 80)
            {
                $valueStr = $valueStr.Substring(0, 80) + " ..."
            }

            Write-Information ("Using {0} value for property {1}:" -f $source, $Property)
            Write-Information $valueStr
        } else {
            Write-Information ("Using {0} value for property {1}: {2}" -f $source, $Property, $valueStr)
        }

        $value
    }
}
