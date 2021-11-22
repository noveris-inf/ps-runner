<#
#>

########
# Global settings
$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"
Set-StrictMode -Version 2

Class RunnerSection
{
    [string]$Name
    [string]$Description
    [ScriptBlock]$Script
    $Data

    RunnerSection([string]$name, [string]$description, [ScriptBlock]$script, [PSObject]$data)
    {
        $this.Name = $name
        $this.Description = $description
        $this.Script = $script
        $this.Data = $data
    }
}

class RunnerSectionContent
{
    [string]$Name
    [string]$Description
    [System.Collections.ArrayList]$Content

    RunnerSectionContent([string]$Name, [string]$Description)
    {
        $this.Name = $name
        $this.Description = $description
        $this.Content = New-Object 'System.Collections.ArrayList'
    }
}

Class RunnerFormatTable
{
    $Content
}

Class RunnerContext
{
    [System.Collections.Generic.List[RunnerSection]]$Entries

    RunnerContext()
    {
        $this.Entries = New-Object 'System.Collections.Generic.List[RunnerSection]'
    }
}

enum RunnerStatus
{
    None = 0
    Info
    Warning
    Error
    InternalError
}

<#
#>
Class RunnerNotice
{
    [RunnerStatus]$Status
    [string]$Description

    RunnerNotice([RunnerStatus]$status, [string]$description)
    {
        $this.Status = $status
        $this.Description = $description
    }

    [string] ToString()
    {
        return ("{0}: {1}" -f $this.Status.ToString().ToUpper(), $this.Description)
    }
}

<#
#>
Function New-RunnerNotice
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Description,

        [Parameter(mandatory=$false)]
        [ValidateNotNull()]
        [RunnerStatus]$Status = [RunnerStatus]::None
    )

    process
    {
        $notice = New-Object RunnerNotice -ArgumentList $Status, $Description

        $notice
    }
}

<#
#>
Function New-RunnerFormatTable
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true)]
        [ValidateNotNull()]
        $Content
    )

    process
    {
        $format = New-Object 'RunnerFormatTable'
        $format.Content = $Content

        $format
    }
}

<#
#>
Function New-RunnerContext
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    [OutputType('RunnerContext')]
    param(
    )

    process
    {
        $obj = New-Object RunnerContext

        $obj
    }
}

<#
#>
Function Add-RunnerContextSection
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [ValidateNotNull()]
        [RunnerContext]$Context,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [AllowEmptyString()]
        [string]$Description = "",

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ScriptBlock]$Script,

        [Parameter(Mandatory=$false)]
        [AllowNull()]
        $Data = $null
    )

    process
    {
        # Add the script to the list of scripts to process
        $entry = New-Object 'RunnerSection' -ArgumentList $Name, $Description, $Script, $Data
        $Context.Entries.Add($entry) | Out-Null
    }
}

<#
#>
Function Invoke-RunnerContext
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [RunnerContext]$Context
    )

    process
    {
        $Context.Entries | ForEach-Object {
            $entry = $_

            # Output a section format object
            $content = New-Object 'RunnerSectionContent' -ArgumentList $entry.Name, $entry.Description

            Invoke-Command -NoNewScope {
                # Run the script block
                try {
                    ForEach-Object -InputObject $entry.Data, -Process $entry.Script
                } catch {
                    New-RunnerNotice -Status InternalError -Description "Error running script: $_"
                }
            } *>&1 | ForEach-Object {
                $content.Content.Add($_) | Out-Null
            }

            $content
        }
    }
}

<#
#>
Function Format-RunnerContentAsHtml
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '')]
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [ValidateNotNull()]
        [RunnerSectionContent]$Section,

        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [ValidateNotNull()]
        [string]$Title = "",

        [Parameter(Mandatory=$false)]
        [bool]$DecodeHtml = $true
    )

    begin
    {
        # Collection of all notices across all sections
        $allNotices = [ordered]@{}

        $allSectionContent = New-Object 'System.Collections.ArrayList'

        # Html preamble
        "<!DOCTYPE html PUBLIC `"-//W3C//DTD XHTML 1.0 Strict//EN`"  `"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd`">"
        "<html xmlns=`"http://www.w3.org/1999/xhtml`">"
        "<head>"
        "<title>$Title</title>"
        "<style>"
        "table {"
        "  font-family: Arial, Helvetica, sans-serif;"
        "  border-collapse: collapse;"
        "  width: 100%;"
        "}"
        "td, th {"
        "  border: 1px solid #ddd;"
        "  padding: 8px;"
        "}"
        "tr:nth-child(even){background-color: #f2f2f2;}"
        "tr:hover {background-color: #ddd;}"
        "th {"
        "  padding-top: 12px;"
        "  padding-bottom: 12px;"
        "  text-align: left;"
        "  background-color: #04AA6D;"
        "  color: white;"
        "}"
        "</style>"
        "</head><body>"
    }

    process
    {
        # Generate string content for this section
        $sectionContent = & {
            $notices = New-Object 'System.Collections.Generic.List[string]'
            $allNotices[$Section.Name] = New-Object 'System.Collections.Generic.List[string]'

            # Display section heading
            ("<b>Section: {0}</b><br>" -f $Section.Name)
            ("<i>{0}</i><br><p />" -f $Section.Description)

            $output = $Section.Content | ForEach-Object {

                # Default message to pass on in pipeline
                $msg = $_

                # Check if it is a string or status object
                if ([RunnerNotice].IsAssignableFrom($msg.GetType()))
                {
                    [RunnerNotice]$notice = $_
                    $noticeStr = $notice.ToString()
                    $notices.Add($noticeStr) | Out-Null

                    # Only add Notices that are issues to the all notices list
                    if ($notice.Status -eq [RunnerStatus]::Warning -or $notice.Status -eq [RunnerStatus]::Error -or
                        $notice.Status -eq [RunnerStatus]::InternalError)
                    {
                        $allNotices[$Section.Name].Add($noticeStr) | Out-Null
                    }

                    # Alter message to notice string representation
                    $msg = $noticeStr
                }

                if ([RunnerFormatTable].IsAssignableFrom($msg.GetType()))
                {
                    $msg = $msg.Content | ConvertTo-Html -As Table -Fragment
                }

                if ([string].IsAssignableFrom($msg.GetType()))
                {
                    $msg += "<br>"
                    if ($DecodeHtml)
                    {
                        $msg = [System.Web.HttpUtility]::HtmlDecode($msg)
                    }
                }

                # Pass message on in the pipeline
                $msg
            }

            # Display notices for this section
            if (($notices | Measure-Object).Count -gt 0)
            {
                "<u>Notices:</u><br><ul>"
                $notices | ForEach-Object {
                    ("<li>{0}</li>" -f $_)
                }
                "</ul><br>"
            }

            # Display output
            "<u>Content:</u><br>"
            $output | Out-String

            "<p />"
        } | Out-String

        $allSectionContent.Add($sectionContent) | Out-Null
    }

    end
    {
        # Display all notices here
        "<u>All Notices:</u><br><ul>"
        $allNotices.Keys | ForEach-Object {
            $key = $_
            ("<li>{0}</li>" -f $key)

            "<ul>"
            $allNotices[$key] | ForEach-Object {
                ("<li>{0}</li>" -f $_)
            }
            "</ul>"
        }
        "</ul><br>"

        # Display all section content
        $allSectionContent | ForEach-Object { $_ }

        # Wrap up HTML
        "</body></html>"
    }
}
