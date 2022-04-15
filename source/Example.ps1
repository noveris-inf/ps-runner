

# Global settings
Set-StrictMode -Version 2
$InformationPreference = "Continue"
$ErrorActionPreference = "Stop"


Import-Module ./ReportRunner/ReportRunner.psm1

$context = New-ReportRunnerContext

Add-ReportRunnerDefinition -Name "example.script" -Script {
    Write-Information "Example Script"
}

Add-ReportRunnerContextSection -Context $context -Name "test1 name" -Description "test1 desc" -Data @{} -Scripts {
    Write-Warning "Standard warning message"
    "test1 message"
    New-ReportRunnerNotice -Status Warning -Description "Warning message 1"
    "Something"
    New-ReportRunnerNotice -Status Info -Description "Status message"

    # New-ReportRunnerFormatTable -Content (Get-Process)
} -LibraryMatches @(
    "example\..*"
)

Add-ReportRunnerContextSection -Context $context -Name "test2 name" -Description "test2 desc" -Data @{} -Scripts {
    "test1 message"
    Write-Information "Info message"
    New-ReportRunnerNotice -Status Warning -Description "Warning message 2"
    New-ReportRunnerNotice -Status Error -Description "Error message 2"

    # New-ReportRunnerFormatTable -Content (Get-Process)
    $obj = [PSCustomObject]@{
        Name = "test1"
        Content = "Content"
    }
    New-ReportRunnerFormatTable -Content $obj
} -LibraryMatches @(
    "example\..*"
)

Invoke-ReportRunnerContext -Context $context | Format-ReportRunnerContentAsHtml -Title "Test Report"
