

# Global settings
Set-StrictMode -Version 2
$InformationPreference = "Continue"
$ErrorActionPreference = "Stop"

Remove-Module ReportRunner -EA SilentlyContinue
Import-Module ./ReportRunner/ReportRunner.psm1

$context = New-ReportRunnerContext

Add-ReportRunnerDefinition -Name "example.script.name" -Script {
    $data = $_

    $message = $data["Message"]

    Write-Information "Example Script: $message"
}

Add-ReportRunnerSection -Context $context -Name "test1 name" -Description "test1 desc" -Data @{
    Message = "Test1 name message"
} -Items @({
    Write-Warning "Standard warning message"
    "test1 message"
    New-ReportRunnerNotice -Status Warning -Description "Warning message 1"
    "Something"
    New-ReportRunnerNotice -Status Info -Description "Status message"

    # New-ReportRunnerFormatTable -Content (Get-Process)
},{
    Write-Information "Second Script"
},"example\..*")

Add-ReportRunnerSection -Context $context -Name "test2 name" -Description "test2 desc" -Data @{
    Message = "Test2 name message"
} -Items @({
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
},{
    Write-Information "Second script"
},"example\..*")

Invoke-ReportRunnerContext -Context $context | Format-ReportRunnerContentAsHtml -Title "Test Report"
