

# Global settings
Set-StrictMode -Version 2
$InformationPreference = "Continue"
$ErrorActionPreference = "Stop"

Remove-Module ReportRunner -EA SilentlyContinue
Import-Module ./ReportRunner/ReportRunner.psm1

# Add an example library function, which can be reused
Add-ReportRunnerLibraryBlock -Id "example.script.name" -Name "Example Script" -Description "Example Script for testing" -Script {
    $data = $_

    Write-Information "Example Script Data (JSON):"
    $data | ConvertTo-Json

    Write-Information "Example Script Data (Format-Table):"
    $data | Format-Table

    Write-Information "Example Script Data (New-ReportRunnerFormatTable):"
    $data.GetEnumerator() | ConvertTo-ReportRunnerFormatTable

    New-ReportRunnerNotice -Status Warning -Description "example script warning"
}

# Create the report runner context, with optional data
$context = New-ReportRunnerContext -Title "Example Report" -Data @{ A=1; B=1; C=1 }

# Create a report runner section, with optional data
$section = New-ReportRunnerSection -Context $context -Name "Section 1" -Description "Section 1 description" -Data @{ B=2; C=2 }

# Add some blocks to this section
New-ReportRunnerBlock -Section $section -LibraryFilter "^example\.script\." -Data @{ C=3 }
New-ReportRunnerBlock -Section $section -Id "example.manual.first" -Name "Manual Block 1" -Description "Manual Block 1 description" -Data @{C=4} -Script {
    $data = $_

    Write-Information "Manual Block 1 Data:"
    $data | ConvertTo-Json
    $data.GetEnumerator() | ConvertTo-ReportRunnerFormatTable

    New-ReportRunnerNotice -Status Info "manual block 1 info notice"
}

# Create a report runner section, with optional data
$section = New-ReportRunnerSection -Context $context -Name "Section 2" -Description "Section 2 description" -Data @{ B=2; C=2 }

# Add some blocks to this section
New-ReportRunnerBlock -Section $section -LibraryFilter "^example\.script\." -Data @{ C=3 }
New-ReportRunnerBlock -Section $section -Id "example.manual.second" -Name "Manual Block 2" -Description "Manual Block 2 description" -Data @{C=4} -Script {
    $data = $_

    Write-Information "Manual Block 1 Data:"
    $data | ConvertTo-Json
    $data.GetEnumerator() | ConvertTo-ReportRunnerFormatTable

    $testNum1 = 6
    $testNum1 = Get-ReportRunnerDataProperty -Data $data -Property TestNum1 -DefaultValue $null
    Write-Information "TestNum1: $testNum1"

    $incoming = Get-ReportRunnerDataProperty -Data $data -Property C -DefaultValue $null
    Write-Information "C: $incoming"

    $testNum2 = Get-ReportRunnerDataProperty -Data $data -Property TestNum2
    Write-Information "TestNum2: $testNum2"

    New-ReportRunnerNotice -Status Info "manual block 2 info notice"
}

Update-ReportRunnerBlockData -Section $section -Id "example.manual.second" -Data @{
    C = 5
}

Invoke-ReportRunnerContext -Context $context
Format-ReportRunnerContextAsHtml -Context $context | Out-String | Out-File Report.html
Format-ReportRunnerContextAsHtml -Context $context -SummaryOnly | Out-String | Out-File Summary.html
