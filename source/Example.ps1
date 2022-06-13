

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

    [PSCustomObject]@{
        First = (@("test", "other") | ForEach-Object { $_ } | Out-String)
        Second = "Second"
    } | ConvertTo-ReportRunnerFormatTable

    New-ReportRunnerNotice -Status Warning -Description "example script warning"
}

# Create the report runner context, with optional data
$context = New-ReportRunnerContext -Title "Example Report" -Data @{ A=1; B=1; C=1 }

# Create a report runner section, with optional data
$section = New-ReportRunnerSection -Context $context -Name "Section b1" -Description "Section 1 description" -Data @{ B=2; C=2 }

# Add some blocks to this section
New-ReportRunnerBlock -Section $section -LibraryFilter "^example\.script\." -Data @{ C=3 }
New-ReportRunnerBlock -Section $section -Id "example.manual.first" -Name "Manual Block 1" -Description "Manual Block 1 description" -Data @{C=4} -Script {
    $data = $_

    Write-Information "Manual Block 1 Data:"
    $data | ConvertTo-Json
    $data.GetEnumerator() | ConvertTo-ReportRunnerFormatTable

    $testdata = Get-ReportRunnerDataProperty -Data $data -Property LargeCollection -DefaultValue (1..20)
    $testdata = Get-ReportRunnerDataProperty -Data $data -Property LargeString1 -DefaultValue ([string](1..20))
    $testdata = Get-ReportRunnerDataProperty -Data $data -Property LargeString1 -DefaultValue ([string](1..28))
    $testdata = Get-ReportRunnerDataProperty -Data $data -Property LargeString2 -DefaultValue ([string](1..100))

    $longValue = [string](1..20)
    $item = [PSCustomObject]@{
        Property1 = $longValue
        Property2 = $longValue
        Property3 = $longValue
        Property4 = $longValue
        Property5 = $longValue
        Property6 = $longValue
        Property7 = $longValue
        Property8 = $longValue
    }

    @($item) | ConvertTo-ReportRunnerFormatTable

    "Default Ignore"
    "<a href=`"https://www.google.com.au`">Google</a>"
    @(@{ Url = "<a href=`"https://www.google.com.au`">Google</a>" }) |
        ConvertTo-Html -As Table -Fragment

    "Encode Status Encode"
    Set-ReportRunnerBlockSetting -EncodeStatus Encode
    "<a href=`"https://www.google.com.au`">Google</a>"
    @(@{ Url = "<a href=`"https://www.google.com.au`">Google</a>" }) |
        ConvertTo-Html -As Table -Fragment

    "Encode Status Decode"
    Set-ReportRunnerBlockSetting -EncodeStatus Decode
    "<a href=`"https://www.google.com.au`">Google</a>"
    @(@{ Url = "<a href=`"https://www.google.com.au`">Google</a>" }) |
        ConvertTo-Html -As Table -Fragment

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
