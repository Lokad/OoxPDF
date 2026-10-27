# Exercise the real line-start gate; no build, PDF conversion or Office.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference='Stop'
$PSNativeCommandUseErrorActionPreference=$false
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/line-start-test-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$root=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
$checks=[Collections.Generic.List[object]]::new()
function Op($X=72,$Y=700,$Page=1,$Text='Visible content') {
    [pscustomobject]@{PageNumber=$Page;X=$X;Y=$Y;DecodedText=$Text}
}
function Check([string] $Id,[object[]] $Reference,[object[]] $Candidate,[int] $Expected,[string[]] $Extra=@(),[switch] $GroupingOnly) {
    foreach ($mode in @('indexed','position')) {
        $prefix=Join-Path $root ($mode+'-'+$Id)
        ConvertTo-Json -InputObject @($Reference) -Depth 4 | Set-Content ($prefix+'-ref.json')
        ConvertTo-Json -InputObject @($Candidate) -Depth 4 | Set-Content ($prefix+'-cand.json')
        $arguments=@('-NoProfile','-File',(Join-Path $PSScriptRoot 'ComparePdfTextLineStarts.ps1'),
            '-Reference',($prefix+'-ref.json'),'-Candidate',($prefix+'-cand.json'),'-StartTolerance','0.15','-LineTolerance','0.5')
        if (!$GroupingOnly) { $arguments+=@('-BaselineTolerance','0.5') }
        if ($mode -eq 'position') { $arguments+='-MatchByPosition' }
        & pwsh @arguments @Extra *> ($prefix+'.log')
        $actual=$LASTEXITCODE
        $checks.Add(@{Id=$Id;Mode=$mode;ExpectedExit=$Expected;ActualExit=$actual})
        if ($actual -ne $Expected) { throw ($mode+'/'+$Id+': unexpected exit '+$actual+'; inspect '+$prefix+'.log') }
    }
}
$same=Op
Check 'identical' @($same) @($same) 0
Check 'within-y' @($same) @((Op 72 700.25)) 0
Check 'exact-y-boundary' @($same) @((Op 72 700.5)) 0
Check 'just-beyond-y' @($same) @((Op 72 700.5000001)) 1
Check 'positive-y-200pt' @($same) @((Op 72 900)) 1
Check 'negative-y-200pt' @($same) @((Op 72 500)) 1
Check 'grouping-without-y-assertion' @($same) @((Op 72 500)) 0 -GroupingOnly
Check 'within-x' @($same) @((Op 72.1 700)) 0
Check 'exact-x-boundary' @($same) @((Op 72.15 700)) 0
Check 'beyond-x' @($same) @((Op 72.150001 700)) 1
Check 'both-axes' @($same) @((Op 74 500)) 1
Check 'page-change' @($same) @((Op 72 700 2)) 1
Check 'missing-line' @($same,(Op 72 600)) @($same) 1
Check 'extra-line' @($same) @($same,(Op 72 600)) 1
Check 'missing-operation' @($same,(Op 100 700)) @($same) 1
Check 'source-order' @($same,(Op 72 600)) @((Op 72 600),$same) 0
Check 'grouped-baselines' @($same,(Op 100 700.25)) @($same,(Op 100 700.25)) 0
Check 'first-start-count' @($same,(Op 100 700)) @($same) 0 @('-FirstStartOnly')
Check 'first-start-y-still-bounded' @($same,(Op 100 700)) @((Op 72 500)) 1 @('-FirstStartOnly')
Check 'whitespace-default' @($same,(Op 72 600 1 ' ')) @($same) 1
Check 'whitespace-ignored' @($same,(Op 72 600 1 ([string][char]0xA0))) @($same) 0 @('-IgnoreWhitespaceOnlyOperations')
foreach ($coordinate in @('X','Y')) {
    foreach ($value in @('NaN','Infinity','-Infinity',$null)) {
        $bad=Op; $bad.$coordinate=$value
        $label=if ($null -eq $value) {'missing'} else {$value.Replace('-','negative-')}
        Check ('candidate-'+$coordinate+'-'+$label) @($same) @($bad) 1
        Check ('reference-'+$coordinate+'-'+$label) @($bad) @($same) 1
    }
}
foreach ($parameter in @('LineTolerance','StartTolerance','BaselineTolerance')) {
    foreach ($value in @('NaN','Infinity','-1')) {
        # Supply each argument once so binding itself cannot mask validation.
        $prefix=Join-Path $root ('invalid-'+$parameter+'-'+$value)
        ConvertTo-Json -InputObject @($same) | Set-Content ($prefix+'.json')
        $arguments=@('-NoProfile','-File',(Join-Path $PSScriptRoot 'ComparePdfTextLineStarts.ps1'),'-Reference',($prefix+'.json'),'-Candidate',($prefix+'.json'),('-'+$parameter),$value)
        & pwsh @arguments *> ($prefix+'.log')
        $checks.Add(@{Id=$parameter+'-'+$value;ExpectedExit=1;ActualExit=$LASTEXITCODE})
        if ($LASTEXITCODE -ne 1 -or (Get-Content ($prefix+'.log') -Raw) -notlike '*finite and non-negative*') { throw 'Invalid tolerance was not refused by the comparison boundary' }
    }
}
@{Checks=$checks;Builds=0;Conversions=0;OfficeExports=0} | ConvertTo-Json -Depth 5 |
    Set-Content (Join-Path $root 'evidence.json')
Write-Output ($checks.Count.ToString()+' line-start checks pass; Y limits, finite inputs, both matching modes and existing X/whitespace behavior covered; no builds or conversions.')
