# Fixture tooling only: build DocxEdit 0.2.0 at the revision documented below.
param(
    [Parameter(Mandatory=$true)][string]$ToolDll,
    [string]$OutputDirectory
)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$recipes=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases/DocxEditInterop
if(!$OutputDirectory){$OutputDirectory=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases}
$ToolDll=(Resolve-Path -LiteralPath $ToolDll).Path
$library=Join-Path (Split-Path -Parent $ToolDll) Lokad.DocxEdit.dll
if([Reflection.AssemblyName]::GetAssemblyName($library).Version -ne [Version]'0.2.0.0'){
    throw 'Use DocxEdit 0.2.0, revision 67e303117f3e51465f3da3d4971c357c091d407c.'
}
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$OutputDirectory=(Resolve-Path -LiteralPath $OutputDirectory).Path
$staging=Join-Path $OutputDirectory ('docxedit-fixtures-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $staging | Out-Null
$seed=Join-Path $staging empty.docx
function Invoke-Producer([string[]]$Arguments){
    $json= & dotnet $ToolDll @Arguments --json --compact
    if($LASTEXITCODE -ne 0){throw ('DocxEdit failed: '+($Arguments -join ' '))}
    $result=$json | ConvertFrom-Json
    if(!$result.Success){throw 'DocxEdit did not report success'}
    return $result
}
# The retained recipes target the single paragraph created by the pinned producer.
$null=Invoke-Producer -Arguments @('create','--output',$seed)
$readback=Invoke-Producer -Arguments @('read',$seed)
if($readback.Paragraphs.Count -ne 1 -or $readback.Paragraphs[0].Id -ne 'M.P0001'){
    throw 'Pinned producer seed paragraph differs'
}
$results=@()
foreach($patch in Get-ChildItem -LiteralPath $recipes -Filter '*.docxpatch' | Sort-Object Name){
    $output=Join-Path $OutputDirectory ($patch.BaseName+'.docx')
    $null=Invoke-Producer -Arguments @('apply',$seed,$patch.FullName,'--output',$output)
    $null=Invoke-Producer -Arguments @('validate',$output)
    $document=Invoke-Producer -Arguments @('read',$output)
    $results+=[ordered]@{File=$output;Sha256=(Get-FileHash -LiteralPath $output).Hash;PatchSha256=(Get-FileHash -LiteralPath $patch.FullName).Hash;Equations=$document.Equations.Count;Images=$document.Images.Count}
}
[ordered]@{ProducerRevision='67e303117f3e51465f3da3d4971c357c091d407c';ProducerLibrarySha256=(Get-FileHash -LiteralPath $library).Hash;Fixtures=$results} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $staging producer-report.json)
Write-Output ('Generated '+$results.Count+' fixtures; producer evidence: '+$staging)
