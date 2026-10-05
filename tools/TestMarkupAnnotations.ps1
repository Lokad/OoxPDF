# Exercise the production markup annotation comparator without Office or a cache.
$ErrorActionPreference = 'Stop'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'CompareCachedDocxMarkupReference.ps1'), [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw 'Markup comparison script did not parse.' }
foreach ($function in $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    . ([scriptblock]::Create($function.Extent.Text))
}
function Assert-Equal($Expected, $Actual, [string] $Name) {
    if ($Expected -ne $Actual) { throw "$Name expected $Expected; actual $Actual" }
    Write-Host "PASS $Name"
}
$output = Join-Path (Split-Path -Parent $PSScriptRoot) ('artifacts/markup-annotation-tests/' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $output -Force | Out-Null
function New-AnnotationFixture([string] $Name, [int] $Offset, [string] $Target, [switch] $ReversePages) {
    # Physical object order deliberately differs from logical /Kids order.
    $catalog = $Offset + 1
    $root = $Offset + 2
    $branch = $Offset + 4
    $first = $Offset + 30
    $second = $Offset + 20
    $annotation = $Offset + 40
    $kids = if ($ReversePages) { "$second 0 R $first 0 R" } else { "$first 0 R $second 0 R" }
    $body = @"
%PDF-1.7
$catalog 0 obj << /Type /Catalog /Pages $root 0 R >> endobj
$second 0 obj << /Type /Page /Parent $branch 0 R >> endobj
$root 0 obj << /Type /Pages /Kids [$branch 0 R] /Count 2 >> endobj
$first 0 obj << /Type /Page /Parent $branch 0 R /Annots [$annotation 0 R] >> endobj
$branch 0 obj << /Type /Pages /Kids [$kids] /Count 2 >> endobj
$annotation 0 obj << /Subtype /Link /Rect [72 600 150 620] $Target >> endobj
trailer << /Root $catalog 0 R >>
%%EOF
"@
    $path = Join-Path $output "$Name.pdf"
    [IO.File]::WriteAllText($path, $body, [Text.Encoding]::ASCII)
    return @(Get-PdfAnnotations $path)[0]
}
function Compare-Target($Reference, $Candidate, [double] $Tolerance = 1) {
    return @(Compare-RectLists @($Reference) @($Candidate) $Tolerance)[0]
}
$reference = New-AnnotationFixture 'reference' 0 '/Dest [30 0 R/XYZ 72 700 0]'
$equivalent = New-AnnotationFixture 'equivalent' 300 '/Dest [330 0 R /XYZ 72.000 +700.0 null]'
Assert-Equal 'ok' (Compare-Target $reference $equivalent).Status 'Object numbers, whitespace, numeric spelling and inherited zoom do not change a destination'
Assert-Equal 1 $reference.Page 'Annotation source page follows the page tree, not object order'
Assert-Equal 1 $reference.DestinationPage 'Destination page follows the page tree'
Assert-Equal $reference.TargetSha256 $equivalent.TargetSha256 'Equivalent destinations have the same canonical hash'
$goTo = New-AnnotationFixture 'goto' 300 '/A << /D [330 0 R /XYZ 72 700 null] /S /GoTo >>'
Assert-Equal 'ok' (Compare-Target $reference $goTo).Status 'Local GoTo action and direct destination are equivalent'
$wrongPage = New-AnnotationFixture 'wrong-page' 0 '/Dest [30 0 R /XYZ 72 700 null]' -ReversePages
Assert-Equal $true (Compare-Target $reference $wrongPage).TargetDelta 'Same object number on a different logical page still fails'
$wrongSourcePage = New-AnnotationFixture 'wrong-source-page' 0 '/Dest [20 0 R /XYZ 72 700 null]' -ReversePages
Assert-Equal $false (Compare-Target $reference $wrongSourcePage).TargetDelta 'Equal destination pages still match when the source page moves'
Assert-Equal 'delta' (Compare-Target $reference $wrongSourcePage).Status 'Moving a link to another source page still fails'
$wrongSubtype = New-AnnotationFixture 'wrong-subtype' 300 '/Dest [330 0 R /XYZ 72 700 null]'
$wrongSubtype.Subtype = 'Widget'
Assert-Equal 'delta' (Compare-Target $reference $wrongSubtype).Status 'A different annotation subtype cannot pass on equal bounds'
$near = New-AnnotationFixture 'near' 300 '/Dest [330 0 R /XYZ 72.5 699.5 null]'
Assert-Equal 'ok' (Compare-Target $reference $near).Status 'Sub-tolerance target position differences pass'
Assert-Equal $true (Compare-Target $reference $near 0.1).TargetDelta 'The supplied target-position tolerance is honored'
$moved = New-AnnotationFixture 'moved' 300 '/Dest [330 0 R /XYZ 75 700 null]'
Assert-Equal $true (Compare-Target $reference $moved).TargetDelta 'A moved target still fails'
Assert-Equal 3 (Compare-Target $reference $moved).DestinationMaxPositionDelta 'Actual target position delta is reported'
$nullLeft = New-AnnotationFixture 'null-left' 300 '/Dest [330 0 R /XYZ null 700 null]'
$zeroLeft = New-AnnotationFixture 'zero-left' 300 '/Dest [330 0 R /XYZ 0 700 null]'
Assert-Equal $true (Compare-Target $nullLeft $zeroLeft).TargetDelta 'Retained position and coordinate zero are different'
$nullTop = New-AnnotationFixture 'null-top' 300 '/Dest [330 0 R /XYZ 700 null null]'
Assert-Equal $true (Compare-Target $nullLeft $nullTop).TargetDelta 'Null coordinate positions remain distinct'
$zoomOne = New-AnnotationFixture 'zoom-one' 300 '/Dest [330 0 R /XYZ 72 700 1]'
$zoomTwo = New-AnnotationFixture 'zoom-two' 300 '/Dest [330 0 R /XYZ 72 700 2]'
Assert-Equal $true (Compare-Target $zoomOne $zoomTwo 100).TargetDelta 'Zoom factors do not use the position tolerance'
$fit = New-AnnotationFixture 'fit' 300 '/Dest [330 0 R /Fit]'
Assert-Equal $true (Compare-Target $reference $fit).TargetDelta 'Different destination views still fail'
$fitEquivalent = New-AnnotationFixture 'fit-equivalent' 0 '/Dest[30 0 R/Fit]'
Assert-Equal 'ok' (Compare-Target $fit $fitEquivalent).Status 'Fit views are independent of object numbering'
Assert-Equal 0 $fit.DestinationParameters.Count 'Parameter-free views retain an empty parameter array'
$fitNull = New-AnnotationFixture 'fit-null' 300 '/Dest [330 0 R /FitH null]'
Assert-Equal 1 $fitNull.DestinationParameters.Count 'A single retained coordinate retains its parameter slot'
Assert-Equal $null $fitNull.DestinationParameters[0] 'A retained coordinate remains null'
$uri = New-AnnotationFixture 'uri' 0 '/A << /S /URI /URI (https://example.test/a) >>'
$escapedUri = New-AnnotationFixture 'escaped-uri' 300 '/A << /URI (https://example.test/\141) /S /URI >>'
Assert-Equal 'ok' (Compare-Target $uri $escapedUri).Status 'Equivalent escaped URI strings pass'
$otherUri = New-AnnotationFixture 'other-uri' 300 '/A << /S /URI /URI (https://example.test/b) >>'
Assert-Equal $true (Compare-Target $uri $otherUri).TargetDelta 'Different URI targets still fail'
$named = New-AnnotationFixture 'named' 0 '/Dest /FirstBookmark'
$otherNamed = New-AnnotationFixture 'other-named' 300 '/Dest /SecondBookmark'
Assert-Equal $true (Compare-Target $named $otherNamed).TargetDelta 'Unresolved named destinations retain exact target checks'
$badGeneration = New-AnnotationFixture 'generation' 300 '/Dest [330 1 R /XYZ 72 700 null]'
Assert-Equal $null $badGeneration.DestinationPage 'An unresolved object generation is not assigned a page'
Assert-Equal $true (Compare-Target $reference $badGeneration).TargetDelta 'Unresolved destinations cannot silently match a resolved target'
$indirect = New-AnnotationFixture 'indirect' 300 '/Dest 350 0 R'
Assert-Equal 'InternalDestination' $indirect.TargetKind 'Indirect destinations are retained as unresolved targets'
Assert-Equal $true (Compare-Target $reference $indirect).TargetDelta 'An unresolved indirect destination cannot silently pass'
$action = New-AnnotationFixture 'action' 0 '/A << /S /Named /N /NextPage >>'
$otherAction = New-AnnotationFixture 'other-action' 300 '/A << /S /Named /N /PrevPage >>'
Assert-Equal $true (Compare-Target $action $otherAction).TargetDelta 'Different unsupported actions retain exact target checks'
$cycle = Join-Path $output 'cycle.pdf'
[IO.File]::WriteAllText($cycle, '1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj 2 0 obj << /Type /Pages /Kids [2 0 R] /Count 1 >> endobj')
$rejectedCycle = $false
try { Get-PdfAnnotations $cycle | Out-Null } catch { $rejectedCycle = $_.Exception.Message -like '*repeated PDF page-tree reference*' }
Assert-Equal $true $rejectedCycle 'Cyclic page trees fail without looping'
$missingTree = Join-Path $output 'missing-tree.pdf'
[IO.File]::WriteAllText($missingTree, '40 0 obj << /Subtype /Link /Rect [72 600 150 620] /Dest [30 0 R /Fit] >> endobj')
$rejectedMissingTree = $false
try { Get-PdfAnnotations $missingTree | Out-Null } catch { $rejectedMissingTree = $_.Exception.Message -like '*page tree is unavailable*' }
Assert-Equal $true $rejectedMissingTree 'Unavailable page trees do not produce a silent partial comparison'
Write-Host 'Markup annotation checks passed.'
