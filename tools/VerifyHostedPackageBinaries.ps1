# Verify candidate symbol identities and code continuity against a frozen .NET 11 build.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $StageDirectory
)
$ErrorActionPreference='Stop'
$repo=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if ((Get-Location).Path -cne $repo) { throw 'Run from the repository root' }
$stageItem=Get-Item -LiteralPath $StageDirectory -Force
$artifacts=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if ($stageItem -isnot [IO.DirectoryInfo] -or !$stageItem.FullName.StartsWith($artifacts,[StringComparison]::Ordinal)) {
    throw 'Expected a qualification directory under artifacts/'
}
for ($parent=$stageItem; $parent.FullName -cne $repo; $parent=$parent.Parent) {
    if ($parent.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Expected a physical qualification directory' }
}
$stage=[IO.Path]::GetRelativePath($repo,$stageItem.FullName).Replace('\','/')
foreach ($name in @('generated-name-difference.json','binary-continuity.json')) {
    if (Test-Path -LiteralPath "$stage/$name") { throw "Qualification receipt already exists at $stage/$name" }
}
$hosted=Get-Content -LiteralPath "$stage/bound-hosted.json" -Raw | ConvertFrom-Json
$local=Get-Content -LiteralPath "$stage/bound-inputs.json" -Raw | ConvertFrom-Json
if ($hosted.Revision -cne $local.Revision) { throw 'Comparison input revision differs' }

function Pin([string] $Path) {
    [ordered]@{Path=$Path.Replace('\','/');Sha256=(Get-FileHash -LiteralPath $Path).Hash}
}

function Test-File($Pin) {
    if ((Get-FileHash -LiteralPath $Pin.Path).Hash -cne $Pin.Sha256) { throw ('Pinned bytes differ at '+$Pin.Path) }
}

function Read-MethodFingerprint([string] $Path) {
    $stream=[IO.File]::OpenRead($Path)
    $pe=[Reflection.PortableExecutable.PEReader]::new($stream)
    $sink=[Security.Cryptography.IncrementalHash]::CreateHash([Security.Cryptography.HashAlgorithmName]::SHA256)
    try {
        $reader=[Reflection.Metadata.PEReaderExtensions]::GetMetadataReader($pe)
        $count=0
        $normalized=0
        $holderPattern='^System\.Text\.RegularExpressions\.Generated\.<RegexGenerator_g>([A-F0-9]{65})__(?:SvgNumberRegex_0|SvgPathTokenRegex_1|SvgTransformTokenRegex_2|Utilities)$'
        $generatedRows=[Collections.Generic.List[string]]::new()
        $holderHashes=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $holderTypes=0
        foreach ($typeHandle in $reader.TypeDefinitions) {
            $type=$reader.GetTypeDefinition($typeHandle)
            $name=$reader.GetString($type.Namespace)+'.'+$reader.GetString($type.Name)
            $canonicalized=$name -cmatch $holderPattern
            $originalName=$name
            if ($canonicalized) {
                $holderHash=$Matches[1]
                $null=$holderHashes.Add($holderHash)
                $holderTypes++
                $name=$name.Replace($holderHash, '<source>')
            }
            foreach ($methodHandle in $type.GetMethods()) {
                $method=$reader.GetMethodDefinition($methodHandle)
                if ($method.RelativeVirtualAddress -eq 0) { continue }
                $body=[Reflection.Metadata.PEReaderExtensions]::GetMethodBody($pe,$method.RelativeVirtualAddress)
                $digest=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($body.GetILBytes()))
                $signature=[Convert]::ToBase64String($reader.GetBlobBytes($method.Signature))
                $row=$name+'::'+$reader.GetString($method.Name)+'::'+$signature+'::'+$digest+"$([char]10)"
                $sink.AppendData([Text.Encoding]::UTF8.GetBytes($row))
                $count++
                if ($canonicalized) {
                    $normalized++
                    $generatedRows.Add($originalName+'::'+$reader.GetString($method.Name)+'::'+$signature+'::'+$digest)
                }
            }
        }
        if ($holderHashes.Count -ne 1 -or $holderTypes -ne 4 -or $normalized -le 0) { throw "Generated-name scope differs: $($holderHashes.Count) hashes, $holderTypes types, $normalized methods at $Path" }
        [ordered]@{Methods=$count;NormalizedNames=$normalized;Sha256=[Convert]::ToHexString($sink.GetHashAndReset());GeneratedRows=$generatedRows.ToArray()}
    }
    finally { $sink.Dispose();$pe.Dispose();$stream.Dispose() }
}

function Bind-Symbols($Library, $Symbols) {
    Test-File $Library
    Test-File $Symbols
    $file=[IO.File]::OpenRead($Library.Path)
    $symbolFile=[IO.File]::OpenRead($Symbols.Path)
    $pe=[Reflection.PortableExecutable.PEReader]::new($file)
    $pdb=[Reflection.Metadata.MetadataReaderProvider]::FromPortablePdbStream($symbolFile)
    try {
        $codeView=@($pe.ReadDebugDirectory() | Where-Object Type -EQ 'CodeView')
        if ($codeView.Count -ne 1) { throw 'Expected one CodeView entry' }
        $view=$pe.ReadCodeViewDebugDirectoryData($codeView[0])
        $reader=$pdb.GetMetadataReader()
        [byte[]]$id=@($reader.DebugMetadataHeader.Id)
        if ($id.Length -ne 20 -or $reader.Documents.Count -eq 0) { throw 'Portable symbol metadata differs' }
        $guid=[Guid]::new([byte[]]($id[0..15]))
        $stamp=[BitConverter]::ToUInt32($id,16)
        if ($guid -ne $view.Guid -or $stamp -ne $codeView[0].Stamp -or $view.Age -ne 1) { throw 'The portable symbols do not bind to the library' }
        [ordered]@{Library=$Library;Pdb=$Symbols;CodeViewPath=$view.Path;CodeViewId=$guid;Stamp=$stamp;Documents=$reader.Documents.Count}
    }
    finally { $pdb.Dispose();$pe.Dispose();$symbolFile.Dispose();$file.Dispose() }
}

Test-File $local.LocalIdentity
$localIdentity=Get-Content -LiteralPath $local.LocalIdentity.Path -Raw | ConvertFrom-Json -AsHashtable
$localLibrary=Join-Path (Split-Path $local.LocalIdentity.Path) 'src/Lokad.OoxPdf/bin/Release/net11.0/Lokad.OoxPdf.dll'
$localPin=Pin $localLibrary
if ($localIdentity.Revision -cne $local.LocalRevision -or $localIdentity.Framework -cne 'net11.0' -or
    $localPin.Sha256 -cne $localIdentity.LibrarySha256) { throw 'Local runtime bytes differ' }
$bindings=[ordered]@{}
foreach ($framework in @('net11.0','net10.0')) {
    $item=$hosted.Identities.$framework
    $bindings[$framework]=Bind-Symbols $item.Library $item.Pdb
}
$first=Read-MethodFingerprint $localLibrary
$next=Read-MethodFingerprint $hosted.Identities.'net11.0'.Library.Path
$delta=@(Compare-Object $first.GeneratedRows $next.GeneratedRows -CaseSensitive)
if ($delta.Count -notin @(0,(2*$first.NormalizedNames))) { throw 'Unexpected scope of generated-name differences' }
$deltaFile="$stage/generated-name-difference.json"
[ordered]@{LocalRows=$first.GeneratedRows;HostedRows=$next.GeneratedRows;Difference=$delta} |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $deltaFile -Encoding utf8
$null=$next.Remove('GeneratedRows')
if ($first.Methods -le 0 -or $first.Methods -ne $next.Methods -or $first.Sha256 -cne $next.Sha256) { throw 'Hosted method signatures and bodies differ' }
$receipt=[ordered]@{
    Revision=$hosted.Revision
    LocalLibrary=$localPin
    Hosted=(Pin "$stage/bound-hosted.json")
    Bindings=$bindings
    MethodFingerprint=$next
    GeneratedNameDifference=(Pin $deltaFile)
}
$receipt | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath "$stage/binary-continuity.json" -Encoding utf8
Write-Output "$($next.Methods) method signatures and bodies and both framework symbol bindings verify."
