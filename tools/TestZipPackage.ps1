# Exercise fixture identity without building, rendering or regenerating the corpus.
[CmdletBinding()]
param(
    [string] $OutputDirectory,
    [string] $ZipPackageScript = (Join-Path $PSScriptRoot 'ZipPackage.ps1')
)
$ErrorActionPreference='Stop'
if (!$OutputDirectory) {
    $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/zip-package-test-'+[guid]::NewGuid().ToString('N'))
}
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$root=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
$writer=(Resolve-Path -LiteralPath $ZipPackageScript).Path
. $writer

# Literal order distinguishes ordinal comparison from en-US/tr-TR collation.
[string[]] $names=@('A.xml','I.xml','[Content_Types].xml','a.xml','z.xml','ä.xml','ı.xml')
$entries=[hashtable]::new([StringComparer]::Ordinal)
foreach ($name in $names) { $entries[$name]='<text>'+ $name + ' α 😀' + "`r`n" + '</text>' }
$entries['a.xml']=''
$reversed=[hashtable]::new([StringComparer]::Ordinal)
for ($index=$names.Length-1; $index -ge 0; $index--) { $reversed[$names[$index]]=$entries[$names[$index]] }
$checks=[ordered]@{OrdinalEntryOrder=$true;FixedZipEpoch=$true;Utf8Payloads=$true;RepeatedIdentity=$true;OverwriteIdentity=$true;ChangedPayloadIdentity=$true}
$rows=@()
$before=[Globalization.CultureInfo]::CurrentCulture
try {
    foreach ($culture in @('en-US','tr-TR','sv-SE')) {
        [Globalization.CultureInfo]::CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo($culture)
        $path=Join-Path $root ($culture+'.zip')
        New-ZipPackage -Path $path -Entries $(if ($culture -eq 'en-US') {$entries} else {$reversed})
        $archive=[IO.Compression.ZipFile]::OpenRead($path)
        try {
            $actual=@($archive.Entries | ForEach-Object FullName)
            if ([string]::Join("`n",$actual) -cne [string]::Join("`n",$names)) { $checks.OrdinalEntryOrder=$false }
            $payloads=@()
            foreach ($entry in $archive.Entries) {
                if ($entry.LastWriteTime.DateTime -ne [datetime]::new(1980,1,1)) { $checks.FixedZipEpoch=$false }
                $input=$entry.Open()
                $buffer=[IO.MemoryStream]::new()
                try {
                    $input.CopyTo($buffer)
                    $bytes=$buffer.ToArray()
                    $expected=[Text.Encoding]::UTF8.GetBytes($entries[$entry.FullName])
                    if ([Convert]::ToHexString($bytes) -cne [Convert]::ToHexString($expected)) { $checks.Utf8Payloads=$false }
                    $payloads+=@{Name=$entry.FullName;Length=$bytes.Length;Sha256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes));Timestamp=$entry.LastWriteTime.ToString('s',[Globalization.CultureInfo]::InvariantCulture)}
                } finally { $buffer.Dispose(); $input.Dispose() }
            }
            $rows+=@{Culture=$culture;ZipSha256=(Get-FileHash -LiteralPath $path).Hash;Entries=$payloads}
        } finally { $archive.Dispose() }
    }
} finally { [Globalization.CultureInfo]::CurrentCulture=$before }
if (@($rows.ZipSha256 | Select-Object -Unique).Count -ne 1) { $checks.RepeatedIdentity=$false }

$overwrite=Join-Path $root 'overwrite.zip'
[IO.File]::WriteAllText($overwrite,'Obsolete generated output')
New-ZipPackage -Path $overwrite -Entries $reversed
if ((Get-FileHash -LiteralPath $overwrite).Hash -ne $rows[0].ZipSha256) { $checks.OverwriteIdentity=$false }
$changed=$entries.Clone()
$changed['z.xml']+=' Changed source.'
$changedPath=Join-Path $root 'changed.zip'
New-ZipPackage -Path $changedPath -Entries $changed
if ((Get-FileHash -LiteralPath $changedPath).Hash -eq $rows[0].ZipSha256) { $checks.ChangedPayloadIdentity=$false }

$failed=@($checks.Keys | Where-Object {!$checks[$_]})
[ordered]@{Writer=$writer;WriterSha256=(Get-FileHash -LiteralPath $writer).Hash;Passed=$checks.Count-$failed.Count;Failed=$failed.Count;Checks=$checks;Rows=$rows;
    Scope='Byte identity within this platform/runtime; ZIP platform metadata and compressor versions may differ across hosts. Entry bytes, names and timestamps must remain equivalent.'} |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $root 'evidence.json') -Encoding utf8
if ($failed.Count) { throw ('ZIP fixture controls failed: '+($failed -join ', ')) }
Write-Output ($checks.Count.ToString()+' ZIP fixture controls pass; no builds, rendering or corpus regeneration.')
