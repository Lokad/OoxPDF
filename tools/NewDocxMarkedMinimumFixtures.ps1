# Public controls for source-owned note markers in minimum-height body slots.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($spec in @(@('note-marked-minimum-line-20',0),@('note-marked-minimum-after-8',8))){
 $entries=@{}
 $archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases note-minimum-body-line-20.docx))
 try{foreach($entry in $archive.Entries){$reader=[IO.StreamReader]::new($entry.Open());try{$entries[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}}finally{$archive.Dispose()}
 [xml]$document=$entries['word/document.xml']
 $ns=[Xml.XmlNamespaceManager]::new($document.NameTable);$ns.AddNamespace('w',$word)
 foreach($paragraph in @($document.SelectNodes('/w:document/w:body/w:p',$ns))){
  $spacing=$paragraph.SelectSingleNode('w:pPr/w:spacing',$ns)
  [void]$spacing.SetAttribute('line',$word,'400')
  [void]$spacing.SetAttribute('lineRule',$word,'atLeast')
  [void]$spacing.SetAttribute('after',$word,([int]$spec[1]*20).ToString([Globalization.CultureInfo]::InvariantCulture))
 }
 $entries['word/document.xml']=$document.OuterXml
 New-ZipPackage -Path (Join-Path $OutputDirectory ($spec[0]+'.docx')) -Entries $entries
}
