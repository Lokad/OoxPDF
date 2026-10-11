# Public plain-footnote controls for minimum-height body line baselines.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$source=@{}
$archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases note-complete-reserve-3.docx))
try {
 foreach($entry in $archive.Entries){
  $reader=[IO.StreamReader]::new($entry.Open())
  try{$source[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}
 }
}finally{$archive.Dispose()}
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($height in @(16,20)){
 $entries=@{}+$source
 [xml]$document=$entries['word/document.xml']
 $ns=[Xml.XmlNamespaceManager]::new($document.NameTable);$ns.AddNamespace('w',$word)
 $body=$document.SelectSingleNode('/w:document/w:body',$ns)
 foreach($p in @($body.SelectNodes('w:p[not(.//w:footnoteReference)]',$ns))){[void]$body.RemoveChild($p)}
 foreach($p in @($body.SelectNodes('w:p',$ns))){$texts=@($p.SelectNodes('.//w:t',$ns));$texts[-1].InnerText+=' AV To WA.'}
 $tags=@('Top','AfterAlpha','AfterBravo','AfterDelta');$positions=@(0,2,4,7)
 foreach($i in 0..3){
  $p=$document.CreateElement('w','p',$word)
  $p.InnerXml="<w:pPr xmlns:w='$word'><w:spacing w:before='0' w:after='0' w:line='$($height*20)' w:lineRule='atLeast'/></w:pPr><w:r xmlns:w='$word'><w:rPr><w:rFonts w:ascii='Calibri' w:hAnsi='Calibri'/><w:sz w:val='24'/></w:rPr><w:t>Public minimum $($tags[$i]). AV To WA.</w:t></w:r>"
  [void]$body.InsertBefore($p,$body.ChildNodes[$positions[$i]])
 }
 [xml]$notes=$entries['word/footnotes.xml']
 $noteNs=[Xml.XmlNamespaceManager]::new($notes.NameTable);$noteNs.AddNamespace('w',$word)
 foreach($note in @($notes.SelectNodes('/w:footnotes/w:footnote[not(@w:type)]',$noteNs))){
  $texts=@($note.SelectNodes('.//w:t',$noteNs));$texts[-1].InnerText+=' AV To WA.'
 }
 $entries['word/document.xml']=$document.OuterXml
 $entries['word/footnotes.xml']=$notes.OuterXml
 New-ZipPackage -Path (Join-Path $OutputDirectory ('note-minimum-body-line-'+$height+'.docx')) -Entries $entries
}
