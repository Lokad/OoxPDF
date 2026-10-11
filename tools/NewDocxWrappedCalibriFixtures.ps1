# Public nominal-threshold Calibri wrapped-body advance/numbering controls.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force -Path $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($spec in @(@('Calibri','run-break','early','decimalZero','note-wrapped-calibri-zero-early',0),@('Calibri','paragraph-break','later','decimalZero','note-wrapped-calibri-unmet-later',26))){
 $entries=@{};$archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases ('note-each-page-'+$spec[1]+'.docx')))
 try {foreach($entry in $archive.Entries){$reader=[IO.StreamReader]::new($entry.Open());try{$entries[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}}finally{$archive.Dispose()}
 [xml]$document=$entries['word/document.xml'];$ns=[Xml.XmlNamespaceManager]::new($document.NameTable);$ns.AddNamespace('w',$word)
 $filler=('Public wrapped words retain authored digits 108. '*9)
 foreach($p in $document.SelectNodes('/w:document/w:body/w:p[w:r/w:footnoteReference]',$ns)){
  $texts=@($p.SelectNodes('.//w:t',$ns));$tag=$texts[0].InnerText.Split(' ')[0]
  $texts[0].InnerText=$tag+' '+$(if($spec[2] -eq 'later'){$filler}else{''})+'before note 105 '
  $texts[-1].InnerText=' and after note 108. '+$(if($spec[2] -eq 'early'){$filler}else{''})
 }
 [void]$document.SelectSingleNode('/w:document/w:body/w:sectPr/w:footnotePr/w:numFmt',$ns).SetAttribute('val',$word,$spec[3])
 $entries['word/document.xml']=$document.OuterXml
 foreach($part in @('word/document.xml','word/footnotes.xml')){
  [xml]$xml=$entries[$part];$ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('w',$word)
  foreach($rp in $xml.SelectNodes('//w:rPr',$ns)){
   $fonts=$rp.SelectSingleNode('w:rFonts',$ns)
   if(!$fonts){$fonts=$xml.CreateElement('w','rFonts',$word);[void]$rp.PrependChild($fonts)}
   foreach($attribute in @('ascii','hAnsi','eastAsia','cs')){[void]$fonts.SetAttribute($attribute,$word,$spec[0])}
  }
  $entries[$part]=$xml.OuterXml
 }
 [xml]$body=$entries['word/document.xml'];$ns=[Xml.XmlNamespaceManager]::new($body.NameTable);$ns.AddNamespace('w',$word)
 foreach($rp in $body.SelectNodes('//w:rPr',$ns)){
  $kern=$rp.SelectSingleNode('w:kern',$ns)
  if(!$kern){$kern=$body.CreateElement('w','kern',$word);[void]$rp.AppendChild($kern)}
  [void]$kern.SetAttribute('val',$word,[string]$spec[5])
 }
 $entries['word/document.xml']=$body.OuterXml
 New-ZipPackage -Path (Join-Path $OutputDirectory ($spec[4]+'.docx')) -Entries $entries
}
