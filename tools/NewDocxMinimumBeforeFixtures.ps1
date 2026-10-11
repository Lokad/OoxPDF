# Public controls for minimum note body before/after gaps and inherited page tops.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($spec in @(@('note-minimum-before-after-4',4,$false),@('note-minimum-before-inherited-first',8,$true))){
 $entries=@{}
 $archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases note-marked-minimum-line-20.docx))
 try{foreach($entry in $archive.Entries){$reader=[IO.StreamReader]::new($entry.Open());try{$entries[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}}finally{$archive.Dispose()}
 [xml]$document=$entries['word/document.xml']
 $ns=[Xml.XmlNamespaceManager]::new($document.NameTable);$ns.AddNamespace('w',$word)
 $paragraphs=@($document.SelectNodes('/w:document/w:body/w:p',$ns))
 foreach($paragraph in $paragraphs){
  $spacing=$paragraph.SelectSingleNode('w:pPr/w:spacing',$ns)
  [void]$spacing.SetAttribute('before',$word,([int]$spec[1]*20).ToString([Globalization.CultureInfo]::InvariantCulture))
  [void]$spacing.SetAttribute('after',$word,'80')
 }
 if($spec[2]){
  [void]$paragraphs[0].ParentNode.RemoveChild($paragraphs[0])
  foreach($paragraph in $paragraphs[1..($paragraphs.Count-1)]){$paragraph.SelectSingleNode('w:pPr/w:spacing',$ns).RemoveAttribute('before',$word)}
  [xml]$styles=$entries['word/styles.xml']
  $sn=[Xml.XmlNamespaceManager]::new($styles.NameTable);$sn.AddNamespace('w',$word)
  $defaults=$styles.SelectSingleNode('/w:styles/w:docDefaults',$sn)
  $pd=$defaults.SelectSingleNode('w:pPrDefault',$sn)
  if(!$pd){$pd=$styles.CreateElement('w','pPrDefault',$word);[void]$defaults.AppendChild($pd)}
  $pp=$pd.SelectSingleNode('w:pPr',$sn)
  if(!$pp){$pp=$styles.CreateElement('w','pPr',$word);[void]$pd.AppendChild($pp)}
  $sp=$pp.SelectSingleNode('w:spacing',$sn)
  if(!$sp){$sp=$styles.CreateElement('w','spacing',$word);[void]$pp.AppendChild($sp)}
  [void]$sp.SetAttribute('before',$word,'160')
  $entries['word/styles.xml']=$styles.OuterXml
  [xml]$notes=$entries['word/footnotes.xml']
  $nn=[Xml.XmlNamespaceManager]::new($notes.NameTable);$nn.AddNamespace('w',$word)
  foreach($p in @($notes.SelectNodes('//w:p',$nn))){
   $pp=$p.SelectSingleNode('w:pPr',$nn)
   if(!$pp){$pp=$notes.CreateElement('w','pPr',$word);[void]$p.PrependChild($pp)}
   $sp=$pp.SelectSingleNode('w:spacing',$nn)
   if(!$sp){$sp=$notes.CreateElement('w','spacing',$word);[void]$pp.AppendChild($sp)}
   [void]$sp.SetAttribute('before',$word,'0')
  }
  $entries['word/footnotes.xml']=$notes.OuterXml
 }
 $entries['word/document.xml']=$document.OuterXml
 New-ZipPackage -Path (Join-Path $OutputDirectory ($spec[0]+'.docx')) -Entries $entries
}
