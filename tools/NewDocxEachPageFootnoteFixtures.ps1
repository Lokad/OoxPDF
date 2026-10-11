# Generates public, single-section footnote numbering controls for two page-break forms.
param([string]$OutputDirectory)

$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo 'tests/Lokad.OoxPdf.Tests/Cases'
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
. (Join-Path $PSScriptRoot 'ZipPackage.ps1')
$source=@{}
$archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases 'note-nav-numbered-ids.docx'))
try{
 foreach($entry in $archive.Entries){
  $reader=[IO.StreamReader]::new($entry.Open())
  try{$source[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}
 }
}finally{$archive.Dispose()}
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($boundary in @('run-break','paragraph-break')){
 $entries=@{}+$source
 [xml]$document=$entries['word/document.xml']
 $ns=[Xml.XmlNamespaceManager]::new($document.NameTable); $ns.AddNamespace('w',$word)
 $body=$document.SelectSingleNode('/w:document/w:body',$ns)
 $paragraphTemplate=$body.SelectSingleNode('w:p',$ns).CloneNode($true)
 $section=$body.SelectSingleNode('w:sectPr',$ns).CloneNode($true)
 foreach($node in @($section.SelectNodes('w:footnotePr|w:endnotePr|w:type',$ns))){[void]$section.RemoveChild($node)}
 $properties=$document.CreateElement('w','footnotePr',$word)
 $properties.InnerXml="<w:pos xmlns:w='$word' w:val='pageBottom'/><w:numFmt xmlns:w='$word' w:val='decimalZero'/><w:numStart xmlns:w='$word' w:val='9'/><w:numRestart xmlns:w='$word' w:val='eachPage'/>"
 [void]$section.PrependChild($properties)
 $body.RemoveAll()
 [xml]$notes=$entries['word/footnotes.xml']
 $noteNs=[Xml.XmlNamespaceManager]::new($notes.NameTable); $noteNs.AddNamespace('w',$word)
 $noteTemplate=$notes.SelectSingleNode('/w:footnotes/w:footnote[@w:id="37"]',$noteNs).CloneNode($true)
 foreach($note in @($notes.SelectNodes('/w:footnotes/w:footnote[not(@w:type)]',$noteNs))){[void]$notes.DocumentElement.RemoveChild($note)}
 $ids=@(37,4,73,8); $labels=@('Alpha','Bravo','Charlie','Delta')
 foreach($index in 0..3){
  if($index -eq 2 -and $boundary -eq 'run-break'){
   $p=$document.CreateElement('w','p',$word)
   $p.InnerXml="<w:r xmlns:w='$word'><w:br w:type='page'/></w:r>"
   [void]$body.AppendChild($p)
  }
  $paragraph=$paragraphTemplate.CloneNode($true)
  foreach($inlineSection in @($paragraph.SelectNodes('w:pPr/w:sectPr',$ns))){[void]$inlineSection.ParentNode.RemoveChild($inlineSection)}
  if($index -eq 2 -and $boundary -eq 'paragraph-break'){
   [void]$paragraph.SelectSingleNode('w:pPr',$ns).AppendChild($document.CreateElement('w','pageBreakBefore',$word))
  }
  $texts=@($paragraph.SelectNodes('.//w:t',$ns))
  $texts[0].InnerText=$labels[$index]+' public body before note '
  $texts[-1].InnerText=' and after note.'
  [void]$paragraph.SelectSingleNode('.//w:footnoteReference',$ns).SetAttribute('id',$word,[string]$ids[$index])
  [void]$body.AppendChild($paragraph)
  $note=$noteTemplate.CloneNode($true)
  [void]$note.SetAttribute('id',$word,[string]$ids[$index])
  $note.SelectSingleNode('.//w:t',$noteNs).InnerText=' Public destination '+$labels[$index]+'.'
  [void]$notes.DocumentElement.AppendChild($note)
 }
 [void]$body.AppendChild($section)
 $entries['word/document.xml']=$document.OuterXml
 $entries['word/footnotes.xml']=$notes.OuterXml
 $entries['word/settings.xml']="<w:settings xmlns:w='$word'><w:footnotePr><w:numFmt w:val='decimal'/><w:numStart w:val='105'/></w:footnotePr></w:settings>"
 New-ZipPackage -Path (Join-Path $OutputDirectory ('note-each-page-'+$boundary+'.docx')) -Entries $entries
}
