# Public inherited and ordinary controls for minimum-height body baselines.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($spec in @(@('note-default-minimum-line-16',16,$false),@('ordinary-minimum-line-20',20,$true))){
 $entries=@{}
 $archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases ('note-minimum-body-line-'+$spec[1]+'.docx')))
 try{foreach($entry in $archive.Entries){$reader=[IO.StreamReader]::new($entry.Open());try{$entries[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}}finally{$archive.Dispose()}
 [xml]$document=$entries['word/document.xml']
 $ns=[Xml.XmlNamespaceManager]::new($document.NameTable);$ns.AddNamespace('w',$word)
 if($spec[2]){
  foreach($ref in @($document.SelectNodes('//w:footnoteReference | //w:footnotePr',$ns))){[void]$ref.ParentNode.RemoveChild($ref)}
  [void]$entries.Remove('word/footnotes.xml')
  foreach($name in @('word/_rels/document.xml.rels','[Content_Types].xml')){
   [xml]$package=$entries[$name]
   foreach($entry in @($package.DocumentElement.ChildNodes|Where-Object {$_.Type -like '*footnotes*' -or $_.PartName -eq '/word/footnotes.xml'})){[void]$entry.ParentNode.RemoveChild($entry)}
   $entries[$name]=$package.OuterXml
  }
  [xml]$settings=$entries['word/settings.xml'];$sn=[Xml.XmlNamespaceManager]::new($settings.NameTable);$sn.AddNamespace('w',$word)
  foreach($property in @($settings.SelectNodes('//w:footnotePr',$sn))){[void]$property.ParentNode.RemoveChild($property)}
  $entries['word/settings.xml']=$settings.OuterXml
 }else{
  [xml]$styles=$entries['word/styles.xml'];$st=[Xml.XmlNamespaceManager]::new($styles.NameTable);$st.AddNamespace('w',$word)
  $defaults=$styles.SelectSingleNode('/w:styles/w:docDefaults',$st)
  $previous=$defaults.SelectSingleNode('w:pPrDefault',$st);if($previous){[void]$defaults.RemoveChild($previous)}
  $properties=$styles.CreateElement('w','pPrDefault',$word)
  $properties.InnerXml="<w:pPr xmlns:w='$word'><w:spacing w:line='320' w:lineRule='atLeast'/></w:pPr>"
  [void]$defaults.AppendChild($properties)
  foreach($paragraph in @($document.SelectNodes('//w:p[starts-with(w:r/w:t,"Public minimum ")]',$ns))){
   $spacing=$paragraph.SelectSingleNode('w:pPr/w:spacing',$ns)
   $spacing.RemoveAttribute('line',$word);$spacing.RemoveAttribute('lineRule',$word)
  }
  $entries['word/styles.xml']=$styles.OuterXml
 }
 $entries['word/document.xml']=$document.OuterXml
 New-ZipPackage -Path (Join-Path $OutputDirectory ($spec[0]+'.docx')) -Entries $entries
}
