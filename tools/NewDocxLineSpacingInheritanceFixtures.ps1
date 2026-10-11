# Public controls for automatic numeric line overrides through defaults/styles.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$source=@{}
$archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases note-minimum-body-line-16.docx))
try {
 foreach($entry in $archive.Entries){
  $reader=[IO.StreamReader]::new($entry.Open())
  try{$source[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}
 }
}finally{$archive.Dispose()}
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($variant in @('default-line-auto','style-line-only')){
 $entries=@{}+$source
 [xml]$document=$entries['word/document.xml']
 [xml]$styles=$entries['word/styles.xml']
 $ns=[Xml.XmlNamespaceManager]::new($document.NameTable);$ns.AddNamespace('w',$word)
 $styleNs=[Xml.XmlNamespaceManager]::new($styles.NameTable);$styleNs.AddNamespace('w',$word)
 if($variant -eq 'default-line-auto'){
  $defaults=$styles.SelectSingleNode('/w:styles/w:docDefaults',$styleNs)
  $previous=$defaults.SelectSingleNode('w:pPrDefault',$styleNs)
  if($previous){[void]$defaults.RemoveChild($previous)}
  $properties=$styles.CreateElement('w','pPrDefault',$word)
  $properties.InnerXml="<w:pPr xmlns:w='$word'><w:spacing w:line='320' w:lineRule='exact'/></w:pPr>"
  [void]$defaults.AppendChild($properties)
 }else{
  $parent=$styles.CreateElement('w','style',$word)
  [void]$parent.SetAttribute('type',$word,'paragraph');[void]$parent.SetAttribute('styleId',$word,'PublicLineParent')
  $parent.InnerXml="<w:basedOn xmlns:w='$word' w:val='Normal'/><w:pPr xmlns:w='$word'><w:spacing w:line='320' w:lineRule='exact'/></w:pPr>"
  [void]$styles.DocumentElement.AppendChild($parent)
 }
 foreach($p in @($document.SelectNodes('/w:document/w:body/w:p[not(.//w:footnoteReference)]',$ns))){
  $properties=$p.SelectSingleNode('w:pPr',$ns)
  $spacing=$properties.SelectSingleNode('w:spacing',$ns)
  if($variant -eq 'default-line-auto'){
   [void]$spacing.SetAttribute('line',$word,'240');[void]$spacing.SetAttribute('lineRule',$word,'auto')
  }else{
   $style=$document.CreateElement('w','pStyle',$word);[void]$style.SetAttribute('val',$word,'PublicLineParent')
   [void]$properties.PrependChild($style)
   [void]$spacing.SetAttribute('line',$word,'480');$spacing.RemoveAttribute('lineRule',$word)
  }
 }
 $entries['word/document.xml']=$document.OuterXml
 $entries['word/styles.xml']=$styles.OuterXml
 New-ZipPackage -Path (Join-Path $OutputDirectory ('note-'+$variant+'.docx')) -Entries $entries
}
