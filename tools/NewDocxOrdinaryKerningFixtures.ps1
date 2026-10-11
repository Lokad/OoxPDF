# Public note-free controls for inherited ordinary-body kerning thresholds.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$source=@{}
$archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases ordinary-minimum-line-20.docx))
try{foreach($entry in $archive.Entries){$reader=[IO.StreamReader]::new($entry.Open());try{$source[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}}finally{$archive.Dispose()}
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($variant in @('default-zero','style-zero')){
 $entries=@{}+$source
 [xml]$document=$entries['word/document.xml'];$ns=[Xml.XmlNamespaceManager]::new($document.NameTable);$ns.AddNamespace('w',$word)
 $body=$document.SelectSingleNode('/w:document/w:body',$ns)
 foreach($paragraph in @($body.SelectNodes('w:p[starts-with(w:r/w:t,"Public minimum ")]',$ns))){[void]$body.RemoveChild($paragraph)}
 foreach($kern in @($document.SelectNodes('//w:rPr/w:kern',$ns))){[void]$kern.ParentNode.RemoveChild($kern)}
 [xml]$styles=$entries['word/styles.xml'];$sn=[Xml.XmlNamespaceManager]::new($styles.NameTable);$sn.AddNamespace('w',$word)
 $defaults=$styles.SelectSingleNode('/w:styles/w:docDefaults',$sn)
 $runDefault=$defaults.SelectSingleNode('w:rPrDefault',$sn)
 if(!$runDefault){$runDefault=$styles.CreateElement('w','rPrDefault',$word);[void]$defaults.PrependChild($runDefault)}
 $properties=$runDefault.SelectSingleNode('w:rPr',$sn)
 if(!$properties){$properties=$styles.CreateElement('w','rPr',$word);[void]$runDefault.AppendChild($properties)}
 foreach($kern in @($properties.SelectNodes('w:kern',$sn))){[void]$properties.RemoveChild($kern)}
 $kern=$styles.CreateElement('w','kern',$word);[void]$kern.SetAttribute('val',$word, $(if($variant -eq 'default-zero'){'0'}else{'16'}));[void]$properties.AppendChild($kern)
 if($variant -eq 'style-zero'){
  $style=$styles.CreateElement('w','style',$word);[void]$style.SetAttribute('type',$word,'paragraph');[void]$style.SetAttribute('styleId',$word,'PublicOrdinaryKern')
  $style.InnerXml="<w:name xmlns:w='$word' w:val='Public ordinary kerning'/><w:basedOn xmlns:w='$word' w:val='Normal'/><w:rPr xmlns:w='$word'><w:kern w:val='0'/></w:rPr>"
  [void]$styles.DocumentElement.AppendChild($style)
  foreach($paragraph in @($body.SelectNodes('w:p',$ns))){
   $ppr=$paragraph.SelectSingleNode('w:pPr',$ns)
   foreach($oldStyle in @($ppr.SelectNodes('w:pStyle',$ns))){[void]$ppr.RemoveChild($oldStyle)}
   $pstyle=$document.CreateElement('w','pStyle',$word);[void]$pstyle.SetAttribute('val',$word,'PublicOrdinaryKern');[void]$ppr.PrependChild($pstyle)
  }
 }
 $entries['word/document.xml']=$document.OuterXml;$entries['word/styles.xml']=$styles.OuterXml
 New-ZipPackage -Path (Join-Path $OutputDirectory ('ordinary-kerning-'+$variant+'.docx')) -Entries $entries
}
