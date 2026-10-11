# Public regular 11pt contextual minimum controls for ordinary and complete-note body flow.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($spec in @(@('ordinary-eleven-point-minimum', 'ordinary-minimum-line-20.docx', $false),@('note-eleven-point-minimum-alternating', 'note-marked-minimum-line-20.docx', $true))){
 $entries=@{}
 $archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases $spec[1]))
 try{foreach($entry in $archive.Entries){$reader=[IO.StreamReader]::new($entry.Open());try{$entries[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}}finally{$archive.Dispose()}
 [xml]$document=$entries['word/document.xml']
 $ns=[Xml.XmlNamespaceManager]::new($document.NameTable);$ns.AddNamespace('w',$word)
 $paragraphs=@($document.SelectNodes('/w:document/w:body/w:p',$ns))
 for($i=0;$i -lt $paragraphs.Count;$i++){
  foreach($run in @($paragraphs[$i].SelectNodes('w:r',$ns))){
   $rp=$run.SelectSingleNode('w:rPr',$ns)
   if(!$rp){$rp=$document.CreateElement('w','rPr',$word);[void]$run.PrependChild($rp)}
   $size=$rp.SelectSingleNode('w:sz',$ns)
   if(!$size){$size=$document.CreateElement('w','sz',$word);[void]$rp.AppendChild($size)}
   [void]$size.SetAttribute('val',$word,'22')
  }
  $pr=$paragraphs[$i].SelectSingleNode('w:pPr',$ns)
  $spacing=$pr.SelectSingleNode('w:spacing',$ns)
  [void]$spacing.SetAttribute('before',$word,'80');[void]$spacing.SetAttribute('after',$word,'80')
  [void]$pr.AppendChild($document.CreateElement('w','contextualSpacing',$word))
  if($spec[2]){
   $style=$document.CreateElement('w','pStyle',$word)
   [void]$style.SetAttribute('val',$word,(@('PublicContextualA','PublicContextualB')[$i%2]))
   [void]$pr.PrependChild($style)
  }
 }
 if($spec[2]){
  [xml]$styles=$entries['word/styles.xml']
  foreach($id in @('PublicContextualA','PublicContextualB')){
   $style=$styles.CreateElement('w','style',$word)
   [void]$style.SetAttribute('type',$word,'paragraph');[void]$style.SetAttribute('styleId',$word,$id)
   $name=$styles.CreateElement('w','name',$word);[void]$name.SetAttribute('val',$word,$id)
   [void]$style.AppendChild($name);[void]$styles.DocumentElement.AppendChild($style)
  }
  $entries['word/styles.xml']=$styles.OuterXml
 }
 $entries['word/document.xml']=$document.OuterXml
 New-ZipPackage -Path (Join-Path $OutputDirectory ($spec[0]+'.docx')) -Entries $entries
}
