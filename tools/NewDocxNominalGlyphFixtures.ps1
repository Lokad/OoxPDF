# Public regular 11pt glyph-advance and automatic-marker controls.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($spec in @(@('ordinary-eleven-point-glyphs','ordinary-eleven-point-minimum.docx',$false),@('note-eleven-point-arial-markers','note-eleven-point-minimum-alternating.docx',$true))){
 $entries=@{}
 $archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases $spec[1]))
 try{foreach($entry in $archive.Entries){$reader=[IO.StreamReader]::new($entry.Open());try{$entries[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}}finally{$archive.Dispose()}
 [xml]$document=$entries['word/document.xml']
 $ns=[Xml.XmlNamespaceManager]::new($document.NameTable);$ns.AddNamespace('w',$word)
 if($spec[2]){
  foreach($part in @('word/document.xml','word/styles.xml','word/footnotes.xml')){
   [xml]$xml=$entries[$part];$map=[Xml.XmlNamespaceManager]::new($xml.NameTable);$map.AddNamespace('w',$word)
   foreach($fonts in @($xml.SelectNodes('//w:rFonts',$map))){foreach($attribute in @('ascii','hAnsi','eastAsia','cs')){[void]$fonts.SetAttribute($attribute,$word,'Arial')}}
   $entries[$part]=$xml.OuterXml
  }
 }else{
  $texts=@(('ProbeA'+('I'*16)),('ProbeB'+('M'*16)),('ProbeC'+('a'*16)),('ProbeD'+('i'*16)),('ProbeE'+('AV'*8)),('ProbeF'+('To'*8)),('ProbeG'+('WA'*8)),('ProbeH'+('0123456789'*2)))
  $paragraphs=@($document.SelectNodes('/w:document/w:body/w:p',$ns))
  if($paragraphs.Count -ne $texts.Count){throw 'Unexpected glyph fixture paragraph count'}
  for($i=0;$i -lt $paragraphs.Count;$i++){
   $p=$paragraphs[$i];$pr=$p.SelectSingleNode('w:pPr',$ns)
   foreach($child in @($p.ChildNodes)){if($child -ne $pr){[void]$p.RemoveChild($child)}}
   $run=$document.CreateElement('w','r',$word);$rp=$document.CreateElement('w','rPr',$word)
   $fonts=$document.CreateElement('w','rFonts',$word);foreach($attribute in @('ascii','hAnsi','eastAsia','cs')){[void]$fonts.SetAttribute($attribute,$word,'Calibri')}
   $size=$document.CreateElement('w','sz',$word);[void]$size.SetAttribute('val',$word,'22')
   [void]$rp.AppendChild($fonts);[void]$rp.AppendChild($size);[void]$run.AppendChild($rp)
   $text=$document.CreateElement('w','t',$word);$text.InnerText=$texts[$i];[void]$run.AppendChild($text);[void]$p.AppendChild($run)
  }
  $entries['word/document.xml']=$document.OuterXml
 }
 New-ZipPackage -Path (Join-Path $OutputDirectory ($spec[0]+'.docx')) -Entries $entries
}
