# Public bounded regular 11pt met-threshold glyph controls.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($spec in @(@('ordinary-eleven-point-times-met-glyphs','ordinary-eleven-point-glyphs.docx','Times New Roman'),@('note-eleven-point-arial-met-glyphs','note-eleven-point-arial-markers.docx','Arial'))){
 $entries=@{}
 $archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases $spec[1]))
 try{foreach($entry in $archive.Entries){$reader=[IO.StreamReader]::new($entry.Open());try{$entries[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}}finally{$archive.Dispose()}
 foreach($part in @('word/document.xml','word/styles.xml','word/footnotes.xml')){
  if(!$entries.ContainsKey($part)){continue}
  [xml]$xml=$entries[$part];$ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('w',$word)
  foreach($fonts in @($xml.SelectNodes('//w:rFonts',$ns))){foreach($attribute in @('ascii','hAnsi','eastAsia','cs')){[void]$fonts.SetAttribute($attribute,$word,$spec[2])}}
  if($part -eq 'word/document.xml'){
   foreach($run in @($xml.SelectNodes('/w:document/w:body/w:p/w:r',$ns))){
    $rp=$run.SelectSingleNode('w:rPr',$ns);if(!$rp){$rp=$xml.CreateElement('w','rPr',$word);[void]$run.PrependChild($rp)}
    $kern=$rp.SelectSingleNode('w:kern',$ns);if(!$kern){$kern=$xml.CreateElement('w','kern',$word);[void]$rp.AppendChild($kern)}
    [void]$kern.SetAttribute('val',$word,'22')
   }
  }
  $entries[$part]=$xml.OuterXml
 }
 New-ZipPackage -Path (Join-Path $OutputDirectory ($spec[0]+'.docx')) -Entries $entries
}
