param([Parameter(Mandatory=$true)][string]$InputDirectory,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $EvidenceDirectory){throw 'Refusing existing evidence'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$rows=@()
function Read-State($pres,[string]$state,[string]$stage){
 $r=@{state=$state;stage=$stage;shapes=@()}
 foreach($i in @(1,2)){
  $s=$pres.Slides.Item($i)
  foreach($name in @('theme-card','direct-card','theme-footer','theme-rule','direct-rule','major-reference','major-direct','minor-reference','minor-direct','latin-reference','latin-direct')){
   try{$q=$s.Shapes.Item($name)}catch{continue}
   $item=@{slide=$i;name=$name}
   try{$item.fillRgb=[int]$q.Fill.ForeColor.RGB;$item.fillTheme=[int]$q.Fill.ForeColor.ObjectThemeColor}catch{}
   if($q.TextFrame.HasText){
    $item.text=$q.TextFrame.TextRange.Text
    try{$item.fontLatin=$q.TextFrame.TextRange.Font.NameAscii;$item.fontEA=$q.TextFrame.TextRange.Font.NameFarEast;$item.textRgb=[int]$q.TextFrame.TextRange.Font.Color.RGB;$item.textTheme=[int]$q.TextFrame.TextRange.Font.Color.ObjectThemeColor;$item.boundWidth=[single]$q.TextFrame2.TextRange.BoundWidth}catch{$item.fontReadError=$_.Exception.Message}
   }
   $r.shapes+=$item
  }
 }
 try{$r.majorEA=$pres.SlideMaster.Theme.ThemeFontScheme.MajorFont(3).Name;$r.minorEA=$pres.SlideMaster.Theme.ThemeFontScheme.MinorFont(3).Name;$r.majorLatin=$pres.SlideMaster.Theme.ThemeFontScheme.MajorFont(1).Name}catch{$r.themeReadError=$_.Exception.Message}
 return $r
}
try{
 foreach($state in @('base','color','font')){
  $file=Join-Path $InputDirectory ('theme-'+$state+'.pptx')
  $p=$app.Presentations.Open([IO.Path]::GetFullPath($file),-1,0,-1)
  $rows+=Read-State $p $state 'first-open'
  foreach($i in @(1,2)){$p.Slides.Item($i).Export((Join-Path $EvidenceDirectory ($state+'-slide-'+$i+'.png')),'PNG',1600,900)}
  $roundtrip=Join-Path $EvidenceDirectory ('roundtrip-'+$state+'.pptx');$p.SaveAs([IO.Path]::GetFullPath($roundtrip),24);$p.Close();$p=$null
  $p=$app.Presentations.Open([IO.Path]::GetFullPath($roundtrip),-1,0,-1)
  $rows+=Read-State $p $state 'save-reopen'
  foreach($i in @(1,2)){$p.Slides.Item($i).Export((Join-Path $EvidenceDirectory ($state+'-reopen-'+$i+'.png')),'PNG',1600,900)}
  $p.Close();$p=$null
 }
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app);[IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'native-readback.json'),($rows|ConvertTo-Json -Depth 7),(New-Object Text.UTF8Encoding($false)))}
