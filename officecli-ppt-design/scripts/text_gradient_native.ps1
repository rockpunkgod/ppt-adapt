param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $EvidenceDirectory){throw 'Fresh evidence required'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$rows=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,-1)
 foreach($i in @(1,2)){
  $s=$p.Slides.Item($i);$shape=$s.Shapes.Item('glyph-gradient');$fill=$shape.TextFrame2.TextRange.Font.Fill;$stops=@()
  for($j=1;$fill.Type -eq 3 -and $j -le $fill.GradientStops.Count;$j++){$g=$fill.GradientStops.Item($j);$stops+=@{position=[single]$g.Position;transparency=[single]$g.Transparency;rgb=[int]$g.Color.RGB}}
  $rows+=@{slide=$i;shapeCount=$s.Shapes.Count;hasTextFrame=[int]$shape.HasTextFrame;text=$shape.TextFrame2.TextRange.Text;font=$shape.TextFrame2.TextRange.Font.Name;fontFillType=[int]$fill.Type;gradientStops=$stops}
  $s.Export((Join-Path $EvidenceDirectory ('native-'+$i+'.png')),'PNG',1600,900)
 }
 $saved=Join-Path $EvidenceDirectory 'roundtrip.pptx';$p.SaveAs([IO.Path]::GetFullPath($saved),24);$p.Close();$p=$null
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($saved),0,0,-1)
 foreach($i in @(1,2)){$s=$p.Slides.Item($i);$shape=$s.Shapes.Item('glyph-gradient');$s.Export((Join-Path $EvidenceDirectory ('reopened-'+$i+'.png')),'PNG',1600,900);$rows+=@{reopened=$i;hasTextFrame=[int]$shape.HasTextFrame;text=$shape.TextFrame2.TextRange.Text;fontFillType=[int]$shape.TextFrame2.TextRange.Font.Fill.Type;gradientStopCount=$(if($shape.TextFrame2.TextRange.Font.Fill.Type -eq 3){$shape.TextFrame2.TextRange.Font.Fill.GradientStops.Count}else{0})}}
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'native-readback.json'),($rows|ConvertTo-Json -Depth 7),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
