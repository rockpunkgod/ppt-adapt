param([string]$Directory,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$log=@()
try{
 foreach($stem in @('gauge','final-reference')){
  $p=$app.Presentations.Open((Join-Path ([IO.Path]::GetFullPath($Directory)) ($stem+'.pptx')),-1,0,0)
  for($i=1;$i -le 2;$i++){
   $s=$p.Slides.Item($i);$effects=@()
   foreach($e in $s.TimeLine.MainSequence){$effects+=@{target=$e.Shape.Name;targetType=[int]$e.Shape.Type;effectType=[int]$e.EffectType;duration=$e.Timing.Duration;delay=$e.Timing.TriggerDelayTime;accelerate=$e.Timing.Accelerate;decelerate=$e.Timing.Decelerate;repeatCount=$e.Timing.RepeatCount;trigger=[int]$e.Timing.TriggerType}}
   $g=$s.Shapes.Item('needle-pivot');$log+=@{deck=$stem;slide=$i;effects=$effects;groupType=[int]$g.Type;groupBox=@($g.Left,$g.Top,$g.Width,$g.Height);groupRotation=$g.Rotation;chartShapeType=[int]$s.Shapes.Item('gauge-data').Type;tens=$s.Shapes.Item('tens').TextFrame2.TextRange.Text;ones=$s.Shapes.Item('ones').TextFrame2.TextRange.Text}
   $s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ($stem+'-'+$i+'.png')),'PNG',1600,900)
  }
  $p.Close();$p=$null
 }
 $log|ConvertTo-Json -Depth 15|Set-Content -LiteralPath (Join-Path $Evidence 'objects.json') -Encoding UTF8
}finally{if($p){$p.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
