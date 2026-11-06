param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck)
$ErrorActionPreference='Stop';$OutputDeck=[IO.Path]::GetFullPath($OutputDeck)
if(Test-Path -LiteralPath $OutputDeck){throw 'Fresh output required'}
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$log=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,0)
 if($p.Slides.Count -ne 3){throw 'Three-page fresh brush fixture required'}
 foreach($i in @(2,3)){
  $s=$p.Slides.Item($i);$cover=$s.Shapes.Item('blurred-cover')
  if($cover.Fill.PictureEffects.Count -ne 0){throw 'Expected untouched picture'}
  $blur=$cover.Fill.PictureEffects.Insert(2);$blur.EffectParameters.Item(1).Value=[single]30
  $cover.PictureFormat.Brightness=[single]0.23
 }
 $s=$p.Slides.Item(3)
 if($s.TimeLine.MainSequence.Count -ne 0){throw 'Fresh timeline required'}
 for($j=0;$j -lt 6;$j++){
  $brush=$s.Shapes.Item('brush-'+$j);$brush.Fill.Background()
  if([int]$brush.Fill.Type -ne 5){throw 'Native background fill required'}
  $trigger=3;if($j -eq 0){$trigger=1}
  $e=$s.TimeLine.MainSequence.AddEffect($brush,22,0,$trigger)
  # Actual playback: Direction2/wipe(right) reveals from the right edge.
  $direction=2;if($j%2 -eq 1){$direction=4}
  $e.EffectParameters.Direction=$direction;$e.Timing.Duration=[single]0.5
  $log+=@{shape=$brush.Name;fillType=[int]$brush.Fill.Type;effect=[int]$e.EffectType;direction=$e.EffectParameters.Direction;duration=$e.Timing.Duration;trigger=[int]$e.Timing.TriggerType}
 }
 $e=$s.TimeLine.MainSequence.AddEffect($s.Shapes.Item('reveal-title'),10,0,3);$e.Timing.Duration=[single]0.5
 $p.SaveAs($OutputDeck,24)
 [IO.File]::WriteAllText($OutputDeck+'.readback.json',($log|ConvertTo-Json -Depth 6))
}finally{if($p){$p.Close()}}
