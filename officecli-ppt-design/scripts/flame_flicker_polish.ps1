param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Officecli)
$ErrorActionPreference='Stop';$OutputDeck=[IO.Path]::GetFullPath($OutputDeck)
if(Test-Path -LiteralPath $OutputDeck){throw 'Fresh output required'}
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$log=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,0)
 if($p.Slides.Count -ne 2 -or $p.Slides.Item(1).TimeLine.MainSequence.Count -ne 0){throw 'Fresh two-page candle fixture required'}
 $s=$p.Slides.Item(1)
 for($j=0;$j -lt 3;$j++){
  $e=$s.TimeLine.MainSequence.AddEffect($s.Shapes.Item('flame-'+$j),59,0,2)
  $e.EffectParameters.Size=[single]70
  # Direction rejects Grow/Shrink on this build; set the native scale behavior.
  $e.Behaviors.Item(1).ScaleEffect.ByX=[single]100;$e.Behaviors.Item(1).ScaleEffect.ByY=[single]70
  $e.Behaviors.Item(1).ScaleEffect.ToX=[single]100;$e.Behaviors.Item(1).ScaleEffect.ToY=[single]70
  $e.Timing.Duration=[single]0.2;$e.Timing.TriggerDelayTime=[single]($j*0.1)
  $e.Timing.Accelerate=[single]0.5;$e.Timing.Decelerate=[single]0.5;$e.Timing.AutoReverse=-1
  $log+=@{shape=$e.Shape.Name;size=$e.EffectParameters.Size;duration=$e.Timing.Duration;delay=$e.Timing.TriggerDelayTime;accelerate=$e.Timing.Accelerate;decelerate=$e.Timing.Decelerate;autoReverse=[int]$e.Timing.AutoReverse;byX=$e.Behaviors.Item(1).ScaleEffect.ByX;byY=$e.Behaviors.Item(1).ScaleEffect.ByY}
 }
 $p.SaveAs($OutputDeck,24)
}finally{if($p){$p.Close()}}
# Same indefinite-repeat authoring mechanism as the existing native loop recipes.
& $Officecli raw-set $OutputDeck '/slide[1]' --xpath "//p:cTn[@presetClass='emph']" --action setattr --xml 'repeatCount=indefinite' --json
if($LASTEXITCODE -ne 0){throw 'Repeat authoring failed'}
& $Officecli close $OutputDeck --json
if($LASTEXITCODE -ne 0){throw 'Close failed'}
[IO.File]::WriteAllText($OutputDeck+'.readback.json',($log|ConvertTo-Json -Depth 6))
