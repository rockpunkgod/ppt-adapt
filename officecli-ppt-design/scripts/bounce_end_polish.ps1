param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $EvidenceDirectory)){throw 'Refusing existing outputs'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$pres=$null;$log=@()
try{
 $pres=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$s=$pres.Slides.Item(1);$s.Select()
 if($pres.Slides.Count -ne 1 -or $s.TimeLine.MainSequence.Count -ne 0){throw 'Requires the fresh one-slide CLI baseline'}
 foreach($name in @('control','bounce')){
  $e=$s.TimeLine.MainSequence.AddEffect($s.Shapes.Item($name),2,0,2)
  $e.EffectParameters.Direction=10
  $e.Timing.Duration=[single]2;$e.Timing.TriggerDelayTime=[single]1
  $e.Timing.SmoothStart=0;$e.Timing.SmoothEnd=0;$e.Timing.Accelerate=[single]0;$e.Timing.Decelerate=[single]0
  $e.Timing.AutoReverse=0;$e.Timing.BounceEnd=0
  if($name -eq 'bounce'){
   $e.Timing.BounceEnd=-1;$e.Timing.BounceEndIntensity=[single]0.75
   for($j=1;$j -le $e.Behaviors.Count;$j++){
    $bh=$e.Behaviors.Item($j)
    if([int]$bh.Type -eq 5){$bh.Timing.BounceEnd=-1;$bh.Timing.BounceEndIntensity=[single]0.75}
    $log+=@{shape=$name;behavior=$j;type=[int]$bh.Type;bounce=[int]$bh.Timing.BounceEnd;intensity=[single]$bh.Timing.BounceEndIntensity}
   }
  }
  if($name -eq 'bounce' -and [Math]::Abs([single]$e.Timing.BounceEndIntensity-0.75) -gt 0.001){throw 'Bounce intensity did not persist'}
  $log+=@{shape=$name;duration=[single]$e.Timing.Duration;delay=[single]$e.Timing.TriggerDelayTime;bounce=[int]$e.Timing.BounceEnd;intensity=[single]$e.Timing.BounceEndIntensity}
 }
 $pres.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 $s.Export((Join-Path $EvidenceDirectory 'edit-state.png'),'PNG',1600,900)
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'readback.json'),($log|ConvertTo-Json -Depth 5),(New-Object Text.UTF8Encoding($false)))
 $movie=Join-Path $EvidenceDirectory 'native-60fps.mp4';$pres.CreateVideo($movie,$false,6,720,60,90)
 $limit=[DateTime]::UtcNow.AddSeconds(55)
 while([int]$pres.CreateVideoStatus -in @(0,1,2) -and [DateTime]::UtcNow -lt $limit){Start-Sleep -Milliseconds 200}
 if([int]$pres.CreateVideoStatus -ne 3){throw ('Native video status '+$pres.CreateVideoStatus)}
}finally{if($pres){$pres.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
