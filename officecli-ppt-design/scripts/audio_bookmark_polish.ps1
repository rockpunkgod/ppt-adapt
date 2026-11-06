param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Evidence)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'New output paths required'}
[void][IO.Directory]::CreateDirectory($Evidence)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$s=$p.Slides.Item(1);$s.Select()
 $audio=$s.Shapes.Item('audio');$audio.Name='sequence-clock'
 $mark=$audio.MediaFormat.MediaBookmarks.Add(0,'CycleStart')
 $audio.AnimationSettings.PlaySettings.LoopUntilStopped=-1
 $audio.AnimationSettings.PlaySettings.HideWhileNotPlaying=-1
 $photo=$s.Shapes.Range(@('photo','photo-scene')).Group();$photo.Name='photo-card'
 $flash=$s.Shapes.Item('flash')
 $seq=$s.TimeLine.InteractiveSequences.Add()
 # Every effect shares the same media bookmark. Delays define one 2-second sequence.
 $rows=@()
 function Add-Bookmarked($shape,[int]$effect,[single]$duration,[single]$delay,[bool]$exit){
  if($seq.Count -eq 0){$e=$seq.AddTriggerEffect($shape,$effect,5,$audio,'CycleStart',0)}
  else{$e=$seq.AddEffect($shape,$effect,0,2)}
  if($exit){$e.Exit=-1}
  $e.Timing.Duration=$duration;$e.Timing.TriggerDelayTime=$delay
  $script:rows+=@{shape=$shape.Name;effect=$effect;duration=$duration;delay=$delay;exit=$exit;trigger=$e.Timing.TriggerType;bookmark=$e.Timing.TriggerBookmark}
  return $e
 }
 $e=Add-Bookmarked $flash 10 .1 0 $false
 $e=Add-Bookmarked $flash 10 .1 .12 $true
 $e=Add-Bookmarked $photo 22 .4 .2 $false
 $e.EffectParameters.Direction=3
 $e.Timing.Duration=.4
 $e=Add-Bookmarked $photo 127 .9 .6 $false
 $e.Behaviors.Item(1).MotionEffect.Path='M 0 0 L 0 0.04 E'
 $e.Timing.Accelerate=0;$e.Timing.Decelerate=0
 $e=Add-Bookmarked $photo 10 .3 1.7 $true
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 $s.Export((Join-Path $Evidence 'native.png'),'PNG',1600,900)
 @{bookmarkPosition=$mark.Position;bookmark=$mark.Name;audioLengthMs=$audio.MediaFormat.Length;loop=$audio.AnimationSettings.PlaySettings.LoopUntilStopped;effects=$rows}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $Evidence 'structure.json') -Encoding UTF8
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
