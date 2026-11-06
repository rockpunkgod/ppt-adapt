param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $EvidenceDirectory)){throw 'Refusing existing outputs'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$pres=$null;$log=@()
try{
 $pres=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 $s=$pres.Slides.Item(1);$s.Select()
 foreach($i in @(1,2,3)){
  $shape=$s.Shapes.Item('focus-'+$i);$effect=$s.TimeLine.MainSequence.AddEffect($shape,59,0,2)
  $before=[single]$effect.EffectParameters.Size
  $pct=85;if($i -eq 2){$pct=120}
  if($before -gt 0 -and $before -le 3){$effect.EffectParameters.Size=[single]($pct/100)}else{$effect.EffectParameters.Size=[single]$pct}
  $effect.Timing.Duration=[single]1.5;$effect.Timing.TriggerDelayTime=[single]1
  $log+=@{shape=$shape.Name;effect='GrowShrink';defaultSize=$before;setSize=[single]$effect.EffectParameters.Size}
  if($i -ne 2){
   $tr=$s.TimeLine.MainSequence.AddEffect($shape,62,0,2);$defaultAmount=[single]$tr.EffectParameters.Amount
   # Preserve native default opacity 0.5; verify the persisted XML. Earlier zero was a PowerShell leading-dot cast bug, not an API limitation.
   $tr.Timing.Duration=[single]1.5;$tr.Timing.TriggerDelayTime=[single]1;$tr.Timing.RewindAtEnd=0
   for($b=1;$b -le $tr.Behaviors.Count;$b++){$tr.Behaviors.Item($b).Timing.RewindAtEnd=0}
   $log+=@{shape=$shape.Name;effect='Transparency';defaultAmount=$defaultAmount;setAmount=[single]$tr.EffectParameters.Amount}
  }
 }
 # Independent effects on later objects should share absolute page start, not cascade three separate 1s delays.
 $s.SlideShowTransition.AdvanceOnTime=-1;$s.SlideShowTransition.AdvanceTime=[single]5
 $s=$pres.Slides.Item(2);$s.Select();$seq=$s.TimeLine.MainSequence
 $e=$seq.AddEffect($s.Shapes.Item('wheel-core'),10,0,2);$e.Timing.Duration=[single]0.5
 $ring=$seq.AddEffect($s.Shapes.Item('wheel-ring'),21,0,3);$ring.Timing.Duration=[single]7
 for($j=1;$j -le $ring.Behaviors.Count;$j++){if([int]$ring.Behaviors.Item($j).Type -eq 7){$ring.Behaviors.Item($j).FilterEffect.Subtype=18}}
 for($i=1;$i -le 6;$i++){
  $node=$seq.AddEffect($s.Shapes.Item('wheel-node-'+$i),10,0,2);$node.Timing.Duration=[single]0.4
  $node.Timing.TriggerDelayTime=[single](($i-1)*7/6)
 }
 $s.SlideShowTransition.AdvanceOnTime=-1;$s.SlideShowTransition.AdvanceTime=[single]9.5
 $pres.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 foreach($idx in @(1,2)){$pres.Slides.Item($idx).Export((Join-Path $EvidenceDirectory "slide-$idx.png"),'PNG',1600,900)}
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'parameter-readback.json'),($log|ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
}finally{if($pres){$pres.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($pres)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null}
