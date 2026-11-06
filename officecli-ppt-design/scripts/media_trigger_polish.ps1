param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Native)
$ErrorActionPreference='Stop';$OutputDeck=[IO.Path]::GetFullPath($OutputDeck);$Native=[IO.Path]::GetFullPath($Native)
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Native)){throw 'Fresh outputs required'};[void][IO.Directory]::CreateDirectory($Native)
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$rows=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$p.Windows.Item(1).Activate();$p.Windows.Item(1).ViewType=9
 if($p.Slides.Count -ne 2){throw 'Two-page scene/control required'}
 for($i=1;$i -le 2;$i++){
  $s=$p.Slides.Item($i);$s.Select();$v=$s.Shapes.Item('confetti-video');$gif=$s.Shapes.Item('fireworks-gif');$cake=$s.Shapes.Item('cake');$title=$s.Shapes.Item('happy-birthday')
  if([int]$v.Type -ne 16 -or [int]$v.MediaType -ne 3){throw 'Real video required'}
  if($s.TimeLine.InteractiveSequences.Count -ne 0){throw 'Do not overwrite existing triggers'}
  $settings=$v.AnimationSettings.PlaySettings;$settings.PlayOnEntry=0;$settings.LoopUntilStopped=0;$settings.HideWhileNotPlaying=0;$settings.RewindMovie=-1
  $main=$s.TimeLine.MainSequence;for($j=$main.Count;$j -ge 1;$j--){if($main.Item($j).Shape.Id -ne $v.Id){throw 'Unexpected non-video effect'};$main.Item($j).Delete()}
  $g=$s.TimeLine.InteractiveSequences.Add(1);$appear=$g.AddTriggerEffect($gif,1,4,$cake,'',0)
  $seq=$s.TimeLine.InteractiveSequences.Add(1);$entry=$seq.AddTriggerEffect($v,2,4,$title,'',0);$entry.EffectParameters.Direction=1;$entry.Timing.Duration=[single]0.5
  $targetSeq=$null
  for($q=1;$q -le $s.TimeLine.InteractiveSequences.Count;$q++){$candidate=$s.TimeLine.InteractiveSequences.Item($q);for($j=1;$j -le $candidate.Count;$j++){$ce=$candidate.Item($j);if($ce.Shape.Id -eq $v.Id -and [int]$ce.EffectType -eq 2){if($targetSeq){throw 'Multiple video Fly groups'};$targetSeq=$candidate}}}
  if(-not $targetSeq){throw 'Actual video trigger group not found'}
  $play=$targetSeq.AddEffect($v,83,0,2);$play.Timing.TriggerType=2;$play.Timing.TriggerDelayTime=0
  if($i -eq 1){$tr=$targetSeq.AddEffect($v,62,0,2);$tr.EffectParameters.Amount=[single]0.7;$tr.Timing.TriggerType=2;$tr.Timing.TriggerDelayTime=0;$tr.Timing.RewindAtEnd=0}
  $s.SlideShowTransition.AdvanceOnClick=0;$s.SlideShowTransition.AdvanceOnTime=0
  $rows+=@{slide=$i;gifTrigger=$appear.Timing.TriggerShape.Name;videoTrigger=$entry.Timing.TriggerShape.Name;flyDirection=[int]$entry.EffectParameters.Direction;flyDuration=$entry.Timing.Duration;interactiveSequences=$s.TimeLine.InteractiveSequences.Count;mainEffects=$main.Count}
 }
 $p.SlideShowSettings.ShowWithAnimation=-1;$p.SaveAs($OutputDeck,24);$p.Close();$p=$null
 $p=$app.Presentations.Open($OutputDeck,-1,0,-1)
 for($i=1;$i -le 2;$i++){
  $s=$p.Slides.Item($i);if($s.TimeLine.MainSequence.Count -ne 0 -or $s.TimeLine.InteractiveSequences.Count -ne 2){throw 'Native trigger groups changed'}
  for($q=1;$q -le 2;$q++){
   $group=$s.TimeLine.InteractiveSequences.Item($q);$trigger=$group.Item(1).Timing.TriggerShape.Name;$types=@();$targets=@()
   for($j=1;$j -le $group.Count;$j++){$e=$group.Item($j);$types+=[int]$e.EffectType;$targets+=$e.Shape.Name}
   if($trigger -eq 'cake'){if(($types -join ',') -ne '1' -or ($targets -join ',') -ne 'fireworks-gif'){throw 'GIF is not isolated in Cake trigger'}}
   elseif($trigger -eq 'happy-birthday'){$wanted=if($i -eq 1){'2,83,62'}else{'2,83'};if(($types -join ',') -ne $wanted -or @($targets|Where-Object{$_ -ne 'confetti-video'}).Count){throw 'Video effects are not in the title group'}}
   else{throw 'Unexpected trigger'}
   $rows+=@{reopenedSlide=$i;trigger=$trigger;effectTypes=$types;targets=$targets}
  }
  $s.Export((Join-Path $Native ('slide-'+$i+'.png')),'PNG',1600,900)
 }
 $rows|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $Native 'triggers.json') -Encoding UTF8
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
