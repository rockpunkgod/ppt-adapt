param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $EvidenceDirectory)){throw 'Fresh outputs required'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$rows=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 foreach($sidx in @(1,2)){
  $s=$p.Slides.Item($sidx);$s.SlideShowTransition.Duration=[single]2;$s.SlideShowTransition.AdvanceOnClick=-1;$s.SlideShowTransition.AdvanceOnTime=0
  foreach($pair in @(@('!!screen-hinge',$(if($sidx -eq 1){90}else{0})),@('!!keyboard-hinge',$(if($sidx -eq 1){270}else{290})))){
   $shape=$s.Shapes.Item([string]$pair[0]);$td=$shape.ThreeD;$td.SetPresetCamera(62);$td.FieldOfView=[single]45;$td.RotationX=[single]0;$td.RotationZ=[single]0;$td.RotationY=[single]$pair[1]
   $rows+=@{slide=$sidx;name=$shape.Name;type=[int]$shape.Type;left=[single]$shape.Left;top=[single]$shape.Top;width=[single]$shape.Width;height=[single]$shape.Height;rotationX=[single]$td.RotationX;rotationY=[single]$td.RotationY;rotationZ=[single]$td.RotationZ;fieldOfView=[single]$td.FieldOfView;presetCamera=[int]$td.PresetCamera;childCount=$shape.GroupItems.Count}
  }
  $s.Export((Join-Path $EvidenceDirectory ('slide-'+$sidx+'.png')),'PNG',1600,900)
 }
 $s=$p.Slides.Item(2);$video=$s.Shapes.Item('screen-video')
 if([int]$video.Type -ne 16 -or [int]$video.MediaType -ne 3){throw 'Expected true native embedded video'}
 $settings=$video.AnimationSettings.PlaySettings;$settings.PlayOnEntry=-1;$settings.PauseAnimation=0;$settings.LoopUntilStopped=-1;$settings.RewindMovie=-1;$settings.StopAfterSlides=1;$settings.HideWhileNotPlaying=0
 $seq=$s.TimeLine.MainSequence
 for($i=$seq.Count;$i -ge 1;$i--){if($seq.Item($i).Shape.Id -ne $video.Id){throw 'Unexpected non-video effect: refusing to clear'};$seq.Item($i).Delete()}
 $fade=$seq.AddEffect($video,10,0,2);$fade.Timing.Duration=[single]1;$fade.Timing.TriggerType=2;$fade.Timing.TriggerDelayTime=[single]0
 $play=$seq.AddEffect($video,83,0,2);$play.Timing.TriggerType=2;$play.Timing.TriggerDelayTime=[single]0
 $rows+=@{video=$video.Name;videoId=$video.Id;type=[int]$video.Type;mediaType=[int]$video.MediaType;lengthMs=$video.MediaFormat.Length;left=[single]$video.Left;top=[single]$video.Top;width=[single]$video.Width;height=[single]$video.Height;fadeDuration=[single]$fade.Timing.Duration;fadeTrigger=[int]$fade.Timing.TriggerType;playTrigger=[int]$play.Timing.TriggerType;playOnEntry=[int]$settings.PlayOnEntry;loop=[int]$settings.LoopUntilStopped;rewind=[int]$settings.RewindMovie;stopAfterSlides=$settings.StopAfterSlides}
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 $p.Close();$p=$null;$p=$app.Presentations.Open([IO.Path]::GetFullPath($OutputDeck),-1,0,-1)
 foreach($i in @(1,2)){
  $s=$p.Slides.Item($i);$screen=$s.Shapes.Item('!!screen-hinge');$keyboard=$s.Shapes.Item('!!keyboard-hinge')
  $rows+=@{reopened=$i;screenY=[single]$screen.ThreeD.RotationY;keyboardY=[single]$keyboard.ThreeD.RotationY;screenFov=[single]$screen.ThreeD.FieldOfView;keyboardFov=[single]$keyboard.ThreeD.FieldOfView;screenChildren=$screen.GroupItems.Count;keyboardChildren=$keyboard.GroupItems.Count}
  $s.Export((Join-Path $EvidenceDirectory ('reopened-'+$i+'.png')),'PNG',1600,900)
 }
 $v=$p.Slides.Item(2).Shapes.Item('screen-video');$rows+=@{reopenedVideo=$v.Name;mediaType=[int]$v.MediaType;type=[int]$v.Type;lengthMs=$v.MediaFormat.Length;effects=$p.Slides.Item(2).TimeLine.MainSequence.Count;loop=[int]$v.AnimationSettings.PlaySettings.LoopUntilStopped}
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory '3d-readback.json'),($rows|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
