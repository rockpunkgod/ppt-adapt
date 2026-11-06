param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Native)
$ErrorActionPreference='Stop'
$InputDeck=[IO.Path]::GetFullPath($InputDeck);$OutputDeck=[IO.Path]::GetFullPath($OutputDeck);$Native=[IO.Path]::GetFullPath($Native)
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Native)){throw 'Fresh outputs required'}
[void][IO.Directory]::CreateDirectory($Native)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open($InputDeck,0,0,-1);$p.Windows.Item(1).ViewType=9;$s=$p.Slides.Item(2);$s.Select();$cm=72/2.54
 $names=@();foreach($z in @(@{n='a';x=16;y=4.7;w=8;h=8},@{n='b';x=21.8;y=4.7;w=8;h=8},@{n='c';x=19;y=6.8;w=8;h=6.5})){$q=$s.Shapes.AddShape(9,[single]($z.x*$cm),[single]($z.y*$cm),[single]($z.w*$cm),[single]($z.h*$cm));$q.Name=$z.n;$names+=$z.n}
 $a=$s.Shapes.Item('a');$range=$s.Shapes.Range([object[]]$names);$range.Select();$range.MergeShapes(1,$a)
 $cut=$s.Shapes.Item($s.Shapes.Count);$cut.Name='custom-outline';$cut.Line.Visible=0
 $v=$s.Shapes.Item('video');$range=$s.Shapes.Range([object[]]@('video','custom-outline'));$range.Select();$range.MergeShapes(3,$v)
 $media=@();for($j=1;$j -le $s.Shapes.Count;$j++){if([int]$s.Shapes.Item($j).Type -eq 16){$media+=,$s.Shapes.Item($j)}}
 if($media.Count -ne 1){throw 'Intersect must retain exactly one native media object'}
 $v=$media[0];$v.Name='video-cropped';$settings=$v.AnimationSettings.PlaySettings
 $settings.PlayOnEntry=-1;$settings.PauseAnimation=0;$settings.StopAfterSlides=1;$settings.LoopUntilStopped=0;$settings.RewindMovie=0;$settings.HideWhileNotPlaying=0
 # PlayOnEntry creates a native Play effect here; reuse it instead of appending another.
 if($s.TimeLine.MainSequence.Count -ne 1){throw 'Expected one native Play after restoring PlayOnEntry'}
 $play=$s.TimeLine.MainSequence.Item(1)
 if([int]$play.EffectType -ne 83 -or [int]$play.Shape.Id -ne [int]$v.Id){throw 'Unexpected effect binding after restoring playback'}
 $play.Timing.TriggerType=2;$play.Timing.TriggerDelayTime=[single]0
 @{type=[int]$v.Type;mediaType=[int]$v.MediaType;durationMs=$v.MediaFormat.Length;mediaName=$v.Name;effectCount=$s.TimeLine.MainSequence.Count;playTrigger=[int]$play.Timing.TriggerType;shapeCount=$s.Shapes.Count}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Native 'readback.json') -Encoding utf8
 $p.Slides.Item(1).Export((Join-Path $Native 'slide-1.png'),'PNG',1600,900)
 $s.Export((Join-Path $Native 'slide-2.png'),'PNG',1600,900)
 $p.SaveAs($OutputDeck,24)
 for($j=1;$j -le $s.Shapes.Count;$j++){$q=$s.Shapes.Item($j);Write-Output ($q.Name+' / '+$q.Type)}
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
