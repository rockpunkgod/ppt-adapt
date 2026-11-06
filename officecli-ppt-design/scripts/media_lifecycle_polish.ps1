param([string]$InputDeck,[string]$OutputDeck,[string]$Evidence,[int]$StopAfterSlides=2)
$ErrorActionPreference='Stop';if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'Fresh outputs required'};[void][IO.Directory]::CreateDirectory($Evidence)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),-1,0,-1);$p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 $s=$p.Slides.Item(1).Shapes.Item('lifecycle-audio');$play=$s.AnimationSettings.PlaySettings
 $play.PlayOnEntry=-1;$play.PauseAnimation=0;$play.StopAfterSlides=$StopAfterSlides;$play.LoopUntilStopped=0
 $m=$s.MediaFormat;if($m.StartPoint -ne 0 -or $m.EndPoint -ne 12000){throw 'Unexpected source bounds'}
 $p.Save();@{playOnEntry=[int]$play.PlayOnEntry;pauseAnimation=[int]$play.PauseAnimation;stopAfterSlides=$play.StopAfterSlides;loopUntilStopped=[int]$play.LoopUntilStopped;start=$m.StartPoint;end=$m.EndPoint;fadeIn=$m.FadeInDuration;fadeOut=$m.FadeOutDuration;length=$m.Length;mediaId=$s.Id}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'native-settings.json') -Encoding UTF8
}finally{if($p){try{$p.Close()}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
