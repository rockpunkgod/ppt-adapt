param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Evidence)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'Fresh output and evidence required'};[void][IO.Directory]::CreateDirectory($Evidence)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),-1,0,-1);$p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24);$s=$p.Slides.Item(1);$ink=$s.Shapes.Item('original-native-ink')
 if($ink.Type -notin @(22,23)){throw 'Expected genuine native Ink'}
 $matches=@();for($i=1;$i -le $s.TimeLine.MainSequence.Count;$i++){$e=$s.TimeLine.MainSequence.Item($i);if($e.Shape.Id -eq $ink.Id){$matches+=$e}}
 if($matches.Count -ne 1){throw 'Expected one preserved native Replay behavior, not EffectType recreation'};$replay=$matches[0]
 $audio=$s.Shapes.Item('audio');$audio.Name='ink-clock';if($audio.MediaFormat.Length -ne 7000){throw 'Expected seven-second original silent clock'}
 [void]$audio.MediaFormat.MediaBookmarks.Add(0,'InkCycle');$audio.AnimationSettings.PlaySettings.LoopUntilStopped=-1;$audio.AnimationSettings.PlaySettings.HideWhileNotPlaying=-1
 $anchor=$s.Shapes.AddShape(1,-50,-50,5,5);$anchor.Name='ink-bookmark-anchor';$anchor.Fill.Visible=0;$anchor.Line.Visible=0
 $group=$s.TimeLine.InteractiveSequences.Add();$anchorEffect=$group.AddTriggerEffect($anchor,10,5,$audio,'InkCycle',0);$anchorEffect.Timing.Duration=[single]0.001
 $replay.MoveAfter($anchorEffect);$replay.Shape=$ink
 $moved=$null;for($i=1;$i -le $s.TimeLine.InteractiveSequences.Count;$i++){$seq=$s.TimeLine.InteractiveSequences.Item($i);for($j=1;$j -le $seq.Count;$j++){$e=$seq.Item($j);if($e.Shape.Id -eq $ink.Id){$moved=$e}}}
 if(-not $moved){throw 'Native Replay did not survive group move'};$moved.Timing.TriggerType=2;$moved.Timing.TriggerDelayTime=0;$moved.Timing.Duration=5
 for($i=$s.TimeLine.MainSequence.Count;$i -ge 1;$i--){$e=$s.TimeLine.MainSequence.Item($i);if($e.Shape.Id -eq $ink.Id){$e.Delete()}}
 $p.Save();$s.Export([IO.Path]::GetFullPath((Join-Path $Evidence 'native.png')),'PNG',1600,900)
 @{inkType=$ink.Type;inkID=$ink.Id;duration=$moved.Timing.Duration;bookmark=$anchorEffect.Timing.TriggerBookmark;audioLength=$audio.MediaFormat.Length;repeat=$moved.Timing.RepeatCount;loop=$audio.AnimationSettings.PlaySettings.LoopUntilStopped;preservedNativeReplay=$true}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'structure.json') -Encoding UTF8
}catch{@{error=$_.Exception.ToString()}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'error.json') -Encoding UTF8;throw}
finally{if($p){try{$p.Close()}catch{@{error=$_.Exception.ToString()}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'close-error.json') -Encoding UTF8}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
