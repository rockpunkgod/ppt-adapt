param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Officecli)
$ErrorActionPreference='Stop'
$OutputDeck=[IO.Path]::GetFullPath($OutputDeck)
if(Test-Path -LiteralPath $OutputDeck){throw 'Fresh output required'}
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,0)
 if($p.Slides.Count -ne 1){throw 'One-page example required'}
 $s=$p.Slides.Item(1);$model=$s.Shapes.Item('interactive-model')
 if([int]$model.Type -ne 30 -or $s.TimeLine.MainSequence.Count -ne 0 -or $s.TimeLine.InteractiveSequences.Count -ne 0){throw 'Fresh native model timeline required'}
 # IDs read from actual native gallery choices on this Office build.
 $seq=$s.TimeLine.MainSequence
 $arrive=$seq.AddEffect($model,151,0,1);$arrive.Timing.Duration=1
 $turn=$seq.AddEffect($model,152,0,2);$turn.Timing.Duration=10;$turn.Timing.Accelerate=0;$turn.Timing.Decelerate=0
 $interactive=$s.TimeLine.InteractiveSequences.Add()
 $jump=$interactive.AddTriggerEffect($model,154,4,$model,'',0);$jump.Timing.Duration=2
 $p.SlideShowSettings.ShowWithAnimation=-1
 $p.SaveAs($OutputDeck,24);$p.Close();$p=$null
 # The native Continuous menu authors indefinite repeat with no acceleration.
 # COM reads RepeatCount as INT_MIN but rejects assigning that value.
 & $Officecli raw-set $OutputDeck '/slide[1]' --xpath "//p:cTn[@presetID='37' and @presetClass='emph']" --action setattr --xml 'repeatCount=indefinite'
 if($LASTEXITCODE -ne 0){throw 'Native continuous repeat failed'}
 & $Officecli close $OutputDeck
 if($LASTEXITCODE -ne 0){throw 'OfficeCLI close failed'}
 $p=$app.Presentations.Open($OutputDeck,0,0,0);$s=$p.Slides.Item(1)
 if($s.TimeLine.MainSequence.Count -ne 2 -or $s.TimeLine.InteractiveSequences.Count -ne 1){throw 'Reopen timeline mismatch'}
 $turn=$s.TimeLine.MainSequence.Item(2);$jump=$s.TimeLine.InteractiveSequences.Item(1).Item(1)
 if([int]$turn.EffectType -ne 152 -or $turn.Timing.Duration -ne 10 -or $turn.Timing.RepeatCount -ne -2147483648 -or [int]$jump.EffectType -ne 154 -or $jump.Timing.TriggerShape.Name -ne 'interactive-model'){throw 'Native 3D readback mismatch'}
 @{turntableSeconds=$turn.Timing.Duration;repeatReadback=$turn.Timing.RepeatCount;jumpTrigger=$jump.Timing.TriggerShape.Name;mainTypes=@($s.TimeLine.MainSequence|ForEach-Object{[int]$_.EffectType});jumpType=[int]$jump.EffectType}|ConvertTo-Json|Set-Content -LiteralPath ($OutputDeck+'.readback.json') -Encoding UTF8
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
