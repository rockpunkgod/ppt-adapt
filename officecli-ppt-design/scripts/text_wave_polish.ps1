param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Native,[Parameter(Mandatory=$true)][string]$Officecli)
$ErrorActionPreference='Stop';$OutputDeck=[IO.Path]::GetFullPath($OutputDeck);$Native=[IO.Path]::GetFullPath($Native)
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Native)){throw 'Fresh outputs required'};[void][IO.Directory]::CreateDirectory($Native)
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$rows=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$p.Windows.Item(1).Activate();$p.Windows.Item(1).ViewType=9
 if($p.Slides.Count -ne 4){throw 'Four-page native Wave foundation required'}
 for($i=1;$i -le 4;$i++){
  $s=$p.Slides.Item($i);if($s.TimeLine.MainSequence.Count -ne 0){throw 'Do not overwrite an existing timeline'}
  if($i -le 3){
   $e=$s.TimeLine.MainSequence.AddEffect($s.Shapes.Item('wave-target'),82,0,1);$e=$s.TimeLine.MainSequence.ConvertToTextUnitEffect($e,1);$e.Timing.Duration=[single]$(if($i -eq 1){0.5}else{2})
   $motions=@($e.Behaviors|Where-Object{[int]$_.Type -eq 1});$rotations=@($e.Behaviors|Where-Object{[int]$_.Type -eq 4})
   if($motions.Count -ne 1 -or $rotations.Count -ne 4){throw 'Unexpected native Wave; stop instead of replacing it with a simple path'}
   $endpoint=if($i -eq 3){'-0.08'}else{'-0.04'};$motions[0].MotionEffect.Path=('M 0 0 L 0 '+$endpoint)
   $rows+=@{slide=$i;duration=$e.Timing.Duration;path=$motions[0].MotionEffect.Path;textUnit=[int]$e.EffectInformation.TextUnitEffect;rotationBehaviors=$rotations.Count}
  }
 }
 $p.SlideShowSettings.ShowWithAnimation=-1;$p.SaveAs($OutputDeck,24);$p.Close();$p=$null
 # UntilEnd belongs to the Wave parent. COM exposes no documented sentinel;
 # preserve all native motion/rotation/letter nodes and set only this attribute.
 for($i=1;$i -le 3;$i++){
  & $Officecli raw-set $OutputDeck ('/slide['+$i+']') --xpath "//p:cTn[@presetID='34' and @presetClass='emph']" --action setattr --xml 'repeatCount=indefinite'
  if($LASTEXITCODE -ne 0){throw 'Wave loop update failed'}
 }
 & $Officecli close $OutputDeck;if($LASTEXITCODE -ne 0){throw 'OfficeCLI close failed'}
 $p=$app.Presentations.Open($OutputDeck,0,0,-1);$p.Windows.Item(1).Activate();$p.Windows.Item(1).ViewType=9
 for($i=1;$i -le 4;$i++){$s=$p.Slides.Item($i);if($s.Shapes.Item('wave-target').TextFrame.TextRange.Text -ne 'DESIGN'){throw 'Editable text changed'};$s.Select();$s.Export((Join-Path $Native ('slide-'+$i+'.png')),'PNG',1600,900)}
 if($p.Slides.Item(4).TimeLine.MainSequence.Count -ne 0){throw 'Static control has animation'}
 $p.Slides.Item(1).Select();$p.Save();$rows|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $Native 'wave.json') -Encoding UTF8
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
