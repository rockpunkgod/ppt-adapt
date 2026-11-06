param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Native,[Parameter(Mandatory=$true)][string]$Officecli)
$ErrorActionPreference='Stop';$OutputDeck=[IO.Path]::GetFullPath($OutputDeck);$Native=[IO.Path]::GetFullPath($Native)
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Native)){throw 'Fresh outputs required'};[void][IO.Directory]::CreateDirectory($Native)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$p.Windows.Item(1).Activate();$p.Windows.Item(1).ViewType=9
 if($p.Slides.Count -ne 2){throw 'Two-page wheel and control required'}
 for($i=1;$i -le 2;$i++){
  $s=$p.Slides.Item($i);$seq=$s.TimeLine.MainSequence;if($seq.Count -ne 0){throw 'Do not overwrite an existing timeline'}
  $e=$seq.AddEffect($s.Shapes.Item('rotor'),61,0,1);$e.EffectParameters.Amount=[single]360;$e.Timing.Duration=[single]1;$e.Timing.Accelerate=0;$e.Timing.Decelerate=0
  $f=$seq.AddEffect($s.Shapes.Item('result'),10,0,1);$f.Timing.Duration=[single]0.2
 }
 $p.SlideShowSettings.ShowWithAnimation=-1;$p.SaveAs($OutputDeck,24);$p.Close();$p=$null
 for($i=1;$i -le 2;$i++){
  & $Officecli raw-set $OutputDeck ('/slide['+$i+']') --xpath "//p:cTn[@presetID='8' and @presetClass='emph']" --action setattr --xml 'repeatCount=indefinite'
  if($LASTEXITCODE -ne 0){throw 'Spin repeat failed'}
 }
 # Exact condition authored by PowerPoint's UntilNextClick option on this machine.
 # Both options return the same COM RepeatCount; never infer semantics from that value.
 & $Officecli raw-set $OutputDeck '/slide[1]' --xpath "//p:cTn[@presetID='8' and @presetClass='emph']/p:childTnLst" --action insertbefore --xml "<p:endCondLst><p:cond evt='onNext' delay='0'><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>"
 if($LASTEXITCODE -ne 0){throw 'UntilNextClick condition failed'}
 & $Officecli close $OutputDeck;if($LASTEXITCODE -ne 0){throw 'OfficeCLI close failed'}
 $p=$app.Presentations.Open($OutputDeck,0,0,-1);$p.Windows.Item(1).Activate();$p.Windows.Item(1).ViewType=9
 for($i=1;$i -le 2;$i++){$s=$p.Slides.Item($i);if($s.TimeLine.MainSequence.Count -ne 2){throw 'Unexpected effects after native reopen'};$s.Select();$s.Export((Join-Path $Native ('slide-'+$i+'.png')),'PNG',1600,900)}
 $p.Slides.Item(1).Select();$p.Save()
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
