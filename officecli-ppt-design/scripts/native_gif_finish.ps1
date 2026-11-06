param([Parameter(Mandatory=$true)][ValidateSet('source','intersect','save')][string]$Stage,[Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Native)
$ErrorActionPreference='Stop';$InputDeck=[IO.Path]::GetFullPath($InputDeck);$OutputDeck=[IO.Path]::GetFullPath($OutputDeck);$Native=[IO.Path]::GetFullPath($Native)
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Native)){throw 'Fresh outputs required'};[void][IO.Directory]::CreateDirectory($Native)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open($InputDeck,0,0,-1);$p.Windows.Item(1).Activate();$p.Windows.Item(1).ViewType=9
 if($Stage -eq 'source'){
  $s=$p.Slides.Item(1);$s.Select();$g=$s.Shapes.Range([object[]]@('square-0','square-1','square-2','square-3','square-4','phase-marker')).Group();$g.Name='animated-vector-texture'
  $spin=$s.TimeLine.MainSequence.AddEffect($g,61,0,2);$spin.EffectParameters.Amount=[single]90;$spin.Timing.Duration=[single]2;$spin.Timing.AutoReverse=-1;$spin.Timing.Accelerate=[single]0.5;$spin.Timing.Decelerate=[single]0.5
  $grow=$s.TimeLine.MainSequence.AddEffect($g,59,0,2);$grow.EffectParameters.Size=[single]130;$grow.Timing.Duration=[single]2;$grow.Timing.AutoReverse=-1;$grow.Timing.Accelerate=[single]0.5;$grow.Timing.Decelerate=[single]0.5
  $s.SlideShowTransition.EntryEffect=0;$s.SlideShowTransition.AdvanceOnTime=-1;$s.SlideShowTransition.AdvanceTime=[single]4
 }elseif($Stage -eq 'intersect'){
  $s=$p.Slides.Item(2);$s.Select();$pic=$s.Shapes.Item('texture-gif');$range=$s.Shapes.Range([object[]]@('texture-gif','vase-cut'));$range.Select();$range.MergeShapes(3,$pic)
  $matches=@();for($j=1;$j -le $s.Shapes.Count;$j++){$v=$s.Shapes.Item($j);if([int]$v.Type -eq 13 -and $v.Name -ne 'bare-gif'){$matches+=,$v}}
  if($matches.Count -ne 1){throw 'Exactly one intersected GIF picture required'};$matches[0].Name='texture-gif'
 }
 $p.SlideShowSettings.ShowWithAnimation=-1;$p.Slides.Item(1).Select();$p.SaveAs($OutputDeck,24)
 for($i=1;$i -le $p.Slides.Count;$i++){$p.Slides.Item($i).Export((Join-Path $Native ('slide-'+$i+'.png')),'PNG',1600,900)}
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
