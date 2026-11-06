param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Native)
$ErrorActionPreference='Stop'
$OutputDeck=[IO.Path]::GetFullPath($OutputDeck);$Native=[IO.Path]::GetFullPath($Native)
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Native)){throw 'Fresh outputs required'}
[void][IO.Directory]::CreateDirectory($Native)
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$log=@()
try{
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$p.Windows.Item(1).Activate();$p.Windows.Item(1).ViewType=9
 for($i=1;$i -le 2;$i++){
  $s=$p.Slides.Item($i);$s.Select();$cut=$s.Shapes.Item('flask-cut')
  $outline=$cut.Duplicate().Item(1);$outline.Name='flask-outline';$outline.Left=$cut.Left;$outline.Top=$cut.Top;$outline.Fill.Visible=0;$outline.Line.Visible=-1;$outline.Line.ForeColor.RGB=15528946;$outline.Line.Weight=[single]2
  $mask=$s.Shapes.AddShape(1,0,0,$p.PageSetup.SlideWidth,$p.PageSetup.SlideHeight);$mask.Name='mask-seed';$mask.Line.Visible=0
  $ids=@{};for($j=1;$j -le $s.Shapes.Count;$j++){$ids[[int]$s.Shapes.Item($j).Id]=$true}
  $range=$s.Shapes.Range([object[]]@('mask-seed','flask-cut'));$range.Select();$range.MergeShapes(4,$mask)
  $hole=$null;for($j=1;$j -le $s.Shapes.Count;$j++){$q=$s.Shapes.Item($j);if(-not $ids.ContainsKey([int]$q.Id)){$hole=$q}}
  if(-not $hole){throw 'No native subtract result'};$hole.Name='flask-hole-mask';$hole.Fill.Background();$hole.Line.Visible=0
  foreach($name in @('flask-outline','lip','kicker','title','body','caption')){$s.Shapes.Item($name).ZOrder(0)}
  $seq=$s.TimeLine.MainSequence
  if($i -eq 1){
   $e=$seq.AddEffect($s.Shapes.Item('water'),2,0,1);$e.EffectParameters.Direction=11;$e.Timing.Duration=[single]2;$e.Timing.Accelerate=0;$e.Timing.Decelerate=0
   $e.Timing.BounceEnd=-1;$e.Timing.BounceEndIntensity=[single]0.75
   for($b=1;$b -le $e.Behaviors.Count;$b++){if([int]$e.Behaviors.Item($b).Type -eq 5){$e.Behaviors.Item($b).Timing.BounceEnd=-1;$e.Behaviors.Item($b).Timing.BounceEndIntensity=[single]0.75}}
   $log+=@{slide=1;duration=$e.Timing.Duration;direction=$e.EffectParameters.Direction;bounce=$e.Timing.BounceEnd;bounceIntensity=$e.Timing.BounceEndIntensity}
  }else{
   foreach($row in @(@{name='wave-back';x=3;y=-8.2;trigger=1},@{name='wave-front';x=-7.6;y=-8.5;trigger=2})){
    $e=$seq.AddEffect($s.Shapes.Item($row.name),127,0,$row.trigger)
    $path='M 0 0 L '+([double]$row.x*72/2.54/$p.PageSetup.SlideWidth).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)+' '+([double]$row.y*72/2.54/$p.PageSetup.SlideHeight).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)+' E'
    $e.Behaviors.Item(1).MotionEffect.Path=$path;$e.Timing.Duration=[single]3;$e.Timing.Accelerate=[single]0.3;$e.Timing.Decelerate=[single]0.7
    $log+=@{slide=2;shape=$row.name;path=$path;duration=$e.Timing.Duration;trigger=$e.Timing.TriggerType}
   }
  }
  $log+=@{slide=$i;maskFillType=$hole.Fill.Type;maskType=$hole.Type;outlineFill=$outline.Fill.Visible;outlineWidth=$outline.Line.Weight}
  $s.Export((Join-Path $Native ('slide-'+$i+'.png')),'PNG',1600,900)
 }
 $p.SaveAs($OutputDeck,24)
 [IO.File]::WriteAllText((Join-Path $Native 'readback.json'),($log|ConvertTo-Json -Depth 5),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
