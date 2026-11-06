param([string]$InputDeck,[string]$OutputDeck,[string]$Evidence)
$ErrorActionPreference='Stop';if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'Fresh outputs required'};[void][IO.Directory]::CreateDirectory($Evidence)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$s=$p.Slides.Item(2);$s.Select()
 $mask=$s.Shapes.Item('portal-mask');$title=$s.Shapes.Item('portal-title');$copy=$title.Duplicate().Item(1);$copy.Name='portal-title-overlay';$copy.Left=$title.Left;$copy.Top=$title.Top
 $before=@{};for($i=1;$i -le $s.Shapes.Count;$i++){$before[[int]$s.Shapes.Item($i).Id]=$true}
 $s.Shapes.Range([object[]]@('portal-mask','portal-title')).MergeShapes(4,$mask)
 $result=$null;for($i=1;$i -le $s.Shapes.Count;$i++){$q=$s.Shapes.Item($i);if(-not $before.ContainsKey([int]$q.Id)){$result=$q}}
 if(-not $result){throw 'No native subtraction output'};$result.Name='portal-cutout';$result.Line.Visible=0
 $result.ZOrder(0);$copy.ZOrder(0);$s.Shapes.Item('portal-caption').ZOrder(0);$s.Shapes.Item('portal-video').ZOrder(1)
 # Prove the hollow geometry separately from the matching foreground title.
 $copy.Visible=0;$s.Shapes.Item('portal-caption').Visible=0;$s.Export((Join-Path $Evidence 'true-hole.png'),'PNG',1600,900);$copy.Visible=-1;$s.Shapes.Item('portal-caption').Visible=-1
 $seq=$s.TimeLine.MainSequence;$log=@()
 foreach($name in @('portal-caption','portal-title-overlay')){
  $e=$seq.AddEffect($s.Shapes.Item($name),2,0,2);$e.Exit=-1;$e.EffectParameters.Direction=1;$e.Timing.Duration=[single]1;$e.Timing.Accelerate=[single]1;$e.Timing.Decelerate=[single]0
  $log+=@{name=$name;effect=[int]$e.EffectType;exit=[int]$e.Exit;direction=[int]$e.EffectParameters.Direction;duration=$e.Timing.Duration;accelerate=$e.Timing.Accelerate}
 }
 $g=$seq.AddEffect($result,59,0,2);$default=$g.EffectParameters.Size;$size=if($default -gt 0 -and $default -le 3){90}else{9000};$g.EffectParameters.Size=[single]$size;$g.Timing.Duration=[single]2;$g.Timing.Accelerate=[single]1;$g.Timing.Decelerate=[single]0
 $log+=@{name='portal-cutout';effect=[int]$g.EffectType;defaultSize=$default;size=$g.EffectParameters.Size;duration=$g.Timing.Duration;accelerate=$g.Timing.Accelerate;nodes=$result.Nodes.Count;type=[int]$result.Type}
 foreach($i in @(1,2)){$p.Slides.Item($i).SlideShowTransition.EntryEffect=0;$p.Slides.Item($i).SlideShowTransition.AdvanceOnTime=0;$p.Slides.Item($i).SlideShowTransition.AdvanceOnClick=-1}
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24);$p.Slides.Item(1).Export((Join-Path $Evidence 'page-1.png'),'PNG',1600,900);$s.Export((Join-Path $Evidence 'initial.png'),'PNG',1600,900);$log|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'readback.json') -Encoding UTF8
}finally{if($p){try{$p.Close()}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
