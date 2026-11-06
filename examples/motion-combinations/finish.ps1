param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$PreviewDirectory,[double]$Angle=-10,[double]$Duration=0.5,[double]$TopDelay=0.25,[double]$BottomDelay=0.4,[double]$TopDistanceCm=0.8,[double]$BottomDistanceCm=0.3,[double]$SeparationCm=0.1)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $PreviewDirectory)){throw 'Fresh output paths required'}
if($Duration -le 0 -or $TopDelay -lt 0 -or $BottomDelay -lt 0 -or [Math]::Abs($Angle) -gt 30 -or $TopDistanceCm -lt 0 -or $BottomDistanceCm -lt 0 -or $SeparationCm -lt 0){throw 'Invalid slice timing, distances or angle'}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($PreviewDirectory))
function Merge($slide,[string[]]$names,[int]$mode,$primary,[string]$name){
 $ids=@{};for($k=1;$k -le $slide.Shapes.Count;$k++){$ids[[int]$slide.Shapes.Item($k).Id]=$true}
 $range=$slide.Shapes.Range([object[]]$names);$range.Select();$range.MergeShapes($mode,$primary)
 for($k=1;$k -le $slide.Shapes.Count;$k++){$q=$slide.Shapes.Item($k);if(-not $ids.ContainsKey([int]$q.Id)){$q.Name=$name;return $q}}
 throw 'No new geometry after merge'
}
function PathEffect($seq,$shape,[string]$path,[double]$seconds,[double]$delay,[int]$trigger){
 $e=$seq.AddEffect($shape,127,0,$trigger);$e.Behaviors.Item(1).MotionEffect.Path=$path
 $e.Timing.Duration=[single]$seconds;$e.Timing.TriggerDelayTime=[single]$delay
 $e.Timing.Accelerate=[single]0.5;$e.Timing.Decelerate=[single]0.5;return $e
}
function Named($slide,[string]$name){
 for($j=1;$j -le $slide.Shapes.Count;$j++){$q=$slide.Shapes.Item($j);if($q.Name -eq $name -or $q.Name -eq ('!!'+$name)){return $q}}
 throw ('Shape not found: '+$name)
}
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 $p.Windows.Item(1).Activate();$p.Windows.Item(1).ViewType=9
 $s=$p.Slides.Item(1);$s.Select();$seed=$s.Shapes.Item('slice-source');$w=[double]$p.PageSetup.SlideWidth;$h=[double]$p.PageSetup.SlideHeight
 $cutY=[double]$seed.TextFrame.TextRange.BoundTop+[double]$seed.TextFrame.TextRange.BoundHeight/2
 $r=$Angle*[Math]::PI/180;$clipH=$h*2;$clipW=$w*3
 $cx=$w/2-[Math]::Sin($r)*$clipH/2;$cy=$cutY+[Math]::Cos($r)*$clipH/2
 foreach($part in @('top','bottom')){
  $copy=$seed.Duplicate().Item(1);$copy.Name='copy-'+$part;$copy.Left=$seed.Left;$copy.Top=$seed.Top
  $clip=$s.Shapes.AddShape(1,[single]($cx-$clipW/2),[single]($cy-$clipH/2),[single]$clipW,[single]$clipH);$clip.Rotation=[single](($Angle+360)%360);$clip.Name='clip-'+$part
  $mode=if($part -eq 'top'){4}else{3};$piece=Merge $s @($copy.Name,$clip.Name) $mode $copy ('slice-'+$part)
  $piece.Fill.Visible=-1;$piece.Fill.Solid();$piece.Fill.ForeColor.RGB=15528946
  $piece.Line.Visible=-1;$piece.Line.ForeColor.RGB=15528946;$piece.Line.Weight=[single]0.15
 }
 $seed.Visible=0
 $dx=$w+170;$x=-85;$y=$cutY+[Math]::Tan($r)*($x-$w/2)
 $flash=$s.Shapes.AddShape(4,[single]($x-15),[single]($y-1.5),30,3);$flash.Name='slice-flash';$flash.Rotation=[single](($Angle+360)%360);$flash.Fill.ForeColor.RGB=7925722;$flash.Line.Visible=0
 $seq=$s.TimeLine.MainSequence
 $e=PathEffect $seq $flash ('M 0 0 L '+($dx/$w).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)+' '+([Math]::Tan($r)*$dx/$h).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)+' E') $Duration 0 1
 foreach($row in @(@{name='slice-top';distance=$TopDistanceCm;normal=-$SeparationCm;delay=$TopDelay},@{name='slice-bottom';distance=$BottomDistanceCm;normal=$SeparationCm;delay=$BottomDelay})){
  $px=([Math]::Cos($r)*$row.distance-[Math]::Sin($r)*$row.normal)*72/2.54
  $py=([Math]::Sin($r)*$row.distance+[Math]::Cos($r)*$row.normal)*72/2.54
  $path='M 0 0 L '+($px/$w).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)+' '+($py/$h).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)+' E'
  $e=PathEffect $seq $s.Shapes.Item($row.name) $path $Duration $row.delay 2
 }
 $s=$p.Slides.Item(2);$s.Select();$card=$s.Shapes.Item('working-card');$seq=$s.TimeLine.MainSequence
 $e=$seq.AddEffect($card,10,0,1);$e.Timing.Duration=[single]0.4
 $e=PathEffect $seq $card 'M 0 0 L 0.47 0 E' 2 0.4 2
 $e=$seq.AddEffect($card,59,0,2);$sz=$e.EffectParameters.Size;$e.EffectParameters.Size=[single]$(if($sz -le 3){1.15}else{115});$e.Timing.Duration=[single]2;$e.Timing.TriggerDelayTime=[single]0.4
 $e=$seq.AddEffect($card,61,0,2);$e.EffectParameters.Amount=[single]15;$e.Timing.Duration=[single]2;$e.Timing.TriggerDelayTime=[single]0.4
 $e=$seq.AddEffect($card,10,0,1);$e.Exit=-1;$e.Timing.Duration=[single]0.4
 for($i=3;$i -le 4;$i++){
  $s=$p.Slides.Item($i);$s.SlideShowTransition.AdvanceOnClick=0;$s.SlideShowTransition.AdvanceOnTime=0
  foreach($q in @(1,2)){$target=$p.Slides.Item($q+2);$nav=Named $s ('nav-q'+$q);$action=$nav.ActionSettings.Item(1);$action.Action=7;$action.Hyperlink.SubAddress=([string]$target.SlideID)+','+([string]$target.SlideIndex)+',Quarter'}
  # Zoom was added after the mask in OfficeCLI. Correct the layer order before scrolling.
  foreach($n in @('fixed-mask','fixed-title','fixed-label','row-0','row-1','row-2','nav-q1','nav-q2','hint')){$q=Named $s $n;$q.ZOrder(0)}
 }
 $p.SlideShowSettings.ShowPresenterView=0;$p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 for($i=1;$i -le $p.Slides.Count;$i++){$p.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($PreviewDirectory)) ('slide-'+$i+'.png')),'PNG',1600,900)}
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
