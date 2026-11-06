param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$Model,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Native)
$ErrorActionPreference='Stop';$OutputDeck=[IO.Path]::GetFullPath($OutputDeck);$Native=[IO.Path]::GetFullPath($Native)
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Native)){throw 'Fresh outputs required'};[void][IO.Directory]::CreateDirectory($Native)
$m=Get-Content -LiteralPath $Model -Raw|ConvertFrom-Json;$app=New-Object -ComObject PowerPoint.Application;$p=$null;$rows=@()
function Replay($s,$ink){$matches=@();for($j=1;$j -le $s.TimeLine.MainSequence.Count;$j++){$e=$s.TimeLine.MainSequence.Item($j);if($e.Shape.Id -eq $ink.Id){$matches+=,$e}};if($matches.Count -ne 1){throw 'Keep the unique native Replay; do not rebuild EffectType0'};return $matches[0]}
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$p.Windows.Item(1).Activate();$p.Windows.Item(1).ViewType=9;$s=$p.Slides.Item(1);$s.Select();$ink=$s.Shapes.Item('original-native-ink');$pen=$s.Shapes.Item('tip-centered-pen');if([int]$ink.Type -ne 23){throw 'True native Ink required'}
 $replay=Replay $s $ink
 for($j=$s.TimeLine.MainSequence.Count;$j -ge 1;$j--){if($s.TimeLine.MainSequence.Item($j).Shape.Id -ne $ink.Id){$s.TimeLine.MainSequence.Item($j).Delete()}}
 $s.Shapes.Item('audio').Delete();$pen.ZOrder(0);$replay.Timing.TriggerType=1;$replay.Timing.Duration=[single]5;$replay.Timing.Accelerate=0;$replay.Timing.Decelerate=0;$replay.Timing.TriggerDelayTime=0
 $path=$s.TimeLine.MainSequence.AddEffect($pen,0,0,2);$motion=$path.Behaviors.Add(1);$motion.MotionEffect.Path=[string]$m.path;$path.Timing.Duration=[single]5;$path.Timing.Accelerate=0;$path.Timing.Decelerate=0;$path.Timing.TriggerDelayTime=0;$path.Timing.RewindAtEnd=0
 $fast=$s.Duplicate().Item(1);$fast.Select();$ink=$fast.Shapes.Item('original-native-ink');$pen=$fast.Shapes.Item('tip-centered-pen');$seq=$fast.TimeLine.MainSequence;$replay=Replay $fast $ink
 $entry=$seq.AddEffect($pen,2,0,1);$entry.EffectParameters.Direction=3;$entry.Timing.Duration=[single]0.5;$entry.Timing.Accelerate=0;$entry.Timing.Decelerate=[single]1;$entry.MoveBefore($replay)
 $replay=Replay $fast $ink;$replay.Timing.TriggerType=3;$replay.Timing.Duration=[single]1;$replay.Timing.Accelerate=0;$replay.Timing.Decelerate=0
 for($j=1;$j -le $seq.Count;$j++){$e=$seq.Item($j);if($e.Shape.Id -eq $pen.Id -and [int]$e.EffectType -eq 0){$e.Timing.Duration=[single]1;$e.Timing.TriggerType=2;$e.Timing.Accelerate=0;$e.Timing.Decelerate=0}}
 $exit=$seq.AddEffect($pen,2,0,3);$exit.Exit=-1;$exit.EffectParameters.Direction=4;$exit.Timing.Duration=[single]0.5;$exit.Timing.Accelerate=[single]1;$exit.Timing.Decelerate=0;$exit.Timing.RewindAtEnd=0
 $fast.Shapes.Item('title').TextFrame.TextRange.Text='写下清晰的路径。';$fast.Shapes.Item('footer').TextFrame.TextRange.Text='单击：笔尖入场，1秒同步绘制，随后向左离场。'
 # Native links revisit the completed animation state. A zero-time hidden gate
 # immediately before each content page makes the subsequent forward entry reset it.
 $gateSlow=$p.Slides.Add(1,12);$gateFast=$p.Slides.Add(3,12)
 foreach($gate in @($gateSlow,$gateFast)){$gate.FollowMasterBackground=0;$gate.Background.Fill.Solid();$gate.Background.Fill.ForeColor.RGB=1901846;$gate.SlideShowTransition.Hidden=-1;$gate.SlideShowTransition.AdvanceOnClick=0;$gate.SlideShowTransition.AdvanceOnTime=-1;$gate.SlideShowTransition.AdvanceTime=0}
 for($i=2;$i -le 4;$i+=2){
  $s=$p.Slides.Item($i);$s.Select();$s.SlideShowTransition.AdvanceOnClick=-1;$s.SlideShowTransition.AdvanceOnTime=0
  foreach($nav in @(@{shape='nav-slow';target=1},@{shape='nav-fast';target=3})){
   $label=$s.Shapes.Item($nav.shape);$label.TextFrame.TextRange.Font.Color.RGB=if($i -eq ($nav.target+1)){7925722}else{12306364}
   $hit=$s.Shapes.AddShape(1,$label.Left,$label.Top,$label.Width,$label.Height);$hit.Name=$nav.shape+'-hit';$hit.Fill.Solid();$hit.Fill.Transparency=[single]1;$hit.Line.Visible=0;$act=$hit.ActionSettings.Item(1);$act.Action=7;$target=$p.Slides.Item($nav.target);$act.Hyperlink.SubAddress=([string]$target.SlideID)+','+([string]$nav.target)+','
  }
  for($j=1;$j -le $s.TimeLine.MainSequence.Count;$j++){$e=$s.TimeLine.MainSequence.Item($j);$rows+=@{slide=$i;order=$j;type=[int]$e.EffectType;shape=$e.Shape.Name;duration=$e.Timing.Duration;trigger=[int]$e.Timing.TriggerType;accelerate=$e.Timing.Accelerate;decelerate=$e.Timing.Decelerate;exit=[int]$e.Exit}}
  $s.Export((Join-Path $Native ('slide-'+$i+'.png')),'PNG',1600,900)
 }
 $p.SlideShowSettings.ShowWithAnimation=-1;$p.SlideShowSettings.AdvanceMode=2;$p.Slides.Item(1).Select();$p.SaveAs($OutputDeck,24);$rows|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $Native 'timeline.json') -Encoding utf8
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
