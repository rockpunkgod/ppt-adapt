param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Native)
$ErrorActionPreference='Stop'
$OutputDeck=[IO.Path]::GetFullPath($OutputDeck);$Native=[IO.Path]::GetFullPath($Native)
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Native)){throw 'Fresh output paths required'}
[void][IO.Directory]::CreateDirectory($Native)
$a=New-Object -ComObject PowerPoint.Application;$p=$null;$log=@()
try{
 $p=$a.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 for($i=1;$i -le 2;$i++){
  $s=$p.Slides.Item($i);$g=$s.Shapes.Item('!!turning-page');$td=$g.ThreeD
  $beforeWidth=[single]$g.Width;$beforeLeft=[single]$g.Left
  for($j=1;$j -le $g.GroupItems.Count;$j++){$child=$g.GroupItems.Item($j);if($child.Name -like '*axis-left-helper'){$child.Flip(0);$child.Visible=0}}
  $td.SetPresetCamera(62);$td.FieldOfView=[single]60;$td.RotationX=[single]$(if($i -eq 1){0}else{180});$td.RotationY=0;$td.RotationZ=0;$td.PresetLighting=18
  $s.SlideShowTransition.Duration=[single]2;$s.SlideShowTransition.AdvanceOnClick=-1;$s.SlideShowTransition.AdvanceOnTime=0
  if([Math]::Abs([double]$g.Width-$beforeWidth) -gt 0.01 -or [Math]::Abs([double]$g.Left-$beforeLeft) -gt 0.01){throw 'Hiding helper changed hinge bounds'}
  $log+=@{slide=$i;left=$g.Left;top=$g.Top;width=$g.Width;height=$g.Height;rotationX=$td.RotationX;rotationY=$td.RotationY;fov=$td.FieldOfView;lighting=$td.PresetLighting;children=$g.GroupItems.Count;helperVisibility=$g.GroupItems.Item(2).Visible;boundsKept=$true}
  $s.Export((Join-Path $Native ('slide-'+$i+'.png')),'PNG',1600,900)
 }
 $p.SaveAs($OutputDeck,24)
 [IO.File]::WriteAllText((Join-Path $Native 'readback.json'),($log|ConvertTo-Json -Depth 5),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)}
