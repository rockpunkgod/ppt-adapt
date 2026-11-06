param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck)
$ErrorActionPreference='Stop'
$OutputDeck=[IO.Path]::GetFullPath($OutputDeck)
if(Test-Path -LiteralPath $OutputDeck){throw 'Fresh output required'}
Add-Type -AssemblyName System.Drawing
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 if($p.Slides.Count -ne 3){throw 'Expected three focus states'}
 for($i=1;$i -le 3;$i++){
  $s=$p.Slides.Item($i);$s.Select()
  $picture=$s.Shapes.Item('focus-image');$circle=$s.Shapes.Item('!!focus-circle')
  $s.Shapes.Range(@('focus-image','!!focus-circle')).MergeShapes(3,$picture)
  $crop=$null;foreach($sh in $s.Shapes){if($sh.Name -notin @('focus-image','focus-circle') -and $sh.Type -eq 13 -and $sh.Width -lt 200){$crop=$sh}}
  if(-not $crop){throw 'Native picture-circle intersection missing'}
  $crop.Name='!!focus-image'
  $x=@(7.62,16.94,26.25)[$i-1];$dx=$x-16.94;$dy=8.47-14.35;$len=[Math]::Sqrt($dx*$dx+$dy*$dy);$angle=[Math]::Atan2($dx,-$dy)*180/[Math]::PI
  $beam=$s.Shapes.Item('!!light-beam');$beam.Rotation=$angle
  $beam.Left=(16.94+$dx/2-2.1)*72/2.54;$beam.Top=(14.35+$dy/2-$len/2)*72/2.54
  # The lamp stays above the beam; text stays readable above image layers.
  if($s.Shapes.Item('!!lamp').Type -ne 30){throw 'Native model required'}
  $s.Shapes.Item('!!lamp').ZOrder(0)
  # Original fixture has a uniquely orange lens. Align its native projected
  # centroid with the beam apex; 3D auto-fit changes with model orientation.
  $anchor=Join-Path ([IO.Path]::GetDirectoryName($OutputDeck)) ('anchor-'+$i+'.png')
  $s.Export($anchor,'PNG',1600,900);$bmp=New-Object Drawing.Bitmap($anchor)
  try{
   $nx=0.0;$ny=0.0;$count=0
   for($yy=590;$yy -lt 850;$yy++){for($xx=580;$xx -lt 1020;$xx++){
    $c=$bmp.GetPixel($xx,$yy)
    if($c.R -gt 110 -and $c.R -gt 1.3*$c.G -and $c.B -lt .7*$c.G){$nx+=$xx;$ny+=$yy;$count++}
   }}
   if($count -lt 8){throw 'Original orange lens not visible; do not infer an arbitrary model anchor'}
   $lamp=$s.Shapes.Item('!!lamp');$lamp.Left+=(16.94/33.87*1600-$nx/$count)*$p.PageSetup.SlideWidth/1600
   $lamp.Top+=(14.35/19.05*900-$ny/$count)*$p.PageSetup.SlideHeight/900
  }finally{$bmp.Dispose()}
 }
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
}finally{if($p){$p.Close()}}
# Shared PowerPoint remains running.
