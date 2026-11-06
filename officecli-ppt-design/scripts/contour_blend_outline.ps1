param([string]$Output,[string]$Text='P',[string]$FontName='Arial Black',[double]$HeightCm=5.0)
$ErrorActionPreference='Stop'
if($Text -ne 'P' -or $FontName -ne 'Arial Black' -or $HeightCm -ne 5){throw 'Only validated default P / Arial Black / 5cm accepted'}
if(Test-Path -LiteralPath $Output){throw 'Refusing overwrite'}
Add-Type -AssemblyName System.Drawing
function Contours($path){
 $items=New-Object 'System.Collections.Generic.List[object]'
 $current=$null;$points=$path.PathPoints;$types=$path.PathTypes
 for($i=0;$i -lt $points.Length;$i++){
  if(($types[$i] -band 7) -eq 0){$current=New-Object 'System.Collections.Generic.List[object]'}
  if(($types[$i] -band 7) -notin 0,1){throw 'Expected flattened line segments'}
  $current.Add(@([double]$points[$i].X,[double]$points[$i].Y))
  if(($types[$i] -band 128) -ne 0){$items.Add($current.ToArray());$current=$null}
 }
 if($null -ne $current){throw 'Open glyph contour unexpected'}
 return ,$items.ToArray()
}
$family=New-Object Drawing.FontFamily($FontName)
$format=[Drawing.StringFormat]::GenericTypographic
$matrix=New-Object Drawing.Drawing2D.Matrix
$path=New-Object Drawing.Drawing2D.GraphicsPath
try{
 $path.FillMode=[Drawing.Drawing2D.FillMode]::Winding
 $path.AddString($Text,$family,0,[single]400,[Drawing.PointF]::new(0,0),$format)
 $bounds=$path.GetBounds();if($bounds.Height -le 0 -or $HeightCm -le 0){throw 'Nonempty glyph and positive target height required'};$unitsToCm=$HeightCm/$bounds.Height
 $path.Flatten($matrix,[single]0.1)
 $original=Contours $path;$strokes=@()
 $result=@{method='Windows GDI+ original glyph contour only, flattened at 0.1 world unit; no Widen geometry used';font=$FontName;text=$Text;unitsToCm=$unitsToCm;bounds=@{x=$bounds.X;y=$bounds.Y;width=$bounds.Width;height=$bounds.Height};original=$original;strokes=$strokes}
 $json=$result|ConvertTo-Json -Depth 12
 [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Output)))|Out-Null
 [IO.File]::WriteAllText([IO.Path]::GetFullPath($Output),$json,[Text.UTF8Encoding]::new($false))
}finally{$path.Dispose();$family.Dispose();$matrix.Dispose();$format.Dispose()}
