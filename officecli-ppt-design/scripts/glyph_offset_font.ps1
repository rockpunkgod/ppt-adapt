param([Parameter(Mandatory=$true)][string]$Output,[string]$Text='6',[string]$FontName='Arial Black',[double]$HeightCm=11.0)
$ErrorActionPreference='Stop'
if($Text -cne '6' -or $FontName -cne 'Arial Black' -or [Math]::Abs($HeightCm-11.0) -gt 0.000001){throw 'Validated fixture only: Text=6, FontName=Arial Black, HeightCm=11'}
if(Test-Path -LiteralPath $Output){throw 'Refusing overwrite'}
Add-Type -AssemblyName System.Drawing
$family=New-Object Drawing.FontFamily($FontName);$format=[Drawing.StringFormat]::GenericTypographic
$matrix=New-Object Drawing.Drawing2D.Matrix;$path=New-Object Drawing.Drawing2D.GraphicsPath
try{
 $path.FillMode=[Drawing.Drawing2D.FillMode]::Winding
 $path.AddString($Text,$family,0,[single]400,[Drawing.PointF]::new(0,0),$format)
 $bounds=$path.GetBounds();if($bounds.Height -le 0 -or $HeightCm -le 0){throw 'Positive height and a visible glyph required'}
 $path.Flatten($matrix,[single]0.1);$contours=New-Object 'System.Collections.Generic.List[object]';$current=$null
 $points=$path.PathPoints;$types=$path.PathTypes
 for($i=0;$i -lt $points.Length;$i++){
  $type=$types[$i] -band 7
  if($type -eq 0){$current=New-Object 'System.Collections.Generic.List[object]'}
  if($type -notin 0,1){throw 'Only flattened line segments expected'}
  $current.Add(@([double]$points[$i].X,[double]$points[$i].Y))
  if(($types[$i] -band 128) -ne 0){$contours.Add($current.ToArray());$current=$null}
 }
 if($null -ne $current){throw 'Glyph contour must close'}
 $result=@{method='Windows GDI+ font outline only, flatten 0.1 world unit';font=$FontName;text=$Text;unitsToCm=$HeightCm/$bounds.Height;bounds=@{x=$bounds.X;y=$bounds.Y;width=$bounds.Width;height=$bounds.Height};original=$contours.ToArray()}
 [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Output)))|Out-Null
 [IO.File]::WriteAllText([IO.Path]::GetFullPath($Output),($result|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
}finally{$path.Dispose();$family.Dispose();$matrix.Dispose();$format.Dispose()}
