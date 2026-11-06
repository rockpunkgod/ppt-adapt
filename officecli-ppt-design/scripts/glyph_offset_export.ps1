param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$before=(Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,0)
 if($p.Slides.Count -ne 5){throw 'Five glyph states expected'}
 $rows=@()
 for($i=1;$i -le 5;$i++){
  $name=$(if($i -eq 1){'original-glyph'}elseif($i -eq 5){'outlined-stroke'}else{'offset-result'})
  $shape=$p.Slides.Item($i).Shapes.Item($name)
  if($shape.Type -ne 5 -or $shape.Line.Visible -ne 0){throw 'True freeform fill without line styling required'}
  $rows+=@{page=$i;type=[int]$shape.Type;nodes=$shape.Nodes.Count;lineVisible=[int]$shape.Line.Visible}
  $p.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('page-'+$i+'.png')),'PNG',1600,900)
 }
 $rows|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'native.json') -Encoding UTF8
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
if((Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash -ne $before){throw 'Read-only export changed source'}
