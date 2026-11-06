param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop';if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'};[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$states=@()
try{
for($pass=1;$pass -le 2;$pass++){
$p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,0);$rows=@()
for($i=1;$i -le $p.Slides.Count;$i++){
if($pass -eq 2){$p.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('slide-'+$i+'.png')),'PNG',1600,900)}
foreach($shape in $p.Slides.Item($i).Shapes){if($shape.Fill.Type -eq 3){$stops=@();for($k=1;$k -le $shape.Fill.GradientStops.Count;$k++){$g=$shape.Fill.GradientStops.Item($k);$stops+=@{position=$g.Position;transparency=$g.Transparency;rgb=$g.Color.RGB}};$rows+=@{slide=$i;name=$shape.Name;type=[int]$shape.Type;fillType=[int]$shape.Fill.Type;stopCount=$shape.Fill.GradientStops.Count;stops=$stops}}}
}
$states+=@{pass=$pass;shapes=$rows};$p.Close();$p=$null
}
@{passed=$true;states=$states;scope='Read-only native reopen and real gradient objects; no media/animation claim'}|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $Evidence 'native-objects.json') -Encoding UTF8
}finally{if($p){$p.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
