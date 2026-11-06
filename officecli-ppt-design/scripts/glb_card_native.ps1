param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,0)
 if($p.Slides.Count -ne 3){throw 'Expected three card states'}
 $rows=@()
 for($i=1;$i -le $p.Slides.Count;$i++){
  $s=$p.Slides.Item($i);$model=$s.Shapes.Item('!!dual-card')
  if([int]$model.Type -ne 30){throw 'Not a native PowerPoint model3D'}
  $expected=$(if($i -eq 2){180}else{0})
  if([Math]::Abs([double]$model.Model3D.RotationY-$expected) -gt 0.01){throw 'Native 3D rotation mismatch'}
  $rows+=@{slide=$i;name=$model.Name;type=[int]$model.Type;rotationY=$model.Model3D.RotationY}
  $s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('slide-'+$i+'.png')),'PNG',1600,900)
  $s.SlideShowTransition.AdvanceOnTime=-1;$s.SlideShowTransition.AdvanceTime=1
 }
 $rows|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $Evidence 'native-objects.json') -Encoding UTF8
 $p.CreateVideo((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'native.mp4'),-1,1,720,30,85)
 $end=[DateTime]::UtcNow.AddSeconds(18)
 do{Start-Sleep -Milliseconds 250;$st=$p.CreateVideoStatus;if([DateTime]::UtcNow -gt $end){throw 'Video timeout'}}while($st -in 1,2)
 if($st -ne 3){throw "Video status $st"}
}finally{if($p){$p.Saved=-1;$p.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
