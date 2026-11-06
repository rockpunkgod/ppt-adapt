param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1)
 if($p.Slides.Count -ne 2){throw 'Two native chart slides required'}
 for($i=1;$i -le 2;$i++){
  $s=$p.Slides.Item($i);$g=$s.Shapes.Item('sales-chart')
  if($g.HasChart -ne -1 -or $g.Type -ne 3){throw 'A real native chart shape is required'}
  $s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('after-'+$i+'.png')),'PNG',1600,900)
 }
 @{readOnly=$true;chartShapes=2;chartDataApiCalled=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'native-audit.json') -Encoding UTF8
}finally{if($p){$p.Close()};[Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null}
