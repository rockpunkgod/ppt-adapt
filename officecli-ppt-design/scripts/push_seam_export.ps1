param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$before=(Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,0)
 if($p.Slides.Count -ne 3){throw 'Expected three route pages'}
 for($i=1;$i -le 3;$i++){$s=$p.Slides.Item($i);$s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('page-'+$i+'.png')),'PNG',1600,900)}
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
if((Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash -ne $before){throw 'Read-only export changed original'}
