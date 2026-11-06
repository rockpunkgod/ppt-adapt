param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop';if(Test-Path -LiteralPath $Evidence){throw 'Fresh evidence required'};[void][IO.Directory]::CreateDirectory($Evidence)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{$app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1);foreach($i in @(1,2,3)){$p.Slides.Item($i).Export([IO.Path]::GetFullPath((Join-Path $Evidence ('slide-'+$i+'.png'))),'PNG',1600,900)}}finally{if($p){try{$p.Close()}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
