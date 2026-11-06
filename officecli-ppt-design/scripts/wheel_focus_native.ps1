param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $OutputDirectory){throw 'Refusing existing evidence directory'}
[IO.Directory]::CreateDirectory($OutputDirectory)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1)
 foreach($i in @(1,2)){$p.Slides.Item($i).Export((Join-Path $OutputDirectory "slide-$i.png"),'PNG',1600,900)}
 $movie=Join-Path $OutputDirectory 'native.mp4'
 $p.CreateVideo($movie,$true,5,720,30,75)
 $limit=[DateTime]::UtcNow.AddSeconds(55)
 while([int]$p.CreateVideoStatus -in @(0,1,2) -and [DateTime]::UtcNow -lt $limit){Start-Sleep -Milliseconds 250}
 if([int]$p.CreateVideoStatus -ne 3){throw ('Native video status '+$p.CreateVideoStatus)}
 [IO.File]::WriteAllText((Join-Path $OutputDirectory 'native.json'),(@{status=3;file=[IO.Path]::GetFullPath($Deck);movie=$movie}|ConvertTo-Json),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
