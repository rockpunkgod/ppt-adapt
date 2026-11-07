param([string]$InputDeck,[string]$OutputDirectory,[switch]$Video)
$ErrorActionPreference='Stop'
$InputDeck=[IO.Path]::GetFullPath($InputDeck)
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(-not(Test-Path -LiteralPath $InputDeck)){throw ('Input missing: '+$InputDeck)}
foreach($reserved in @('preview.jpg','.preview-pages')+$(if($Video){@('preview.mp4')}else{@()})){
 if(Test-Path -LiteralPath (Join-Path $OutputDirectory $reserved)){throw ('Refusing to overwrite preview output: '+$reserved)}
}
Add-Type -AssemblyName System.Drawing
$app=New-Object -ComObject PowerPoint.Application
$p=$null;$canvas=$null;$g=$null
try {
 $p=$app.Presentations.Open($InputDeck,-1,0,0)
 $count=$p.Slides.Count
 $canvas=New-Object Drawing.Bitmap(1600,([int][Math]::Ceiling($count/2.0)*450))
 $g=[Drawing.Graphics]::FromImage($canvas);$g.Clear([Drawing.Color]::White)
 $pageDir=Join-Path $OutputDirectory '.preview-pages'
 [void][IO.Directory]::CreateDirectory($pageDir)
 for($i=1;$i -le $count;$i++){
  $path=Join-Path $pageDir ('slide-'+$i+'.png')
  $p.Slides.Item($i).Export($path,'PNG',1600,900)
  $img=[Drawing.Image]::FromFile($path)
  try {$g.DrawImage($img,(($i-1)%2)*800,([int][Math]::Floor(($i-1)/2.0))*450,800,450)}finally{$img.Dispose()}
 }
 $canvas.Save((Join-Path $OutputDirectory 'preview.jpg'),[Drawing.Imaging.ImageFormat]::Jpeg)
 if($Video){
  $videoPath=Join-Path $OutputDirectory 'preview.mp4'
  $p.CreateVideo($videoPath,$false,5,720,24,80)
  $deadline=[DateTime]::UtcNow.AddMinutes(8)
  do {
   Start-Sleep -Milliseconds 750
   $status=[int]$p.CreateVideoStatus
   if([DateTime]::UtcNow -gt $deadline){throw 'PowerPoint video export timed out.'}
  } while($status -in @(1,2))
  if($status -ne 3 -or -not(Test-Path -LiteralPath $videoPath)){throw "PowerPoint video export failed: $status"}
 }
} finally {
 if($g){$g.Dispose()};if($canvas){$canvas.Dispose()}
 if($p){$p.Close()}
}
