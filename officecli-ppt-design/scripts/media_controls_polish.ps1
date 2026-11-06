param([string]$InputDeck,[string]$OutputDeck,[string]$Evidence)
$ErrorActionPreference='Stop';if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'Fresh output/evidence required'};[void][IO.Directory]::CreateDirectory($Evidence)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
function Get-MediaControlsInfo($shape){$m=$shape.MediaFormat;$books=@();for($i=1;$i -le $m.MediaBookmarks.Count;$i++){$b=$m.MediaBookmarks.Item($i);$books+=@{name=$b.Name;position=$b.Position}};return @{name=$shape.Name;id=$shape.Id;mediaType=[int]$shape.MediaType;length=$m.Length;start=$m.StartPoint;end=$m.EndPoint;fadeIn=$m.FadeInDuration;fadeOut=$m.FadeOutDuration;bookmarks=$books}}
try{
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),-1,0,-1);$p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 $before=@();foreach($i in @(1,2,3)){$name=@('trim-video','fade-audio','bookmark-video')[$i-1];$before+=Get-MediaControlsInfo ($p.Slides.Item($i).Shapes.Item($name))};$before|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'before.json') -Encoding UTF8
 # Trim is authored by OfficeCLI. Never repair it silently here.
 foreach($i in @(1,2)){if($before[$i-1].start -ne 2000 -or $before[$i-1].end -ne 10000 -or $before[$i-1].length -ne 12000){throw 'CLI retained trim failed native readback'}}
 $audio=$p.Slides.Item(2).Shapes.Item('fade-audio');$audio.MediaFormat.FadeInDuration=2000;$audio.MediaFormat.FadeOutDuration=2000
 $video=$p.Slides.Item(3).Shapes.Item('bookmark-video');[void]$video.MediaFormat.MediaBookmarks.Add(6000,'source-six');[void]$video.MediaFormat.MediaBookmarks.Add(9000,'source-nine')
 $p.Save();$after=@();foreach($i in @(1,2,3)){$name=@('trim-video','fade-audio','bookmark-video')[$i-1];$after+=Get-MediaControlsInfo ($p.Slides.Item($i).Shapes.Item($name));$p.Slides.Item($i).Export([IO.Path]::GetFullPath((Join-Path $Evidence ('slide-'+$i+'.png'))),'PNG',1600,900)};$after|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'after.json') -Encoding UTF8
}catch{@{error=$_.Exception.ToString()}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'error.json') -Encoding UTF8;throw}
finally{if($p){try{$p.Close()}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
