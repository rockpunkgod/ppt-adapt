param([string]$InputDeck,[string]$OutputDeck,[string]$Evidence)
$ErrorActionPreference='Stop';if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'Fresh paths required'};[void][IO.Directory]::CreateDirectory($Evidence)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),-1,0,-1);$p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24);$s=$p.Slides.Item(2);$m=$s.Shapes.Item('fade-audio').MediaFormat
 $before=@{length=$m.Length;start=$m.StartPoint;end=$m.EndPoint;fadeIn=$m.FadeInDuration;fadeOut=$m.FadeOutDuration};$m.FadeInDuration=0;$m.FadeOutDuration=0
 $s.Shapes.Item('title').TextFrame.TextRange.Text='Matched native audio control | fades disabled';$s.Shapes.Item('fade-label').TextFrame.TextRange.Text='Same WAV / trim / volume / playback chain'
 $p.Save();@{before=$before;after=@{length=$m.Length;start=$m.StartPoint;end=$m.EndPoint;fadeIn=$m.FadeInDuration;fadeOut=$m.FadeOutDuration}}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $Evidence 'readback.json') -Encoding UTF8
}catch{@{error=$_.Exception.ToString()}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'error.json') -Encoding UTF8;throw}
finally{if($p){try{$p.Close()}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
