param([string]$InputDeck,[string]$OutputDeck,[string]$Evidence)
$ErrorActionPreference='Stop';if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'Fresh outputs required'};[void][IO.Directory]::CreateDirectory($Evidence)
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$log=@()
try{
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$s=$p.Slides.Item(2);$s.Select();$fill=$s.Shapes.Item('reference-image').Fill
 if($fill.PictureEffects.Count -ne 0){throw 'Expected untouched reference picture'};$e=$fill.PictureEffects.Insert(2)
 if($e.EffectParameters.Count -ne 1){throw 'Unexpected native Blur parameters'};$e.EffectParameters.Item(1).Value=[single]30
 if($e.EffectParameters.Item(1).Value -ne 30){throw 'Native Blur radius did not persist'}
 $log+=@{target='reference-picture';effectType=[int]$e.Type;parameter=$e.EffectParameters.Item(1).Name;radiusAfter=$e.EffectParameters.Item(1).Value}
 foreach($j in @(3,4)){$s=$p.Slides.Item($j);$s.Select();foreach($i in @(1,2)){$q=$s.Shapes.Item('glass-'+$i);$q.Fill.Background();if([int]$q.Fill.Type -ne 5){throw 'Not native background fill'};$log+=@{slide=$j;target=$q.Name;fillType=[int]$q.Fill.Type}}}
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24);$log|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'native-readback.json') -Encoding UTF8
}catch{@{error=$_.Exception.ToString();readback=$log}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'error.json') -Encoding UTF8;throw}
finally{if($p){try{$p.Close()}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
