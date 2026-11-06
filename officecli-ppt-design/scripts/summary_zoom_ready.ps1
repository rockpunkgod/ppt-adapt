# Native visible-Normal thumbnail readiness, for this three-color original fixture.
function Wait-SummaryPosters($Slide,[string]$Folder,[string]$Stem,[int]$ExpectedColors=3){
 Add-Type -AssemblyName System.Drawing
 $clock=[Diagnostics.Stopwatch]::StartNew();$states=@();$i=0;$ready=$false
 do{
  $image=Join-Path $Folder ($Stem+'-'+$i+'.png');if(Test-Path -LiteralPath $image){throw 'Refusing overwrite readiness evidence'}
  $Slide.Export($image,'PNG',1600,900);$bitmap=[Drawing.Bitmap]::FromFile($image);$counts=@(0,0,0);$colors=@(@(29,119,116),@(211,109,85),@(90,100,181))
  try{for($y=100;$y -lt 800;$y+=8){for($x=200;$x -lt 1390;$x+=8){$c=$bitmap.GetPixel($x,$y);for($j=0;$j -lt $ExpectedColors;$j++){if([Math]::Abs($c.R-$colors[$j][0]) -le 2 -and [Math]::Abs($c.G-$colors[$j][1]) -le 2 -and [Math]::Abs($c.B-$colors[$j][2]) -le 2){$counts[$j]++}}}}}finally{$bitmap.Dispose()}
  $ready=$true;for($j=0;$j -lt $ExpectedColors;$j++){if($counts[$j] -le 100){$ready=$false}}
  $states+=@{image=$image;time=$clock.Elapsed.TotalSeconds;counts=$counts;ready=$ready};$i++;if(-not $ready){Start-Sleep -Milliseconds 150}
 }while(-not $ready -and $clock.Elapsed.TotalSeconds -lt 4)
 $result=@{passed=$ready;visibleNativeExport=$true;states=$states;finalImage=$image};$result|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $Folder ($Stem+'.json')) -Encoding UTF8
 if(-not $ready){throw 'Visible Summary poster readiness exceeded bound'}
 return $result
}
