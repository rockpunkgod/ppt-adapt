param([string]$Deck,[string]$Evidence,[string]$Ffmpeg,[int]$Seconds=8,[string]$Helpers=$PSScriptRoot)
$ErrorActionPreference='Stop';if(Test-Path -LiteralPath $Evidence){throw 'Fresh evidence required'};[void][IO.Directory]::CreateDirectory($Evidence)
if(-not ('OfficeCliMediaControlsWindow' -as [type])){. (Join-Path $Helpers 'media_controls_window.ps1')};if(-not ('OfficeCliMediaControls.ProcessLoopback' -as [type])){Add-Type -Path (Join-Path $Helpers 'media_controls_capture.cs')}
Add-Type -Path 'C:/Windows/assembly/GAC_MSIL/office/15.0.0.0__71e9bce111e9429c/OFFICE.DLL';$pia=Get-ChildItem -LiteralPath 'C:/Windows/assembly/GAC_MSIL/Microsoft.Office.Interop.PowerPoint' -Recurse -Filter '*.dll'|Select-Object -First 1;Add-Type -Path $pia.FullName
$oldDpi=[OfficeCliMediaControlsWindow]::SetThreadDpiAwarenessContext([IntPtr](-4));$app=New-Object -ComObject PowerPoint.Application;$p=$null;$show=$null;$movieRec=$null;$audioRec=$null;$events=@();$clock=[Diagnostics.Stopwatch]::StartNew()
function Wait-At($second){while($clock.Elapsed.TotalSeconds-$videoLaunch -lt $second){Start-Sleep -Milliseconds 10}}
try{
 if($app.SlideShowWindows.Count -ne 0){throw 'Another PowerPoint slideshow is running; coordinate the shared process before capture'}
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1);$p.Windows.Item(1).Activate()
 $media=$p.Slides.Item(1).Shapes.Item('lifecycle-audio')
 $hw=[Microsoft.Office.Interop.PowerPoint.DocumentWindow].GetProperty('HWND').GetValue($p.Windows.Item(1),$null);$target=[uint32]0;[void][OfficeCliMediaControlsWindow]::GetWindowThreadProcessId([IntPtr]$hw,[ref]$target)
 if((Get-Process -Id $target).ProcessName -ne 'POWERPNT'){throw 'Own PowerPoint PID guard failed'}
 $audioRec=New-Object OfficeCliMediaControls.ProcessLoopback ([int]$target),([IO.Path]::GetFullPath((Join-Path $Evidence 'powerpoint-audio.wav')));$audioRec.Start();$audioReady=$clock.Elapsed.TotalSeconds
 $settings=$p.SlideShowSettings;$settings.ShowPresenterView=0;$settings.ShowType=1;$settings.RangeType=2;$settings.StartingSlide=1;$settings.EndingSlide=3
 $show=$settings.Run();if(-not $show){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
 if($show.View.Slide.SlideIndex -ne 1){throw 'Native show did not start requested media experiment slide'}
 $until=(Get-Date).AddSeconds(4);$h=[IntPtr]::Zero;while($h -eq [IntPtr]::Zero -and (Get-Date) -lt $until){$h=[OfficeCliMediaControlsWindow]::Find([IO.Path]::GetFileNameWithoutExtension($Deck));if($h -eq [IntPtr]::Zero){Start-Sleep -Milliseconds 50}}
 if($h -eq [IntPtr]::Zero){throw 'Own slideshow missing'};[void][OfficeCliMediaControlsWindow]::SetForegroundWindow($h);if([OfficeCliMediaControlsWindow]::GetForegroundWindow() -ne $h){throw 'Own slideshow not foreground'}
 $rect=New-Object OfficeCliMediaControlsWindow+RECT;[void][OfficeCliMediaControlsWindow]::GetClientRect($h,[ref]$rect);$origin=New-Object OfficeCliMediaControlsWindow+POINT;[void][OfficeCliMediaControlsWindow]::ClientToScreen($h,[ref]$origin);$w=$rect.Right-$rect.Left;$height=$rect.Bottom-$rect.Top
 $movie=[IO.Path]::GetFullPath((Join-Path $Evidence 'native.mp4'));$args='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$origin.X+' -offset_y '+$origin.Y+' -video_size '+$w+'x'+$height+' -i desktop -t '+$Seconds+' -vf "pad=ceil(iw/2)*2:ceil(ih/2)*2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$movie+'"'
 $videoLaunch=$clock.Elapsed.TotalSeconds;$movieRec=Start-Process -FilePath $Ffmpeg -ArgumentList $args -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg.txt')
 $events+=@{action='initial';elapsed=$clock.Elapsed.TotalSeconds-$videoLaunch;absoluteSeconds=$clock.Elapsed.TotalSeconds;slide=$show.View.Slide.SlideIndex}
 foreach($targetTime in @(2.0,4.5)){
  Wait-At $targetTime;$before=$show.View.Slide.SlideIndex;$t=$clock.Elapsed.TotalSeconds
  $show.View.Next();$events+=@{action='native-next';elapsed=$clock.Elapsed.TotalSeconds-$videoLaunch;absoluteSeconds=$clock.Elapsed.TotalSeconds;callStartSeconds=$t;before=$before;slide=$show.View.Slide.SlideIndex}
 }
 if(-not $movieRec.WaitForExit(($Seconds+8)*1000)){throw 'Own desktop recorder timeout'};if($movieRec.ExitCode -ne 0){throw 'Desktop recording failed'}
 $audioRec.Stop()
 @{experiment="cross-slide-stop";mediaId=$media.Id;targetPid=$target;recordScriptSHA256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash;captureHelperSHA256=(Get-FileHash -LiteralPath (Join-Path $Helpers 'media_controls_capture.cs') -Algorithm SHA256).Hash;audioReadySeconds=$audioReady;videoLaunchSeconds=$videoLaunch;audioBytes=$audioRec.AudioBytes;audioPackets=$audioRec.Packets;audioDiscontinuities=$audioRec.Discontinuities;audioTimestampErrors=$audioRec.TimestampErrors;finalSlide=$show.View.Slide.SlideIndex;events=$events;client=@{x=$origin.X;y=$origin.Y;width=$w;height=$height};note='WAV begins before slideshow Run; MP4 begins after Run and encoder startup. Separate capture clocks are logged; no frame-exact A/V synchronization claimed.'}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'recording.json') -Encoding UTF8
}catch{@{error=$_.Exception.ToString()}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'error.json') -Encoding UTF8;throw}
finally{if($audioRec){try{$audioRec.Dispose()}catch{}};if($movieRec -and -not $movieRec.HasExited){$movieRec.Kill()};if($show){try{$show.View.Exit()}catch{}};if($p){try{$p.Close()}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app);[void][OfficeCliMediaControlsWindow]::SetThreadDpiAwarenessContext($oldDpi)}
