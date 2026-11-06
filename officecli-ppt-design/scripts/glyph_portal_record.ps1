param([string]$Deck,[string]$Evidence,[string]$Ffmpeg,[string]$Helpers=$PSScriptRoot)
$ErrorActionPreference='Stop';if(Test-Path -LiteralPath $Evidence){throw 'Fresh evidence required'};[void][IO.Directory]::CreateDirectory($Evidence)
. (Join-Path $Helpers 'media_controls_window.ps1');$oldDpi=[OfficeCliMediaControlsWindow]::SetThreadDpiAwarenessContext([IntPtr](-4));$app=New-Object -ComObject PowerPoint.Application;$p=$null;$show=$null;$rec=$null;$clock=[Diagnostics.Stopwatch]::StartNew()
try{
 if($app.SlideShowWindows.Count -ne 0){throw 'Another slideshow already running'};$app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1);$p.Windows.Item(1).Activate()
 $settings=$p.SlideShowSettings;$settings.ShowPresenterView=0;$settings.ShowType=1;$settings.RangeType=2;$settings.StartingSlide=1;$settings.EndingSlide=2;$show=$settings.Run()
 if(-not $show){throw 'No owned slideshow'}
 $until=(Get-Date).AddSeconds(3);$hw=[IntPtr]::Zero;while($hw -eq [IntPtr]::Zero -and (Get-Date) -lt $until){$hw=[OfficeCliMediaControlsWindow]::Find([IO.Path]::GetFileNameWithoutExtension($Deck));if($hw -eq [IntPtr]::Zero){Start-Sleep -Milliseconds 50}}
 if($hw -eq [IntPtr]::Zero){throw 'Owned screenClass not found'};[void][OfficeCliMediaControlsWindow]::SetForegroundWindow($hw);if([OfficeCliMediaControlsWindow]::GetForegroundWindow() -ne $hw){throw 'Owned slideshow not foreground'}
 $r=New-Object OfficeCliMediaControlsWindow+RECT;[void][OfficeCliMediaControlsWindow]::GetClientRect($hw,[ref]$r);$o=New-Object OfficeCliMediaControlsWindow+POINT;[void][OfficeCliMediaControlsWindow]::ClientToScreen($hw,[ref]$o)
 $movie=[IO.Path]::GetFullPath((Join-Path $Evidence 'native.mp4'));$args='-hide_banner -loglevel error -f gdigrab -framerate 60 -offset_x '+$o.X+' -offset_y '+$o.Y+' -video_size '+$r.Right+'x'+$r.Bottom+' -i desktop -t 6 -vf "pad=ceil(iw/2)*2:ceil(ih/2)*2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$movie+'"'
 $launch=$clock.Elapsed.TotalSeconds;$rec=Start-Process -FilePath $Ffmpeg -ArgumentList $args -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg.txt')
 while($clock.Elapsed.TotalSeconds-$launch -lt 1.0){Start-Sleep -Milliseconds 10};$before=$clock.Elapsed.TotalSeconds;$show.View.Next();$after=$clock.Elapsed.TotalSeconds
 if($show.View.Slide.SlideIndex -ne 2){throw 'Native Next did not enter page 2'}
 if(-not $rec.WaitForExit(12000)){throw 'Recorder timeout'};if($rec.ExitCode -ne 0){throw 'Recording failed'}
 @{videoLaunch=$launch;nextStart=$before;nextReturned=$after;slide=$show.View.Slide.SlideIndex;captureFPS=60;note='Actual desktop capture; native COM Next lifecycle, no physical click claim. Capture encoder startup is not frame-exact synchronized to Next.'}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'recording.json') -Encoding UTF8
}finally{if($rec -and -not $rec.HasExited){$rec.Kill()};if($show){try{$show.View.Exit()}catch{}};if($p){try{$p.Close()}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app);[void][OfficeCliMediaControlsWindow]::SetThreadDpiAwarenessContext($oldDpi)}
