param([string]$Deck,[string]$Evidence,[string]$Ffmpeg,[switch]$ProbeOnly)
$ErrorActionPreference='Stop';if(Test-Path -LiteralPath $Evidence){throw 'Fresh evidence required'};[void][IO.Directory]::CreateDirectory($Evidence)
. (Join-Path $PSScriptRoot 'gif_menu_window.ps1');$dpi=[GifMenuNative]::SetThreadDpiAwarenessContext([IntPtr](-4));$app=New-Object -ComObject PowerPoint.Application;$p=$null;$show=$null;$rec=$null;$events=@();$clock=[Diagnostics.Stopwatch]::StartNew()
try{
 if($app.SlideShowWindows.Count -ne 0){throw 'Another slideshow is active'};$app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1);$p.Windows.Item(1).Activate();$p.SlideShowSettings.ShowPresenterView=0;$p.SlideShowSettings.ShowType=1;$show=$p.SlideShowSettings.Run()
 if(-not $show){throw 'No owned slideshow'};$hwnd=[GifMenuNative]::Find($p.Name);if($hwnd -eq [IntPtr]::Zero){throw 'Owned visible slideshow missing'};[void][GifMenuNative]::SetForegroundWindow($hwnd)
 $r=New-Object GifMenuNative+R;[void][GifMenuNative]::GetClientRect($hwnd,[ref]$r);$w=$r.Rt;$h=$r.B;$scale=[Math]::Min($w/33.87,$h/19.05);$ox=($w-33.87*$scale)/2;$oy=($h-19.05*$scale)/2;$origin=New-Object GifMenuNative+P;[void][GifMenuNative]::ClientToScreen($hwnd,[ref]$origin)
 function Click([double]$x,[double]$y,[string]$name,[int]$want){
  if([GifMenuNative]::GetForegroundWindow() -ne $hwnd){throw 'Owned show lost foreground'};$pt=New-Object GifMenuNative+P;$pt.X=[int]($origin.X+$ox+$x*$scale);$pt.Y=[int]($origin.Y+$oy+$y*$scale)
  if([GifMenuNative]::GetAncestor([GifMenuNative]::WindowFromPoint($pt),2) -ne $hwnd){throw 'Physical point outside owned slideshow'}
  [void][GifMenuNative]::SetCursorPos($pt.X,$pt.Y);[GifMenuNative]::mouse_event(1,1,0,0,[UIntPtr]::Zero);[GifMenuNative]::mouse_event(1,-1,0,0,[UIntPtr]::Zero);Start-Sleep -Milliseconds 50;$actual=New-Object GifMenuNative+P;[void][GifMenuNative]::GetCursorPos([ref]$actual)
  $physical=New-Object GifMenuNative+R;[void][GifMenuNative]::DwmGetWindowAttribute($hwnd,9,[ref]$physical,16)
  if($ProbeOnly -or $actual.X -ne $pt.X -or $actual.Y -ne $pt.Y){@{expected=@($pt.X,$pt.Y);actual=@($actual.X,$actual.Y);client=@{width=$w;height=$h;originX=$origin.X;originY=$origin.Y};dwm=@($physical.L,$physical.T,$physical.Rt,$physical.B);showCOM=@{left=$show.Left;top=$show.Top;width=$show.Width;height=$show.Height;hwnd=$show.HWND};scale=$scale;offset=@($ox,$oy);probeOnly=[bool]$ProbeOnly}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'pointer-mismatch.json') -Encoding UTF8;throw 'Pointer diagnostic stop before clicking'};$before=$show.View.Slide.SlideIndex;$click=$clock.Elapsed.TotalSeconds-$launch
  [GifMenuNative]::mouse_event(2,0,0,0,[UIntPtr]::Zero);[GifMenuNative]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
  $limit=(Get-Date).AddMilliseconds(900);while($show.View.Slide.SlideIndex -ne $want -and (Get-Date) -lt $limit){Start-Sleep -Milliseconds 30}
  $slide=$show.View.Slide.SlideIndex;$script:events+=@{name=$name;kind='physical-click';before=$before;actualSlide=$slide;expectedSlide=$want;clickElapsed=$click;returnElapsed=$clock.Elapsed.TotalSeconds-$launch;point=@($pt.X,$pt.Y);actualPoint=@($actual.X,$actual.Y)}
  if($slide -ne $want){throw 'Native hyperlink target failed'}
 }
 $movie=[IO.Path]::GetFullPath((Join-Path $Evidence 'native.mp4'));$args='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$origin.X+' -offset_y '+$origin.Y+' -video_size '+$w+'x'+$h+' -i desktop -t 20 -vf "scale=1280:-2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$movie+'"'
 $launch=$clock.Elapsed.TotalSeconds;$rec=Start-Process -FilePath $Ffmpeg -ArgumentList $args -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg.txt')
 foreach($step in @(@(1,16.9,9.5,'first-home',2),@(4,16.935,7.8,'first-world',3),@(7,26.97,7.8,'first-create',4),@(10,6.9,7.8,'reenter-home',2),@(13,16.935,7.8,'reenter-world',3),@(16,26.97,7.8,'reenter-create',4))){
  while($clock.Elapsed.TotalSeconds-$launch -lt [double]$step[0]){Start-Sleep -Milliseconds 10};Click ([double]$step[1]) ([double]$step[2]) ([string]$step[3]) ([int]$step[4])
 }
 if(-not $rec.WaitForExit(12000)){throw 'Recorder timeout'};if($rec.ExitCode -ne 0){throw 'Recording failed'}
 @{passed=$true;events=$events;width=$w;height=$h;videoLaunch=$launch;note='Actual physical clicks; GIF phase must be analyzed from visible cyan marker separately. Event clocks are not frame-exact synchronized with encoder startup. Current-slide self-link behavior is untested.'}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'events.json') -Encoding UTF8
}catch{@{passed=$false;error=$_.Exception.ToString();events=$events}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'error.json') -Encoding UTF8;throw}
finally{if($rec -and -not $rec.HasExited){$rec.Kill()};if($show){try{$show.View.Exit()}catch{}};if($p){try{$p.Close()}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app);[void][GifMenuNative]::SetThreadDpiAwarenessContext($dpi)}
