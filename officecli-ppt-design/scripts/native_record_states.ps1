param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$Plan,[Parameter(Mandatory=$true)][string]$Directory,[Parameter(Mandatory=$true)][string]$Ffmpeg)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Directory){throw 'Choose a new recording directory'}
$deckPath=[IO.Path]::GetFullPath($Deck);$ffPath=[IO.Path]::GetFullPath($Ffmpeg)
$planData=[IO.File]::ReadAllText([IO.Path]::GetFullPath($Plan))|ConvertFrom-Json
$duration=[double]$planData.duration;$next=@($planData.next)
if($duration -le 0 -or $duration -gt 55){throw 'Recording segment must be 0-55 seconds'}
$last=0.0;foreach($t in $next){if([double]$t -le $last -or [double]$t -ge $duration){throw 'Advance times must increase within the segment'};$last=[double]$t}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($Directory))
Add-Type @'
using System;using System.Text;using System.Runtime.InteropServices;
public static class PptStateRecorder {
 public delegate bool EnumProc(IntPtr h,IntPtr p);
 [StructLayout(LayoutKind.Sequential)] public struct RECT {public int Left,Top,Right,Bottom;}
 [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc c,IntPtr p);
 [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h,int key,out RECT r,int n);
 public static IntPtr Find(string name){IntPtr result=IntPtr.Zero;EnumWindows((h,p)=>{var c=new StringBuilder(256);var t=new StringBuilder(1024);GetClassName(h,c,256);GetWindowText(h,t,1024);if(IsWindowVisible(h)&&c.ToString().ToLower().Contains("screenclass")&&t.ToString().IndexOf(name,StringComparison.OrdinalIgnoreCase)>=0)result=h;return true;},IntPtr.Zero);return result;}
}
'@
$app=New-Object -ComObject PowerPoint.Application;$pres=$null;$show=$null;$rec=$null;$events=@()
try{
 $app.Visible=-1;$pres=$app.Presentations.Open($deckPath,$true,$false,$true)
 $pres.SlideShowSettings.RangeType=1;$pres.SlideShowSettings.StartingSlide=1;$pres.SlideShowSettings.EndingSlide=$pres.Slides.Count
 $pres.SlideShowSettings.AdvanceMode=1;$pres.SlideShowSettings.ShowType=1;$pres.SlideShowSettings.ShowPresenterView=0
 $show=$pres.SlideShowSettings.Run();Start-Sleep -Milliseconds 700
 if(-not $show -and $app.SlideShowWindows.Count -gt 0){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
 $hwnd=[PptStateRecorder]::Find($pres.Name)
 if($hwnd -eq [IntPtr]::Zero){throw 'Own slideshow window not found'}
 [void][PptStateRecorder]::SetForegroundWindow($hwnd)
 if([PptStateRecorder]::GetForegroundWindow() -ne $hwnd){throw 'Own slideshow is not foreground'}
 $rect=New-Object PptStateRecorder+RECT
 if([PptStateRecorder]::DwmGetWindowAttribute($hwnd,9,[ref]$rect,16) -ne 0){throw 'Physical recording bounds unavailable'}
 $width=$rect.Right-$rect.Left;$height=$rect.Bottom-$rect.Top
 $movie=Join-Path $Directory 'native.mp4'
 $argsText='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$rect.Left+' -offset_y '+$rect.Top+' -video_size '+$width+'x'+$height+' -i desktop -t '+$duration.ToString([Globalization.CultureInfo]::InvariantCulture)+' -vf "scale=1280:-2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+[IO.Path]::GetFullPath($movie)+'"'
 $rec=Start-Process -FilePath $ffPath -ArgumentList $argsText -PassThru -WindowStyle Hidden -RedirectStandardError (Join-Path $Directory 'recorder.txt')
 $clock=[Diagnostics.Stopwatch]::StartNew();$index=0
 while($clock.Elapsed.TotalSeconds -lt ($duration+.2)){
  if([PptStateRecorder]::GetForegroundWindow() -ne $hwnd){throw 'Foreground changed; stop recording before further capture'}
  if($index -lt $next.Count -and $clock.Elapsed.TotalSeconds -ge [double]$next[$index]){
   $show.View.Next();$events+=@{type='COM-Next';at=$clock.Elapsed.TotalSeconds;requested=[double]$next[$index]};$index++
  }
  Start-Sleep -Milliseconds 25
 }
 $rec.Handle|Out-Null
 if(-not $rec.WaitForExit(10000)){throw 'Recorder timed out'}
 $rec.Refresh();if($null -ne $rec.ExitCode -and $rec.ExitCode -ne 0){throw 'Recorder failed'}
 if(-not(Test-Path -LiteralPath $movie) -or (Get-Item -LiteralPath $movie).Length -eq 0){throw 'Recording missing or empty'}
 @{recorded=$true;deck=$deckPath;events=$events;duration=$duration;limit='COM Next verifies motion states, not physical click interaction. Movie zero is later than slideshow/autoplay start; inspect actual middle frames.'}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Directory 'record.json') -Encoding utf8
}finally{
 if($rec -and -not $rec.HasExited){$rec.Kill();$rec.WaitForExit()}
 if($show){$show.View.Exit()}
 if($pres){$pres.Close()}
 [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
