param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$Evidence,[Parameter(Mandatory=$true)][string]$Ffmpeg)
$ErrorActionPreference='Stop';if(Test-Path -LiteralPath $Evidence){throw 'Evidence exists'};[void][IO.Directory]::CreateDirectory($Evidence)
Add-Type @'
using System;using System.Text;using System.Runtime.InteropServices;
public static class BookmarkNative {
 public delegate bool Call(IntPtr h,IntPtr x);
 [StructLayout(LayoutKind.Sequential)] public struct RECT {public int Left,Top,Right,Bottom;}
 [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr c);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string c,string t);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] public static extern bool EnumWindows(Call c,IntPtr x);
 [DllImport("user32.dll")] public static extern int GetClassName(IntPtr h,StringBuilder s,int n);
 public static IntPtr Find(){IntPtr result=IntPtr.Zero;EnumWindows((h,x)=>{var c=new StringBuilder(256);GetClassName(h,c,256);if(c.ToString().ToLower().Contains("screenclass"))result=h;return true;},IntPtr.Zero);return result;}
 [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h,int a,out RECT r,int size);
}
'@
[BookmarkNative]::SetProcessDpiAwarenessContext([IntPtr](-4))|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$show=$null;$rec=$null;$log=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1);$app.Visible=-1
 $p.SlideShowSettings.ShowPresenterView=0;$p.SlideShowSettings.ShowType=1
 $show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 1000
 if(-not $show){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
 $h=[BookmarkNative]::Find();if($h -eq [IntPtr]::Zero){throw 'Missing slideshow window'}
 [BookmarkNative]::SetForegroundWindow($h)|Out-Null;$r=New-Object BookmarkNative+RECT
 [void][BookmarkNative]::DwmGetWindowAttribute($h,9,[ref]$r,16)
 $movie=Join-Path $Evidence 'native.mp4'
 $args='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$r.Left+' -offset_y '+$r.Top+' -video_size '+($r.Right-$r.Left)+'x'+($r.Bottom-$r.Top)+' -i desktop -t 10 -vf "scale=1280:-2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$movie+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $args -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg.txt')
 $audio=$p.Slides.Item(1).Shapes.Item('sequence-clock');$timer=[Diagnostics.Stopwatch]::StartNew()
 while($timer.Elapsed.TotalSeconds -lt 10){
  try{$player=$show.View.Player($audio.Id);$log+=@{elapsed=$timer.Elapsed.TotalSeconds;position=$player.CurrentPosition;state=$player.State}}catch{$log+=@{elapsed=$timer.Elapsed.TotalSeconds;error=$_.Exception.Message}}
  Start-Sleep -Milliseconds 100
 }
 [void]$rec.WaitForExit(5000)
 $log|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $Evidence 'player-samples.json') -Encoding UTF8
 if(-not(Test-Path -LiteralPath $movie)){throw 'Missing movie'}
}finally{if($rec -and -not $rec.HasExited){$rec.Kill()};if($show){$show.View.Exit()};if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
