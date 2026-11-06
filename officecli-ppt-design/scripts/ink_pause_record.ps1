param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$Evidence,[Parameter(Mandatory=$true)][string]$Ffmpeg,[int]$Seconds=29)
$ErrorActionPreference='Stop';if(Test-Path -LiteralPath $Evidence){throw 'Fresh evidence required'};[void][IO.Directory]::CreateDirectory($Evidence)
Add-Type @'
using System;using System.Text;using System.Runtime.InteropServices;
public static class InkPauseRecording {
 public delegate bool Call(IntPtr h,IntPtr x);
 [StructLayout(LayoutKind.Sequential)]public struct RECT {public int Left,Top,Right,Bottom;}
 [StructLayout(LayoutKind.Sequential)]public struct POINT {public int X,Y;}
 [DllImport("user32.dll")]public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
 [DllImport("user32.dll")]public static extern bool EnumWindows(Call c,IntPtr x);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]public static extern int GetWindowText(IntPtr h,StringBuilder s,int n);
 [DllImport("user32.dll")]public static extern int GetClassName(IntPtr h,StringBuilder s,int n);
 [DllImport("user32.dll")]public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")]public static extern bool SetWindowPos(IntPtr h,IntPtr after,int x,int y,int w,int z,uint flags);
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]public static extern bool GetClientRect(IntPtr h,out RECT r);
 [DllImport("user32.dll")]public static extern bool ClientToScreen(IntPtr h,ref POINT p);
 public static IntPtr Find(string title){IntPtr result=IntPtr.Zero;EnumWindows((h,x)=>{var c=new StringBuilder(256);var t=new StringBuilder(1024);GetClassName(h,c,256);GetWindowText(h,t,1024);if(IsWindowVisible(h)&&c.ToString().ToLower().Contains("screenclass")&&t.ToString().ToLower().Contains(title.ToLower()))result=h;return true;},IntPtr.Zero);return result;}
}
'@
[void][InkPauseRecording]::SetThreadDpiAwarenessContext([IntPtr](-4))
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$show=$null;$rec=$null
try{
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1);$p.Windows.Item(1).Activate()
 $p.SlideShowSettings.ShowPresenterView=0;$p.SlideShowSettings.ShowType=1
 $show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 500
 if(-not $show){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
 $h=[InkPauseRecording]::Find([IO.Path]::GetFileNameWithoutExtension($Deck));if($h -eq [IntPtr]::Zero){throw 'Own native slideshow not found'}
 [void][InkPauseRecording]::SetForegroundWindow($h)
 if([InkPauseRecording]::GetForegroundWindow() -ne $h){throw 'Own slideshow not foreground; refuse desktop capture'}
 $rect=New-Object InkPauseRecording+RECT;[void][InkPauseRecording]::GetClientRect($h,[ref]$rect)
 $origin=New-Object InkPauseRecording+POINT;[void][InkPauseRecording]::ClientToScreen($h,[ref]$origin)
 $w=$rect.Right-$rect.Left;$height=$rect.Bottom-$rect.Top
 $movie=[IO.Path]::GetFullPath((Join-Path $Evidence 'native.mp4'))
 $args='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$origin.X+' -offset_y '+$origin.Y+' -video_size '+$w+'x'+$height+' -i desktop -t '+$Seconds+' -vf "pad=ceil(iw/2)*2:ceil(ih/2)*2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$movie+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $args -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg.txt')
 if(-not $rec.WaitForExit(($Seconds+8)*1000)){throw 'Recorder timeout'}
 if($rec.ExitCode -ne 0){throw 'Recorder failed'}
 @{nativeFullscreenSlideshow=$true;seconds=$Seconds;client=@{x=$origin.X;y=$origin.Y;width=$w;height=$height};captureBeginsAfterRun=$true;firstCycleIsPartial=$true;slides=$p.Slides.Count;deck=[IO.Path]::GetFullPath($Deck)}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'recording.json') -Encoding UTF8
}catch{@{error=$_.Exception.ToString()}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'error.json') -Encoding UTF8;throw}
finally{
 if($rec -and -not $rec.HasExited){$rec.Kill()}
 if($show){try{$show.View.Exit()}catch{}}
 if($p){try{$p.Close()}catch{@{error=$_.Exception.ToString()}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'close-error.json') -Encoding UTF8}}
 [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
