param([string]$Deck,[string]$Evidence,[string]$Ffmpeg)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;using System.Text;using System.Collections.Generic;using System.Runtime.InteropServices;
public static class ChartBuildWin{
 [StructLayout(LayoutKind.Sequential)]public struct RECT{public int L,T,R,B;}
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")]public static extern void mouse_event(uint f,uint x,uint y,uint d,UIntPtr e);
 [DllImport("dwmapi.dll")]public static extern int DwmGetWindowAttribute(IntPtr h,int a,out RECT r,int s);
 [DllImport("user32.dll")]public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr v);
 public delegate bool Call(IntPtr h,IntPtr x);
 [DllImport("user32.dll")]public static extern bool EnumWindows(Call c,IntPtr x);
 [DllImport("user32.dll")]public static extern int GetClassName(IntPtr h,StringBuilder s,int n);
 [DllImport("user32.dll")]public static extern bool IsWindowVisible(IntPtr h);
 public static IntPtr Find(){var found=new List<IntPtr>();EnumWindows((h,x)=>{var c=new StringBuilder(256);GetClassName(h,c,256);if(IsWindowVisible(h)&&c.ToString().ToLower().Contains("screenclass"))found.Add(h);return true;},IntPtr.Zero);if(found.Count!=1)throw new Exception("Expected exactly one visible projection window; found "+found.Count);return found[0];}
}
'@
$oldDpi=[ChartBuildWin]::SetThreadDpiAwarenessContext([IntPtr](-4))
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$show=$null;$rec=$null;$log=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1)
 $p.SlideShowSettings.ShowType=1;$p.SlideShowSettings.ShowPresenterView=0;$p.SlideShowSettings.AdvanceMode=1
 $show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 650
 if(-not $show -and $app.SlideShowWindows.Count -gt 0){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
 if(-not $show){throw 'Run did not return an own slideshow window'}
 if($show.Presentation.FullName -ne $p.FullName){throw 'Slideshow does not belong to this task presentation'}
 $show.View.PointerType=1
 if([int]$show.View.PointerType -ne 1){throw 'Arrow pointer did not persist'}
 $h=[ChartBuildWin]::Find();[ChartBuildWin]::SetForegroundWindow($h)|Out-Null;Start-Sleep -Milliseconds 200
 if([ChartBuildWin]::GetForegroundWindow() -ne $h){throw 'Own slideshow not foreground'}
 $r=New-Object ChartBuildWin+RECT;[ChartBuildWin]::DwmGetWindowAttribute($h,9,[ref]$r,16)|Out-Null
 $width=$r.R-$r.L;$height=$r.B-$r.T
 $movie=Join-Path ([IO.Path]::GetFullPath($Evidence)) 'chart-build-live.mp4'
 $taskRecordArgs='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$r.L+' -offset_y '+$r.T+' -video_size '+$width+'x'+$height+' -i desktop -t 12 -vf "scale=1600:-2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$movie+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $taskRecordArgs -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg-error.txt')
 function Capture([string]$name){$bmp=New-Object Drawing.Bitmap($width,$height);$g=[Drawing.Graphics]::FromImage($bmp);try{$g.CopyFromScreen($r.L,$r.T,0,0,$bmp.Size);$bmp.Save((Join-Path $Evidence ($name+'.png')),[Drawing.Imaging.ImageFormat]::Png)}finally{$g.Dispose();$bmp.Dispose()}}
 for($slide=1;$slide -le 2;$slide++){
  $show.View.GotoSlide($slide,-1);Start-Sleep -Milliseconds 550
  Capture ('s'+$slide+'-c0')
  $log+=@{slide=$slide;click=0;actualSlide=[int]$show.View.Slide.SlideIndex;index=[int]$show.View.GetClickIndex()}
  $n=$(if($slide -eq 1){2}else{4})
  for($k=1;$k -le $n;$k++){
   if([ChartBuildWin]::GetForegroundWindow() -ne $h){throw 'Lost own show foreground'}
   [ChartBuildWin]::SetCursorPos($r.L+$width-90,$r.T+$height-90)|Out-Null
   [ChartBuildWin]::mouse_event(2,0,0,0,[UIntPtr]::Zero);[ChartBuildWin]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
   Start-Sleep -Milliseconds 750;Capture ('s'+$slide+'-c'+$k)
   $log+=@{slide=$slide;click=$k;actualSlide=[int]$show.View.Slide.SlideIndex;index=[int]$show.View.GetClickIndex()}
   if([int]$show.View.Slide.SlideIndex -ne $slide){throw 'Click advanced slide before chart data groups completed'}
  }
 }
 $rec.WaitForExit();if(-not(Test-Path -LiteralPath $movie)){throw 'Movie missing'}
 if((Get-Item -LiteralPath $movie).Length -eq 0){throw 'Movie is empty'}
 $log|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $Evidence 'clicks.json') -Encoding UTF8
 Write-Output 'Native physical-click chart build recording complete.'
}catch{
 $log|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $Evidence 'partial-clicks.json') -Encoding UTF8
 @{passed=$false;error=$_.Exception.Message}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'failure.json') -Encoding UTF8
 throw
}finally{if($rec -and -not $rec.HasExited){$rec.Kill();$rec.WaitForExit()};if($show){try{$show.View.Exit()}catch{}};if($p){$p.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($p)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null;[ChartBuildWin]::SetThreadDpiAwarenessContext($oldDpi)|Out-Null}
