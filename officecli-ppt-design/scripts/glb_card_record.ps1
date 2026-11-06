param([string]$Deck,[string]$Evidence,[string]$Ffmpeg)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;using System.Text;using System.Collections.Generic;using System.Runtime.InteropServices;
public static class GlbCardWin{
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
$oldDpi=[GlbCardWin]::SetThreadDpiAwarenessContext([IntPtr](-4))
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
 $h=[GlbCardWin]::Find();[GlbCardWin]::SetForegroundWindow($h)|Out-Null;Start-Sleep -Milliseconds 200
 if([GlbCardWin]::GetForegroundWindow() -ne $h){throw 'Own slideshow not foreground'}
 $r=New-Object GlbCardWin+RECT;[GlbCardWin]::DwmGetWindowAttribute($h,9,[ref]$r,16)|Out-Null
 $width=$r.R-$r.L;$height=$r.B-$r.T
 $movie=Join-Path ([IO.Path]::GetFullPath($Evidence)) 'card-live.mp4'
 $taskRecordArgs='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$r.L+' -offset_y '+$r.T+' -video_size '+$width+'x'+$height+' -i desktop -t 8 -vf "scale=1600:-2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$movie+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $taskRecordArgs -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg-error.txt')
 function Capture([string]$name){$bmp=New-Object Drawing.Bitmap($width,$height);$g=[Drawing.Graphics]::FromImage($bmp);try{$g.CopyFromScreen($r.L,$r.T,0,0,$bmp.Size);$bmp.Save((Join-Path $Evidence ($name+'.png')),[Drawing.Imaging.ImageFormat]::Png)}finally{$g.Dispose();$bmp.Dispose()}}
 Start-Sleep -Milliseconds 700;Capture 'front'
 for($k=1;$k -le 2;$k++){
  if([GlbCardWin]::GetForegroundWindow() -ne $h){throw 'Lost own show foreground'}
  [GlbCardWin]::SetCursorPos($r.L+$width-90,$r.T+$height-90)|Out-Null
  [GlbCardWin]::mouse_event(2,0,0,0,[UIntPtr]::Zero);[GlbCardWin]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
  Start-Sleep -Milliseconds 900;Capture ('middle-'+$k)
  Start-Sleep -Milliseconds 1400;Capture ('end-'+$k)
  $log+=@{click=$k;actualSlide=[int]$show.View.Slide.SlideIndex}
  if([int]$show.View.Slide.SlideIndex -ne ($k+1)){throw 'Physical click did not reach expected card state'}
 }
 if(-not $rec.WaitForExit(15000)){throw 'Screen recording timeout'};if(-not(Test-Path -LiteralPath $movie)){throw 'Movie missing'}
 if((Get-Item -LiteralPath $movie).Length -eq 0){throw 'Movie is empty'}
 $log|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $Evidence 'clicks.json') -Encoding UTF8
 Write-Output 'Native physical-click 3D card recording complete.'
}catch{
 $log|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $Evidence 'partial-clicks.json') -Encoding UTF8
 @{passed=$false;error=$_.Exception.Message}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'failure.json') -Encoding UTF8
 throw
}finally{if($rec -and -not $rec.HasExited){$rec.Kill();$rec.WaitForExit()};if($show){try{$show.View.Exit()}catch{}};if($p){$p.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($p)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null;[GlbCardWin]::SetThreadDpiAwarenessContext($oldDpi)|Out-Null}
