param([string]$Deck,[string]$Evidence,[string]$Ffmpeg)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;using System.Text;using System.Collections.Generic;using System.Runtime.InteropServices;
public static class MasterLoopWin{
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
$oldDpi=[MasterLoopWin]::SetThreadDpiAwarenessContext([IntPtr](-4))
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
 $h=[MasterLoopWin]::Find();[MasterLoopWin]::SetForegroundWindow($h)|Out-Null;Start-Sleep -Milliseconds 200
 if([MasterLoopWin]::GetForegroundWindow() -ne $h){throw 'Own slideshow not foreground'}
 $r=New-Object MasterLoopWin+RECT;[MasterLoopWin]::DwmGetWindowAttribute($h,9,[ref]$r,16)|Out-Null
 $width=$r.R-$r.L;$height=$r.B-$r.T
 $movie=Join-Path ([IO.Path]::GetFullPath($Evidence)) 'master-loop-live.mp4'
 $taskRecordArgs='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$r.L+' -offset_y '+$r.T+' -video_size '+$width+'x'+$height+' -i desktop -t 18 -vf "scale=1600:-2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$movie+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $taskRecordArgs -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg-error.txt')
 function Capture([string]$name){$bmp=New-Object Drawing.Bitmap($width,$height);$g=[Drawing.Graphics]::FromImage($bmp);try{$g.CopyFromScreen($r.L,$r.T,0,0,$bmp.Size);$bmp.Save((Join-Path $Evidence ($name+'.png')),[Drawing.Imaging.ImageFormat]::Png)}finally{$g.Dispose();$bmp.Dispose()}}
 $clock=[Diagnostics.Stopwatch]::StartNew()
 function Until([double]$at){while($clock.Elapsed.TotalSeconds -lt $at){Start-Sleep -Milliseconds 15}}
 function Event([string]$kind){$script:log+=@{kind=$kind;time=$clock.Elapsed.TotalSeconds;slide=[int]$show.View.Slide.SlideIndex;clickIndex=[int]$show.View.GetClickIndex()}}
 function Click(){
  if([MasterLoopWin]::GetForegroundWindow() -ne $h){throw 'Own show lost foreground'}
  [MasterLoopWin]::SetCursorPos($r.L+$width-90,$r.T+$height-90)|Out-Null
  [MasterLoopWin]::mouse_event(2,0,0,0,[UIntPtr]::Zero);[MasterLoopWin]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
  Start-Sleep -Milliseconds 60
 }
 Event 'initial';Until 0.35;Capture 'page1-before'
 Until 1.2;Click;Event 'local-click';Until 1.8;Capture 'page1-after'
 Until 7.25;Click;Event 'next';if($show.View.Slide.SlideIndex -ne 2){throw 'Expected natural advance to A2'}
 $entry2=$clock.Elapsed.TotalSeconds;Until ($entry2+0.2);Capture 'page2-before'
 Until ($entry2+1.2);Click;Event 'local-click';Until ($entry2+1.8);Capture 'page2-after'
 Until ($entry2+3.3);Click;Event 'next';if($show.View.Slide.SlideIndex -ne 3){throw 'Expected natural advance to B3'}
 $entry3=$clock.Elapsed.TotalSeconds;Until ($entry3+0.2);Capture 'page3-before'
 Until ($entry3+1.2);Click;Event 'local-click';Until ($entry3+1.8);Capture 'page3-after'
 Until ($entry3+3.3);$show.View.GotoSlide(1,-1);Event 'explicit-reset-return'
 $entry4=$clock.Elapsed.TotalSeconds;Until ($entry4+0.2);Capture 'return1-before'
 Until ($entry4+2.2);Capture 'return1-loop'
 if(-not $rec.WaitForExit(25000)){throw 'Recording timeout'};if(-not(Test-Path -LiteralPath $movie)){throw 'Movie missing'}
 if((Get-Item -LiteralPath $movie).Length -eq 0){throw 'Movie is empty'}
 $log|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $Evidence 'clicks.json') -Encoding UTF8
 Write-Output 'Native inherited-loop natural advances and explicit return/reset recording complete.'
}catch{
 $log|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $Evidence 'partial-clicks.json') -Encoding UTF8
 @{passed=$false;error=$_.Exception.Message}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'failure.json') -Encoding UTF8
 throw
}finally{if($rec -and -not $rec.HasExited){$rec.Kill();$rec.WaitForExit()};if($show){try{$show.View.Exit()}catch{}};if($p){$p.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($p)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null;[MasterLoopWin]::SetThreadDpiAwarenessContext($oldDpi)|Out-Null}
