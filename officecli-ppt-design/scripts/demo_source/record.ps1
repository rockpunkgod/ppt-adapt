param([string]$Deck,[string]$Evidence,[string]$Ffmpeg)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing to overwrite evidence'}
[void][IO.Directory]::CreateDirectory($Evidence)
Add-Type @'
using System;using System.Text;using System.Collections.Generic;using System.Runtime.InteropServices;
public static class PaperShow{
 [StructLayout(LayoutKind.Sequential)]public struct RECT{public int L,T,R,B;}
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();
 [DllImport("dwmapi.dll")]public static extern int DwmGetWindowAttribute(IntPtr h,int a,out RECT r,int s);
 [DllImport("user32.dll")]public static extern void keybd_event(byte k,byte s,uint f,UIntPtr e);
 [DllImport("user32.dll")]public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")]public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr v);
 public delegate bool Call(IntPtr h,IntPtr x);
 [DllImport("user32.dll")]public static extern bool EnumWindows(Call c,IntPtr x);
 [DllImport("user32.dll")]public static extern int GetClassName(IntPtr h,StringBuilder s,int n);
 [DllImport("user32.dll")]public static extern bool IsWindowVisible(IntPtr h);
 public static IntPtr Find(){var found=new List<IntPtr>();EnumWindows((h,x)=>{var c=new StringBuilder(256);GetClassName(h,c,256);if(IsWindowVisible(h)&&c.ToString().ToLower().Contains("screenclass"))found.Add(h);return true;},IntPtr.Zero);if(found.Count!=1)throw new Exception("Expected exactly one slideshow window");return found[0];}
}
'@
$dpi=[PaperShow]::SetThreadDpiAwarenessContext([IntPtr](-4))
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$show=$null;$rec=$null;$log=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1)
 $p.SlideShowSettings.ShowType=1;$p.SlideShowSettings.ShowPresenterView=0;$p.SlideShowSettings.AdvanceMode=1
 $show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 750
 if(-not $show){throw 'No own slideshow returned'}
 if($show.Presentation.FullName -ne $p.FullName){throw 'Slideshow ownership mismatch'}
 $h=[PaperShow]::Find();$showPid=[uint32]0
 [PaperShow]::GetWindowThreadProcessId($h,[ref]$showPid)|Out-Null
 $focus=New-Object -ComObject WScript.Shell;$focus.AppActivate([int]$showPid)|Out-Null
 Start-Sleep -Milliseconds 250
 [PaperShow]::SetForegroundWindow($h)|Out-Null;Start-Sleep -Milliseconds 200
 if([PaperShow]::GetForegroundWindow() -ne $h){throw 'Own show is not foreground'}
 $r=New-Object PaperShow+RECT;[PaperShow]::DwmGetWindowAttribute($h,9,[ref]$r,16)|Out-Null
 $width=$r.R-$r.L;$height=$r.B-$r.T
 [PaperShow]::SetCursorPos($r.R-3,$r.B-3)|Out-Null
 $movie=Join-Path ([IO.Path]::GetFullPath($Evidence)) 'playback.mp4'
 $recordArgs='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$r.L+' -offset_y '+$r.T+' -video_size '+$width+'x'+$height+' -i desktop -t 43 -vf "crop=iw:iw*9/16:0:(ih-iw*9/16)/2,scale=1280:720" -c:v libx264 -preset fast -crf 20 -pix_fmt yuv420p "'+$movie+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $recordArgs -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg-error.txt')
 $recordHandle=$rec.Handle
 Start-Sleep -Milliseconds 1900
 $log+=@{slide=1;input='initial';actualSlide=[int]$show.View.Slide.SlideIndex}
 for($i=2;$i -le 9;$i++){
  if([PaperShow]::GetForegroundWindow() -ne $h){throw 'Lost slideshow foreground'}
  [PaperShow]::keybd_event(0x27,0,0,[UIntPtr]::Zero)
  Start-Sleep -Milliseconds 60
  [PaperShow]::keybd_event(0x27,0,2,[UIntPtr]::Zero)
  Start-Sleep -Milliseconds 4100
  $actual=[int]$show.View.Slide.SlideIndex
  $log+=@{slide=$i;input='physical-right';actualSlide=$actual}
  if($actual -ne $i){throw "Slide navigation mismatch: expected $i actual $actual"}
 }
 if(-not $rec.WaitForExit(15000)){throw 'Recording did not complete'}
 $rec.Refresh()
 if($null -ne $rec.ExitCode -and $rec.ExitCode -ne 0){throw 'Recording failed'}
 & $Ffmpeg -hide_banner -loglevel error -i $movie -f null -
 if($LASTEXITCODE -ne 0){throw 'Recorded video cannot be decoded'}
 @{passed=$true;route=$log;slides=9;physicalRight=8;durationSeconds=43;presentation=$p.FullName}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Evidence 'native-playback.json') -Encoding utf8
}finally{
 if($rec -and -not $rec.HasExited){Stop-Process -Id $rec.Id}
 if($show){try{$show.View.Exit()}catch{}}
 if($p){$p.Close()}
 [PaperShow]::SetThreadDpiAwarenessContext($dpi)|Out-Null
}
