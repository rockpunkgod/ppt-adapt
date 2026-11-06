param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$Report,[Parameter(Mandatory=$true)][string]$OutputDirectory,[Parameter(Mandatory=$true)][string]$Ffmpeg)
if((Test-Path -LiteralPath $Report) -or (Test-Path -LiteralPath $OutputDirectory)){throw 'Refusing existing output'}
[IO.Directory]::CreateDirectory($OutputDirectory)|Out-Null
$ErrorActionPreference='Stop'
Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class PptZoomWin32 {
  public delegate bool Call(IntPtr h,IntPtr x);
  [StructLayout(LayoutKind.Sequential)] public struct RECT {public int Left,Top,Right,Bottom;}
  [StructLayout(LayoutKind.Sequential)] public struct POINT {public int X,Y;}
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr hwnd,out RECT rect);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hwnd,uint msg,IntPtr wParam,IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(Call c,IntPtr x);
  [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h,int attribute,out RECT r,int size);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h,StringBuilder s,int n);
  [DllImport("user32.dll")] public static extern int GetClassName(IntPtr h,StringBuilder s,int n);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h,ref POINT p);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint flags,uint dx,uint dy,uint data,UIntPtr extra);
  public static IntPtr Find(){IntPtr found=IntPtr.Zero;EnumWindows((h,x)=>{var c=new StringBuilder(256);GetClassName(h,c,256);if(c.ToString().ToLower().Contains("screenclass"))found=h;return true;},IntPtr.Zero);return found;}
}
'@
$taskApp=New-Object -ComObject PowerPoint.Application
$taskP=$null;$taskShow=$null;$taskLog=@()
try {
 $taskP=$taskApp.Presentations.Open([IO.Path]::GetFullPath($Deck),$true,$false,$true)
 $taskApp.Visible=-1
 for($idx=1;$idx -le $taskP.Slides.Count;$idx++){$taskP.Slides.Item($idx).Export((Join-Path $OutputDirectory "slide-$idx.png"),'PNG',1600,900)}
 $taskP.SlideShowSettings.ShowType=1
 $taskP.SlideShowSettings.ShowPresenterView=0
 $taskShow=$taskP.SlideShowSettings.Run()
 Start-Sleep -Milliseconds 700
 if(-not $taskShow -and $taskApp.SlideShowWindows.Count -gt 0){$taskShow=$taskApp.SlideShowWindows.Item($taskApp.SlideShowWindows.Count)}
 $taskHwnd=[PptZoomWin32]::Find()
 if($taskShow.HWND){$taskHwnd=[IntPtr]$taskShow.HWND}
 if($taskHwnd -eq [IntPtr]::Zero){throw 'No screenClass slideshow window found'}
 [PptZoomWin32]::SetForegroundWindow($taskHwnd)|Out-Null
 $taskRect=New-Object PptZoomWin32+RECT
 [PptZoomWin32]::GetClientRect($taskHwnd,[ref]$taskRect)|Out-Null
 $taskWidth=$taskRect.Right-$taskRect.Left;$taskHeight=$taskRect.Bottom-$taskRect.Top
 $taskScale=[Math]::Min($taskWidth/33.87,$taskHeight/19.05)
 $taskOffsetX=($taskWidth-33.87*$taskScale)/2;$taskOffsetY=($taskHeight-19.05*$taskScale)/2
 $taskCaption=New-Object Text.StringBuilder 512
 [PptZoomWin32]::GetWindowText($taskHwnd,$taskCaption,512)|Out-Null
 $taskMovie=Join-Path $OutputDirectory 'section-native.mp4'
 $taskFF=[IO.Path]::GetFullPath($Ffmpeg)
 $taskPhysicalRect=New-Object PptZoomWin32+RECT
 [PptZoomWin32]::DwmGetWindowAttribute($taskHwnd,9,[ref]$taskPhysicalRect,16)|Out-Null
 $taskVW=$taskPhysicalRect.Right-$taskPhysicalRect.Left;$taskVH=$taskPhysicalRect.Bottom-$taskPhysicalRect.Top
 $taskArgs='-hide_banner -loglevel error -f gdigrab -framerate 20 -offset_x '+$taskPhysicalRect.Left+' -offset_y '+$taskPhysicalRect.Top+' -video_size '+$taskVW+'x'+$taskVH+' -i desktop -t 13 -vf "scale=1280:-2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$taskMovie+'"'
 $taskRecorder=Start-Process -FilePath $taskFF -ArgumentList $taskArgs -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $OutputDirectory 'recording-error.txt')
 Start-Sleep -Milliseconds 400
 function Click-PptPoint([double]$X,[double]$Y,[int]$Expected,[string]$Name){
   $taskPx=[int]($taskOffsetX+$X*$taskScale);$taskPy=[int]($taskOffsetY+$Y*$taskScale)
   $taskPoint=[IntPtr]($taskPx -bor ($taskPy -shl 16))
   $taskScreenPoint=New-Object PptZoomWin32+POINT
   $taskScreenPoint.X=$taskPx;$taskScreenPoint.Y=$taskPy
   [PptZoomWin32]::ClientToScreen($taskHwnd,[ref]$taskScreenPoint)|Out-Null
   [PptZoomWin32]::SetCursorPos($taskScreenPoint.X,$taskScreenPoint.Y)|Out-Null
   Start-Sleep -Milliseconds 100
   [PptZoomWin32]::mouse_event(2,0,0,0,[UIntPtr]::Zero)
   [PptZoomWin32]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
   Start-Sleep -Milliseconds 3200
   $taskActual=[int]$taskShow.View.Slide.SlideIndex
   $script:taskLog+=@{name=$Name;expected=$Expected;actual=$taskActual;passed=($taskActual -eq $Expected);clientPoint=@($taskPx,$taskPy)}
   if($taskActual -ne $Expected){throw "Click $Name expected slide $Expected, actual $taskActual"}
 }
 # Native Section Zoom must traverse auto entry + content, then return at the section end.
 foreach($card in @(@{x=9.2;target=3;name='section-A'},@{x=24.2;target=5;name='section-B'})){
  $taskShow.View.GotoSlide(1)
  Start-Sleep -Milliseconds 250
  Click-PptPoint $card.x 9.4 $card.target $card.name
  $taskShow.View.Next()
  Start-Sleep -Milliseconds 1400
  $actual=[int]$taskShow.View.Slide.SlideIndex
  $script:taskLog+=@{name=($card.name+'-range-end-return');expected=1;actual=$actual;passed=($actual -eq 1)}
  if($actual -ne 1){throw "Section end did not return: $actual"}
 }
 $taskRecorder.Handle|Out-Null
 $taskRecorder.WaitForExit()
 if($null -ne $taskRecorder.ExitCode -and $taskRecorder.ExitCode -ne 0){throw "Recorder failed"}
 if(-not(Test-Path -LiteralPath $taskMovie)){throw 'Recording output missing'}
 if((Get-Item -LiteralPath $taskMovie).Length -eq 0){throw 'Recording output empty'}
 $taskResult=@{passed=$true;checks=$taskLog;client=@{width=$taskWidth;height=$taskHeight};file=[IO.Path]::GetFullPath($Deck)}
} catch {
 $taskResult=@{passed=$false;checks=$taskLog;error=$_.Exception.Message;file=[IO.Path]::GetFullPath($Deck)}
 throw
} finally {
 if($taskRecorder -and -not $taskRecorder.HasExited){$taskRecorder.Kill();$taskRecorder.WaitForExit()}
 if($Report){[IO.File]::WriteAllText([IO.Path]::GetFullPath($Report),($taskResult|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))}
 if($taskShow){$taskShow.View.Exit();[Runtime.InteropServices.Marshal]::ReleaseComObject($taskShow)|Out-Null}
 if($taskP){try{$taskP.Close()}catch{};[Runtime.InteropServices.Marshal]::ReleaseComObject($taskP)|Out-Null}
 [Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)|Out-Null
}
