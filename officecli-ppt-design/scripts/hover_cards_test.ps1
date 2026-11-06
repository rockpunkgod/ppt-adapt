param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$Directory,[Parameter(Mandatory=$true)][string]$Ffmpeg)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Directory){throw 'Fresh output directory required'}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($Directory))
Add-Type @'
using System;using System.Text;using System.Runtime.InteropServices;
public static class HoverNative {
 public delegate bool EnumProc(IntPtr h,IntPtr p);
 [StructLayout(LayoutKind.Sequential)] public struct RECT {public int Left,Top,Right,Bottom;}
 [StructLayout(LayoutKind.Sequential)] public struct POINT {public int X,Y;}
 [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
 [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc c,IntPtr p);
 [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT p);
 [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h,uint f);
 [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h,out RECT r);
 [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h,ref POINT p);
 [DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
 [DllImport("user32.dll")] public static extern void mouse_event(uint flags,int dx,int dy,uint data,UIntPtr extra);
 [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h,int key,out RECT r,int n);
 public static IntPtr Find(string name){IntPtr result=IntPtr.Zero;EnumWindows((h,p)=>{var c=new StringBuilder(256);var t=new StringBuilder(1024);GetClassName(h,c,256);GetWindowText(h,t,1024);if(IsWindowVisible(h)&&c.ToString().ToLower().Contains("screenclass")&&t.ToString().IndexOf(name,StringComparison.OrdinalIgnoreCase)>=0)result=h;return true;},IntPtr.Zero);return result;}
}
'@
[void][HoverNative]::SetThreadDpiAwarenessContext([IntPtr](-4))
Add-Type -AssemblyName System.Drawing
$app=New-Object -ComObject PowerPoint.Application;$pres=$null;$show=$null;$rec=$null;$log=@();$moves=@();$result=@{passed=$false}
try{
 $app.Visible=-1;$pres=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1)
 $pres.SlideShowSettings.RangeType=1;$pres.SlideShowSettings.ShowType=1;$pres.SlideShowSettings.ShowPresenterView=0;$pres.SlideShowSettings.AdvanceMode=1
 $show=$pres.SlideShowSettings.Run();Start-Sleep -Milliseconds 400
 if(-not $show){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
 $show.View.PointerType=1;$hwnd=[HoverNative]::Find($pres.Name)
 if($hwnd -eq [IntPtr]::Zero){throw 'Own slideshow not found'}
 [void][HoverNative]::SetForegroundWindow($hwnd)
 $r=New-Object HoverNative+RECT;[void][HoverNative]::GetClientRect($hwnd,[ref]$r)
 $width=$r.Right;$height=$r.Bottom;$scale=[Math]::Min($width/33.87,$height/19.05)
 $ox=($width-33.87*$scale)/2;$oy=($height-19.05*$scale)/2
 $origin=New-Object HoverNative+POINT;[void][HoverNative]::ClientToScreen($hwnd,[ref]$origin)
 function Move-To([double]$X,[double]$Y){
  if([HoverNative]::GetForegroundWindow() -ne $hwnd){throw 'Own slideshow lost foreground'}
  $point=New-Object HoverNative+POINT;$point.X=[int]($origin.X+$ox+$X*$scale);$point.Y=[int]($origin.Y+$oy+$Y*$scale)
  if([HoverNative]::GetAncestor([HoverNative]::WindowFromPoint($point),2) -ne $hwnd){throw 'Pointer target is outside own slideshow'}
  if(-not [HoverNative]::SetCursorPos($point.X,$point.Y)){throw 'Pointer positioning failed'}
  [HoverNative]::mouse_event(1,1,0,0,[UIntPtr]::Zero)
  [HoverNative]::mouse_event(1,-1,0,0,[UIntPtr]::Zero)
  Start-Sleep -Milliseconds 50
  $actualPoint=New-Object HoverNative+POINT;[void][HoverNative]::GetCursorPos([ref]$actualPoint)
  $script:moves+=@{expected=@($point.X,$point.Y);actual=@($actualPoint.X,$actualPoint.Y);buttonsInjected=$false}
  if($actualPoint.X -ne $point.X -or $actualPoint.Y -ne $point.Y){throw 'Actual pointer position differs from own target'}
 }
 Move-To 32 18;$show.View.GotoSlide(1);Start-Sleep -Milliseconds 2000
 $physical=New-Object HoverNative+RECT
 if([HoverNative]::DwmGetWindowAttribute($hwnd,9,[ref]$physical,16) -ne 0){throw 'No physical screen bounds'}
 $movie=Join-Path $Directory 'hover-live.mp4';$vw=$physical.Right-$physical.Left;$vh=$physical.Bottom-$physical.Top
 $argsText='-hide_banner -loglevel error -f gdigrab -framerate 25 -offset_x '+$physical.Left+' -offset_y '+$physical.Top+' -video_size '+$vw+'x'+$vh+' -i desktop -t 12 -vf "scale=1280:-2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+[IO.Path]::GetFullPath($movie)+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $argsText -PassThru -WindowStyle Hidden -RedirectStandardError (Join-Path $Directory 'recorder.txt')
 $clock=[Diagnostics.Stopwatch]::StartNew();Start-Sleep -Milliseconds 500
 function Check-State([string]$Name,[int]$Expected){
  if([HoverNative]::GetForegroundWindow() -ne $hwnd){throw 'Own slideshow lost foreground'}
  $actual=[int]$show.View.Slide.SlideIndex
  $bmp=New-Object Drawing.Bitmap($width,$height);$g=[Drawing.Graphics]::FromImage($bmp)
  try{
   $g.CopyFromScreen($origin.X,$origin.Y,0,0,$bmp.Size);$bmp.Save((Join-Path $Directory ($Name+'.png')))
   $colors=@();$pixelPass=$true
   for($j=1;$j -le 4;$j++){
    $rgb=$bmp.GetPixel([int]($ox+(1.7+($j-1)*7.8)*$scale),[int]($oy+7*$scale))
    $want=if($Expected -eq $j+1){@(28,101,97)}else{@(213,226,219)}
    $error=[Math]::Abs([int]$rgb.R-$want[0])+[Math]::Abs([int]$rgb.G-$want[1])+[Math]::Abs([int]$rgb.B-$want[2])
    $pixelPass=$pixelPass -and ($error -le 6);$colors+=@{card=$j;rgb=@([int]$rgb.R,[int]$rgb.G,[int]$rgb.B);error=$error}
   }
  }finally{$g.Dispose();$bmp.Dispose()}
  $script:log+=@{name=$Name;expected=$Expected;actual=$actual;colors=$colors;passed=($actual -eq $Expected -and $pixelPass);at=$clock.Elapsed.TotalSeconds;input='pointer movement only'}
  if($actual -ne $Expected -or -not $pixelPass){throw "Hover state failed: $Name, actual $actual"}
 }
 Check-State 'initial-blank' 1
 $n=0
 foreach($card in @(1,2,3,4,1,4,2)){
  $n++;Move-To (5+($card-1)*7.8) 10
  $wait=[Diagnostics.Stopwatch]::StartNew()
  while([int]$show.View.Slide.SlideIndex -ne $card+1 -and $wait.Elapsed.TotalSeconds -lt 2){Start-Sleep -Milliseconds 100}
  Start-Sleep -Milliseconds 100;Check-State ("hover-$n-card-$card") ($card+1)
  Move-To 32 18;Start-Sleep -Milliseconds 250;Check-State ("blank-$n") ($card+1)
 }
 while($clock.Elapsed.TotalSeconds -lt 12.2){if([HoverNative]::GetForegroundWindow() -ne $hwnd){throw 'Own slideshow lost foreground'};Start-Sleep -Milliseconds 100}
 $rec.Handle|Out-Null;if(-not $rec.WaitForExit(10000)){throw 'Recorder timed out'}
 if(-not(Test-Path -LiteralPath $movie) -or (Get-Item -LiteralPath $movie).Length -eq 0){throw 'Recording missing'}
 $result=@{passed=$true;events=$log;moves=$moves;client=@($width,$height);origin=@($origin.X,$origin.Y);physical=@($physical.Left,$physical.Top,$vw,$vh);limit='No clicks, GotoSlide used only once before baseline. Seven actual hover entries, all four targets, repeated branches and blank negative checks.'}
}catch{$result=@{passed=$false;events=$log;moves=$moves;client=@($width,$height);origin=@($origin.X,$origin.Y);physical=@($physical.Left,$physical.Top,$vw,$vh);error=$_.Exception.Message};throw}
finally{
 if($rec -and -not $rec.HasExited){$rec.Kill();$rec.WaitForExit()}
 $result|ConvertTo-Json -Depth 9|Set-Content -LiteralPath (Join-Path $Directory 'report.json') -Encoding utf8
 if($show){$show.View.Exit()};if($pres){$pres.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
