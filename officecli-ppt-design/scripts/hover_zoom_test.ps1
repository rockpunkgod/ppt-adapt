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
 public static int[][] TrackMove(IntPtr expected,int tx,int ty){
  var samples=new System.Collections.Generic.List<int[]>();var clock=System.Diagnostics.Stopwatch.StartNew();
  while(true){
   if(GetForegroundWindow()!=expected)throw new Exception("Lost own foreground during trajectory");
   POINT cur;GetCursorPos(out cur);samples.Add(new int[]{cur.X,cur.Y,(int)clock.ElapsedMilliseconds});
   int dx=tx-cur.X,dy=ty-cur.Y;if(Math.Abs(dx)<=1&&Math.Abs(dy)<=1)break;
   if(clock.ElapsedMilliseconds>8000)throw new Exception("Trajectory exceeded 8 second bound");
   dx=Math.Max(-4,Math.Min(4,dx));dy=Math.Max(-4,Math.Min(4,dy));POINT next=new POINT();next.X=cur.X+dx;next.Y=cur.Y+dy;
   if(GetAncestor(WindowFromPoint(next),2)!=expected)throw new Exception("Trajectory leaves own show");
   mouse_event(1,dx,dy,0,UIntPtr.Zero);System.Threading.Thread.Sleep(2);
  }
  return samples.ToArray();
 }

}
'@
[void][HoverNative]::SetThreadDpiAwarenessContext([IntPtr](-4))
Add-Type -AssemblyName System.Drawing
$app=New-Object -ComObject PowerPoint.Application;$pres=$null;$show=$null;$rec=$null;$log=@();$moves=@();$result=@{passed=$false}
try{
 $app.Visible=-1;$pres=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1)
 $pres.SlideShowSettings.RangeType=1;$pres.SlideShowSettings.ShowType=1;$pres.SlideShowSettings.ShowPresenterView=0;$pres.SlideShowSettings.AdvanceMode=1
 $show=$pres.SlideShowSettings.Run();Start-Sleep -Milliseconds 500
 if(-not $show){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
 $show.View.PointerType=1;$hwnd=[HoverNative]::Find($pres.Name)
 if($hwnd -eq [IntPtr]::Zero){throw 'Own slideshow missing'}
 [void][HoverNative]::SetForegroundWindow($hwnd)
 $r=New-Object HoverNative+RECT;[void][HoverNative]::GetClientRect($hwnd,[ref]$r)
 $width=$r.Right;$height=$r.Bottom;$scale=[Math]::Min($width/33.87,$height/19.05);$ox=($width-33.87*$scale)/2;$oy=($height-19.05*$scale)/2
 $origin=New-Object HoverNative+POINT;[void][HoverNative]::ClientToScreen($hwnd,[ref]$origin)
 function Move-To([double]$X,[double]$Y){
  if($clock -and $clock.Elapsed.TotalSeconds -ge 120){throw "Input deadline reached; no further tested events"}
  $pt=New-Object HoverNative+POINT;$pt.X=[int]($origin.X+$ox+$X*$scale);$pt.Y=[int]($origin.Y+$oy+$Y*$scale)
  if([HoverNative]::GetAncestor([HoverNative]::WindowFromPoint($pt),2) -ne $hwnd){throw 'Target outside own show'}
  $trace=[HoverNative]::TrackMove($hwnd,$pt.X,$pt.Y);$actual=New-Object HoverNative+POINT;[HoverNative]::GetCursorPos([ref]$actual)|Out-Null
  if([Math]::Abs($pt.X-$actual.X) -gt 1 -or [Math]::Abs($pt.Y-$actual.Y) -gt 1){throw 'Trajectory final cursor mismatch'}
  $pix=New-Object Drawing.Bitmap(1,1);$g=[Drawing.Graphics]::FromImage($pix);try{$g.CopyFromScreen($actual.X,$actual.Y,0,0,$pix.Size);$c=$pix.GetPixel(0,0);$rgb=@([int]$c.R,[int]$c.G,[int]$c.B)}finally{$g.Dispose();$pix.Dispose()}
  $script:moves+=@{cm=@($X,$Y);expected=@($pt.X,$pt.Y);actual=@($actual.X,$actual.Y);trace=$trace;input='small relative mouse MOVE only';rgbAtActual=$rgb;actualInsideSlide=($actual.X -ge $origin.X+$ox -and $actual.X -lt $origin.X+$width-$ox -and $actual.Y -ge $origin.Y+$oy -and $actual.Y -lt $origin.Y+$height-$oy)}
 }
 function Click-At([double]$X,[double]$Y){Move-To $X $Y;if([HoverNative]::GetForegroundWindow() -ne $hwnd){throw 'Lost foreground before click'};[HoverNative]::mouse_event(2,0,0,0,[UIntPtr]::Zero);[HoverNative]::mouse_event(4,0,0,0,[UIntPtr]::Zero)}
 function State([string]$Name,[int]$Expected){
  if($clock -and $clock.Elapsed.TotalSeconds -ge 120){throw "Input deadline reached before next state"}
  $wait=[Diagnostics.Stopwatch]::StartNew();while([int]$show.View.Slide.SlideIndex -ne $Expected -and $wait.Elapsed.TotalSeconds -lt 2){Start-Sleep -Milliseconds 40}
  Start-Sleep -Milliseconds 600
  $samples=@();for($j=0;$j -lt 6;$j++){$samples+=[int]$show.View.Slide.SlideIndex;Start-Sleep -Milliseconds 30}
  $ok=(@($samples|Where-Object{$_ -ne $Expected}).Count -eq 0)
  $bmp=New-Object Drawing.Bitmap($width,$height);$g=[Drawing.Graphics]::FromImage($bmp)
  try{$g.CopyFromScreen($origin.X,$origin.Y,0,0,$bmp.Size);$bmp.Save((Join-Path $Directory ($Name+'.png')))}finally{$g.Dispose();$bmp.Dispose()}
  $script:log+=@{name=$Name;expected=$Expected;samples=$samples;passed=$ok;at=$clock.Elapsed.TotalSeconds}
  if(-not $ok){throw "Unexpected hover/Zoom state $Name expected $Expected got $samples"}
 }
 $initial=New-Object HoverNative+POINT;$initial.X=[int]($origin.X+$ox+2*$scale);$initial.Y=[int]($origin.Y+$oy+6*$scale)
 if([HoverNative]::GetAncestor([HoverNative]::WindowFromPoint($initial),2) -ne $hwnd){throw 'Initial neutral point outside own show'}
 [HoverNative]::SetCursorPos($initial.X,$initial.Y)|Out-Null;[HoverNative]::mouse_event(1,1,0,0,[UIntPtr]::Zero);[HoverNative]::mouse_event(1,-1,0,0,[UIntPtr]::Zero)
 $physical=New-Object HoverNative+RECT;if([HoverNative]::DwmGetWindowAttribute($hwnd,9,[ref]$physical,16) -ne 0){throw 'No physical bounds'}
 $vw=$physical.Right-$physical.Left;$vh=$physical.Bottom-$physical.Top;$movie=Join-Path $Directory 'hover-zoom-live.mp4'
 $argsText='-hide_banner -loglevel error -f gdigrab -framerate 25 -offset_x '+$physical.Left+' -offset_y '+$physical.Top+' -video_size '+$vw+'x'+$vh+' -i desktop -t 120 -vf "scale=1280:-2" -c:v libx264 -preset ultrafast -g 25 -flush_packets 1 -pix_fmt yuv420p -movflags +frag_keyframe+empty_moov+default_base_moof "'+[IO.Path]::GetFullPath($movie)+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $argsText -PassThru -WindowStyle Hidden -RedirectStandardError (Join-Path $Directory 'recorder.txt');$clock=[Diagnostics.Stopwatch]::StartNew();Start-Sleep -Milliseconds 300
 State 'initial-home' 1
 foreach($run in @(1,2,3,4)){
  $branch=if($run%2 -eq 1){1}else{2};$x=if($branch -eq 1){4.0}else{24.0}
  # Approach through the upper neutral strip, avoiding the other card en route.
  $cur=New-Object HoverNative+POINT;[HoverNative]::GetCursorPos([ref]$cur)|Out-Null;$cmx=($cur.X-$origin.X-$ox)/$scale
  if($run -gt 1){Move-To $cmx 4;Move-To $x 4}
  Move-To $x 6;State ("$run-hover-$branch") ($branch+1)
  Move-To $x 11;State ("$run-inside-hole-$branch") ($branch+1)
  Click-At $x 11;State ("$run-zoom-in-$branch") ($branch+3)
  $returnX=if($branch -eq 1){4.0}else{19.3}
  Click-At $returnX 15.2;State ("$run-zoom-return-$branch") ($branch+1)
  $outsideX=if($branch -eq 1){1.0}else{32.0}
  Move-To $outsideX 11;State ("$run-background-home") 1
 }
 while($clock.Elapsed.TotalSeconds -lt 120.2){if([HoverNative]::GetForegroundWindow() -ne $hwnd){throw 'Lost foreground during recording'};Start-Sleep -Milliseconds 100}
 if(-not $rec.WaitForExit(10000)){throw 'Recorder not finished'}
 $result=@{passed=$true;events=$log;moves=$moves;physical=@($physical.Left,$physical.Top,$vw,$vh);client=@($width,$height);scale=$scale;offset=@($ox,$oy);limit='Original C04 combined hover/Slide Zoom, two branches repeated twice with continuous paths. One neutral initial SetCursorPos only; all tested movement uses logged small-step relative mouse MOVE without any later pointer warp. Final point tolerance 1 physical pixel. No GotoSlide/Next or later SetCursorPos; all tested transitions driven by mouse MOVE/click.';slidePhysical=@(($origin.X+$ox),($origin.Y+$oy),($width-2*$ox),($height-2*$oy));documentAspect=([double]$pres.PageSetup.SlideWidth/[double]$pres.PageSetup.SlideHeight);showComBounds=@($show.Left,$show.Top,$show.Width,$show.Height)}
}catch{$result=@{passed=$false;events=$log;moves=$moves;error=$_.Exception.ToString()};throw}
finally{
 if($rec -and -not $rec.HasExited){$rec.Kill();$rec.WaitForExit()}
 [IO.File]::WriteAllText((Join-Path $Directory 'report.json'),($result|ConvertTo-Json -Depth 9),(New-Object Text.UTF8Encoding($false)))
 if($show){try{$show.View.Exit()}catch{}};if($pres){$pres.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
