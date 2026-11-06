param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory,[Parameter(Mandatory=$true)][string]$Ffmpeg)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $EvidenceDirectory){throw 'Fresh evidence directory required'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
Add-Type -AssemblyName System.Drawing
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
 [DllImport("user32.dll")] public static extern void keybd_event(byte vk,byte scan,uint flags,UIntPtr extra);
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
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$show=$null;$rec=$null;$events=@();$players=@();$keys=@();$result=@{passed=$false}
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1);$v=$p.Slides.Item(2).Shapes.Item('screen-video');$videoID=$v.Id
 if([int]$v.Type -ne 16 -or [int]$v.MediaType -ne 3){throw 'True video required'}
 $p.SlideShowSettings.RangeType=1;$p.SlideShowSettings.ShowType=1;$p.SlideShowSettings.AdvanceMode=1;$p.SlideShowSettings.ShowPresenterView=0
 $show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 500
 if(-not $show){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
 $show.View.PointerType=1;$hwnd=[HoverNative]::Find($p.Name)
 if($hwnd -eq [IntPtr]::Zero){throw 'Own show missing'}
 [void][HoverNative]::SetForegroundWindow($hwnd)
 $client=New-Object HoverNative+RECT;[void][HoverNative]::GetClientRect($hwnd,[ref]$client);$width=$client.Right;$height=$client.Bottom
 $origin=New-Object HoverNative+POINT;[void][HoverNative]::ClientToScreen($hwnd,[ref]$origin)
 $physical=New-Object HoverNative+RECT;if([HoverNative]::DwmGetWindowAttribute($hwnd,9,[ref]$physical,16) -ne 0){throw 'No physical bounds'}
 $scale=[Math]::Min($width/33.87,$height/19.05);$ox=($width-33.87*$scale)/2;$oy=($height-19.05*$scale)/2
 function Guard{if([HoverNative]::GetForegroundWindow() -ne $hwnd){throw 'Own slideshow lost foreground'}}
 function Sample-Player{
  Guard;$index=[int]$show.View.Slide.SlideIndex
  if($index -eq 2){
   try{$pl=$show.View.Player($videoID);$script:players+=@{at=$clock.Elapsed.TotalSeconds;slide=$index;positionMs=$pl.CurrentPosition;state=[int]$pl.State;available=$true}}
   catch{$script:players+=@{at=$clock.Elapsed.TotalSeconds;slide=$index;available=$false;error=$_.Exception.HResult}}
  }
 }
 function Wait-To([double]$t){while($clock.Elapsed.TotalSeconds -lt $t){Sample-Player;Start-Sleep -Milliseconds 80}}
 function Shot([string]$name,[int]$expected){
  Guard;$actual=[int]$show.View.Slide.SlideIndex
  $b=New-Object Drawing.Bitmap($width,$height);$g=[Drawing.Graphics]::FromImage($b)
  try{$g.CopyFromScreen($origin.X,$origin.Y,0,0,$b.Size);$b.Save((Join-Path $EvidenceDirectory ($name+'.png')))}finally{$g.Dispose();$b.Dispose()}
  $script:events+=@{name=$name;at=$clock.Elapsed.TotalSeconds;actual=$actual;expected=$expected;passed=($actual -eq $expected)}
  if($actual -ne $expected){throw "Unexpected actual page for $name"}
 }
 function Key([byte]$vk,[string]$name){Guard;$before=[int]$show.View.Slide.SlideIndex;[HoverNative]::keybd_event($vk,0,0,[UIntPtr]::Zero);[HoverNative]::keybd_event($vk,0,2,[UIntPtr]::Zero);$script:keys+=@{name=$name;key=$vk;at=$clock.Elapsed.TotalSeconds;before=$before;input='physical keyboard'}}
 $movie=Join-Path $EvidenceDirectory 'laptop-live.mp4';$argsText='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$physical.Left+' -offset_y '+$physical.Top+' -video_size '+($physical.Right-$physical.Left)+'x'+($physical.Bottom-$physical.Top)+' -i desktop -t 22 -vf "scale=1280:-2" -c:v libx264 -preset ultrafast -g 30 -flush_packets 1 -pix_fmt yuv420p -movflags +frag_keyframe+empty_moov+default_base_moof "'+[IO.Path]::GetFullPath($movie)+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $argsText -PassThru -WindowStyle Hidden -RedirectStandardError (Join-Path $EvidenceDirectory 'recorder.txt');$clock=[Diagnostics.Stopwatch]::StartNew()
 Wait-To 0.7;Shot 'closed-before' 1
 Wait-To 1.0;Key 0x27 'open-first'
 Wait-To 3.3;Shot 'open-fade-first' 2
 Wait-To 4.4;Shot 'open-live-first' 2
 Wait-To 8.8;Shot 'open-live-late' 2
 Wait-To 9.4;Key 0x24 'close-home'
 Wait-To 11.8;Shot 'closed-after-video' 1
 Wait-To 12.2;Key 0x27 'open-second'
 Wait-To 14.5;Shot 'open-fade-second' 2
 Wait-To 15.6;Shot 'open-live-second' 2
 Wait-To 20.8;Shot 'open-second-late' 2
 Wait-To 22.2;if(-not $rec.WaitForExit(5000)){throw 'Recorder exceeded time bound'}
 $result=@{passed=$true;events=$events;players=$players;keys=$keys;videoId=$videoID;videoLengthMs=$v.MediaFormat.Length;client=@($width,$height);physical=@($physical.Left,$physical.Top,($physical.Right-$physical.Left),($physical.Bottom-$physical.Top));scale=$scale;offset=@($ox,$oy);limit='Right/Home/Right physical keys open, close and reopen the real 2.5D scene. Player samples are read-only, never seek/start programmatically.'}
}catch{$result=@{passed=$false;events=$events;players=$players;keys=$keys;error=$_.Exception.ToString()};throw}
finally{
 if($rec -and -not $rec.HasExited){$rec.Kill();$rec.WaitForExit()}
 if($show){try{$show.View.Exit()}catch{}};if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'report.json'),($result|ConvertTo-Json -Depth 10),(New-Object Text.UTF8Encoding($false)))
}
