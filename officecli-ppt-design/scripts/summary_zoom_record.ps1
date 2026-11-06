param([string]$Deck,[string]$Plan,[string]$Evidence,[string]$Ffmpeg)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'summary_zoom_ready.ps1')
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$model=Get-Content -LiteralPath $Plan -Raw -Encoding UTF8|ConvertFrom-Json
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @'
using System;using System.Text;using System.Collections.Generic;using System.Runtime.InteropServices;
public static class SummaryRecordWin{
 [StructLayout(LayoutKind.Sequential)]public struct RECT{public int L,T,R,B;}
 [StructLayout(LayoutKind.Sequential)]public struct POINT{public int X,Y;}
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]public static extern bool SetWindowPos(IntPtr h,IntPtr after,int x,int y,int w,int hh,uint flags);
 [DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")]public static extern bool GetCursorPos(out POINT p);
 [DllImport("user32.dll")]public static extern void mouse_event(uint f,uint x,uint y,uint d,UIntPtr e);
 [DllImport("user32.dll")]public static extern IntPtr WindowFromPoint(POINT p);
 [DllImport("user32.dll")]public static extern IntPtr GetAncestor(IntPtr h,uint f);
 [DllImport("user32.dll")]public static extern bool GetClientRect(IntPtr h,out RECT r);
 [DllImport("user32.dll")]public static extern bool ClientToScreen(IntPtr h,ref POINT p);
 [DllImport("dwmapi.dll")]public static extern int DwmGetWindowAttribute(IntPtr h,int a,out RECT r,int s);
 [DllImport("user32.dll")]public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr v);
 public delegate bool Call(IntPtr h,IntPtr x);
 [DllImport("user32.dll")]public static extern bool EnumWindows(Call c,IntPtr x);
 [DllImport("user32.dll")]public static extern int GetClassName(IntPtr h,StringBuilder s,int n);
 [DllImport("user32.dll")]public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint pid);
 public static IntPtr Find(uint pid){var found=new List<IntPtr>();EnumWindows((h,x)=>{var c=new StringBuilder(256);GetClassName(h,c,256);uint p;GetWindowThreadProcessId(h,out p);if(p==pid&&IsWindowVisible(h)&&c.ToString().ToLower().Contains("screenclass"))found.Add(h);return true;},IntPtr.Zero);if(found.Count!=1)throw new Exception("Expected exactly one visible own projection; found "+found.Count);return found[0];}
}
'@
$oldDpi=[SummaryRecordWin]::SetThreadDpiAwarenessContext([IntPtr](-4))
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$show=$null;$rec=$null;$log=@();$before=(Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1)
 $leaf=[IO.Path]::GetFileName($Deck)
 $own=@([Windows.Automation.AutomationElement]::RootElement.FindAll([Windows.Automation.TreeScope]::Children,[Windows.Automation.Condition]::TrueCondition)|Where-Object {$_.Current.ClassName -eq 'PPTFrameClass' -and $_.Current.Name -like ('*'+$leaf+'*')})
 if($own.Count -ne 1){throw 'Own normal window not uniquely identified'}
 $ownPid=[uint32]$own[0].Current.ProcessId
 try{$own[0].SetFocus()}catch{}
 $taskShell=New-Object -ComObject WScript.Shell;$normalActivated=$taskShell.AppActivate($own[0].Current.Name);Start-Sleep -Milliseconds 150
 $posterReady=Wait-SummaryPosters $p.Slides.Item(1) ([IO.Path]::GetFullPath($Evidence)) 'poster-ready' 3
 for($i=1;$i -le $p.Slides.Count;$i++){$p.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('native-'+$i+'.png')),'PNG',1600,900)}
 $p.SlideShowSettings.ShowType=1;$p.SlideShowSettings.ShowPresenterView=0;$p.SlideShowSettings.AdvanceMode=1
 $show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 650
 if(-not $show -and $app.SlideShowWindows.Count -gt 0){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
 if(-not $show -or $show.Presentation.FullName -ne $p.FullName){throw 'No own slideshow'}
 $show.View.PointerType=1;if([int]$show.View.PointerType -ne 1){throw 'Arrow pointer did not persist'}
 $h=[SummaryRecordWin]::Find($ownPid);$screen=@([Windows.Forms.Screen]::AllScreens|Where-Object {-not $_.Primary})[0];if(-not $screen){throw 'Second screen required for guarded non-overlapping test'};$work=$screen.WorkingArea;$moved=[SummaryRecordWin]::SetWindowPos($h,[IntPtr]::Zero,$work.Left+40,$work.Top+40,1280,760,68);Start-Sleep -Milliseconds 250;$foregroundSet=[SummaryRecordWin]::SetForegroundWindow($h);Start-Sleep -Milliseconds 200
 if([SummaryRecordWin]::GetForegroundWindow() -ne $h){throw 'Own slideshow not foreground; no inputs sent'}
 $r=New-Object SummaryRecordWin+RECT;[SummaryRecordWin]::DwmGetWindowAttribute($h,9,[ref]$r,16)|Out-Null
 $client=New-Object SummaryRecordWin+RECT;[SummaryRecordWin]::GetClientRect($h,[ref]$client)|Out-Null
 $cw=$client.R-$client.L;$ch=$client.B-$client.T;$scale=[Math]::Min($cw/$model.widthCm,$ch/$model.heightCm);$ox=($cw-$model.widthCm*$scale)/2;$oy=($ch-$model.heightCm*$scale)/2
 $width=$r.R-$r.L;$height=$r.B-$r.T
 $movie=Join-Path ([IO.Path]::GetFullPath($Evidence)) 'summary-live.mp4'
 $taskRecordArgs='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$r.L+' -offset_y '+$r.T+' -video_size '+$width+'x'+$height+' -i desktop -t 16 -vf "scale=1600:-2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$movie+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $taskRecordArgs -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg-error.txt')
 function Capture([string]$name){$b=New-Object Drawing.Bitmap($width,$height);$g=[Drawing.Graphics]::FromImage($b);try{$g.CopyFromScreen($r.L,$r.T,0,0,$b.Size);$b.Save((Join-Path $Evidence ($name+'.png')),[Drawing.Imaging.ImageFormat]::Png)}finally{$g.Dispose();$b.Dispose()}}
 $clock=[Diagnostics.Stopwatch]::StartNew();Start-Sleep -Milliseconds 750;Capture 'start'
 foreach($step in $model.steps){
  if([SummaryRecordWin]::GetForegroundWindow() -ne $h){throw 'Own show lost foreground'}
  $point=New-Object SummaryRecordWin+POINT;$point.X=[int]($ox+$step.x*$scale);$point.Y=[int]($oy+$step.y*$scale);[SummaryRecordWin]::ClientToScreen($h,[ref]$point)|Out-Null
  if([SummaryRecordWin]::GetAncestor([SummaryRecordWin]::WindowFromPoint($point),2) -ne $h){throw 'Target point is not in own projection; no input sent'}
  [SummaryRecordWin]::SetCursorPos($point.X,$point.Y)|Out-Null;Start-Sleep -Milliseconds 50
  $actualCursor=New-Object SummaryRecordWin+POINT;[SummaryRecordWin]::GetCursorPos([ref]$actualCursor)|Out-Null
  if([Math]::Abs($actualCursor.X-$point.X)+[Math]::Abs($actualCursor.Y-$point.Y) -gt 1){throw 'Cursor mismatch'}
  $eventTime=$clock.Elapsed.TotalSeconds
  [SummaryRecordWin]::mouse_event(2,0,0,0,[UIntPtr]::Zero);[SummaryRecordWin]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
  Start-Sleep -Milliseconds 1300;Capture $step.name
  $actual=[int]$show.View.Slide.SlideIndex;$log+=@{name=$step.name;inputTime=$eventTime;observedTime=$clock.Elapsed.TotalSeconds;expected=[int]$step.expected;actual=$actual;physicalPoint=@($point.X,$point.Y);passed=($actual -eq [int]$step.expected)}
  if($actual -ne [int]$step.expected){throw ('Expected '+$step.expected+' actual '+$actual+' at '+$step.name)}
 }
 if(-not $rec.WaitForExit(10000)){throw 'Own recorder timeout'}
 if(-not(Test-Path -LiteralPath $movie) -or (Get-Item -LiteralPath $movie).Length -eq 0){throw 'Movie missing/empty'}
 @{passed=$true;events=$log;ownPid=$ownPid;normalAppActivate=$normalActivated;window=[int64]$h;foregroundSet=$foregroundSet;physicalDwm=@($r.L,$r.T,$r.R,$r.B);client=@($cw,$ch);contentRect=@($ox,$oy,($model.widthCm*$scale),($model.heightCm*$scale));ownWindowMoved=$moved;deckShaBefore=$before;scope='Nine real left clicks, all three native Summary sections and automatic section-end return'}|ConvertTo-Json -Depth 9|Set-Content -LiteralPath (Join-Path $Evidence 'clicks.json') -Encoding UTF8
}catch{@{passed=$false;error=$_.Exception.Message;events=$log}|ConvertTo-Json -Depth 9|Set-Content -LiteralPath (Join-Path $Evidence 'failure.json') -Encoding UTF8;throw}
finally{if($rec -and -not $rec.HasExited){$rec.Kill();$rec.WaitForExit()};if($show){try{$show.View.Exit()}catch{}};if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app);[SummaryRecordWin]::SetThreadDpiAwarenessContext($oldDpi)|Out-Null}
if((Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash -ne $before){throw 'Read-only recording changed deck'}
