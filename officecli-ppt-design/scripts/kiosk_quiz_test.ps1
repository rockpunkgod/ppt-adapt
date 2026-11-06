param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory,[Parameter(Mandatory=$true)][ValidateSet(4,5,6)][int]$CorrectAnswer,[Parameter(Mandatory=$true)][string]$Ffmpeg,[switch]$SpeakerControl)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $EvidenceDirectory){throw 'Fresh evidence directory required'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;using System.Text;using System.Runtime.InteropServices;
public static class QuizNative{
 public delegate bool Call(IntPtr h,IntPtr x);
 [StructLayout(LayoutKind.Sequential)]public struct R{public int L,T,Rt,B;}
 [StructLayout(LayoutKind.Sequential)]public struct P{public int X,Y;}
 [DllImport("user32.dll")]public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
 [DllImport("user32.dll")]public static extern bool EnumWindows(Call c,IntPtr x);
 [DllImport("user32.dll")]public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]public static extern int GetClassName(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]public static extern int GetWindowText(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]public static extern bool GetClientRect(IntPtr h,out R r);
 [DllImport("user32.dll")]public static extern bool ClientToScreen(IntPtr h,ref P p);
 [DllImport("user32.dll")]public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")]public static extern bool GetCursorPos(out P p);
 [DllImport("user32.dll")]public static extern IntPtr WindowFromPoint(P p);
 [DllImport("user32.dll")]public static extern IntPtr GetAncestor(IntPtr h,uint f);
 [DllImport("user32.dll")]public static extern void mouse_event(uint f,int x,int y,uint d,UIntPtr e);
 [DllImport("user32.dll")]public static extern void keybd_event(byte vk,byte scan,uint flags,UIntPtr e);
 [DllImport("dwmapi.dll")]public static extern int DwmGetWindowAttribute(IntPtr h,int a,out R r,int size);
 public static IntPtr Find(string name){IntPtr found=IntPtr.Zero;EnumWindows((h,x)=>{var c=new StringBuilder(256);var t=new StringBuilder(1024);GetClassName(h,c,256);GetWindowText(h,t,1024);if(IsWindowVisible(h)&&c.ToString().ToLower().Contains("screenclass")&&t.ToString().IndexOf(name,StringComparison.OrdinalIgnoreCase)>=0)found=h;return true;},IntPtr.Zero);return found;}
}
'@
[QuizNative]::SetThreadDpiAwarenessContext([IntPtr](-4))|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$show=$null;$rec=$null;$log=@();$states=@();$result=@{passed=$false};$colors=@{4='193C40';5='193C40';6='193C40'}
function HexRgb([string]$hex){return @([Convert]::ToInt32($hex.Substring(0,2),16),[Convert]::ToInt32($hex.Substring(2,2),16),[Convert]::ToInt32($hex.Substring(4,2),16))}
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1)
 if([int]$p.SlideShowSettings.ShowType -ne 3){throw 'Document is not saved in true Kiosk mode'}
 $q=$p.Slides.Item(1).Shapes.Item('question').TextFrame.TextRange.Text
 if($q -notmatch '(\d+)\s*\+\s*(\d+)' -or ([int]$Matches[1]+[int]$Matches[2]) -ne $CorrectAnswer){throw 'Question/answer mismatch'}
 $p.SlideShowSettings.RangeType=1;$p.SlideShowSettings.StartingSlide=1;$p.SlideShowSettings.EndingSlide=2;$p.SlideShowSettings.ShowPresenterView=0
 $show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 700
 if(-not $show){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
 $show.View.PointerType=1;$hwnd=[QuizNative]::Find($p.Name)
 if($hwnd -eq [IntPtr]::Zero){throw 'Own visible slideshow not found'}
 [QuizNative]::SetForegroundWindow($hwnd)|Out-Null
 $r=New-Object QuizNative+R;[QuizNative]::GetClientRect($hwnd,[ref]$r)|Out-Null
 $w=$r.Rt;$h=$r.B;$scale=[Math]::Min($w/33.87,$h/19.05);$ox=($w-33.87*$scale)/2;$oy=($h-19.05*$scale)/2
 $origin=New-Object QuizNative+P;[QuizNative]::ClientToScreen($hwnd,[ref]$origin)|Out-Null
 function Guard{if([QuizNative]::GetForegroundWindow() -ne $hwnd){throw 'Own slideshow lost foreground'}}
 function Click([double]$x,[double]$y,[string]$name,[int]$wait=250){
  Guard;$pt=New-Object QuizNative+P;$pt.X=[int]($origin.X+$ox+$x*$scale);$pt.Y=[int]($origin.Y+$oy+$y*$scale)
  if([QuizNative]::GetAncestor([QuizNative]::WindowFromPoint($pt),2) -ne $hwnd){throw 'Click target is outside own slideshow'}
  [QuizNative]::SetCursorPos($pt.X,$pt.Y)|Out-Null;[QuizNative]::mouse_event(1,1,0,0,[UIntPtr]::Zero);[QuizNative]::mouse_event(1,-1,0,0,[UIntPtr]::Zero);Start-Sleep -Milliseconds 50
  $actual=New-Object QuizNative+P;[QuizNative]::GetCursorPos([ref]$actual)|Out-Null
  if($actual.X -ne $pt.X -or $actual.Y -ne $pt.Y){throw 'Pointer did not return exactly to target'}
  Guard;$before=[int]$show.View.Slide.SlideIndex
  [QuizNative]::mouse_event(2,0,0,0,[UIntPtr]::Zero);[QuizNative]::mouse_event(4,0,0,0,[UIntPtr]::Zero);Start-Sleep -Milliseconds $wait
  $script:log+=@{name=$name;kind='physical-click';before=$before;after=[int]$show.View.Slide.SlideIndex;expectedPoint=@($pt.X,$pt.Y);actualPoint=@($actual.X,$actual.Y)}
 }
 function Key([byte]$vk,[string]$name){Guard;[QuizNative]::keybd_event($vk,0,0,[UIntPtr]::Zero);[QuizNative]::keybd_event($vk,0,2,[UIntPtr]::Zero);Start-Sleep -Milliseconds 100;$actual=[int]$show.View.Slide.SlideIndex;$script:log+=@{name=$name;kind='physical-key';actual=$actual;passed=($actual -eq 1)};if($actual -ne 1){throw ('Kiosk key bypass: '+$name)}}
 function Shot([string]$name,[bool]$nextVisible,[int]$expectedSlide=1){
  Guard;$bmp=New-Object Drawing.Bitmap($w,$h);$g=[Drawing.Graphics]::FromImage($bmp)
  try{
   $g.CopyFromScreen($origin.X,$origin.Y,0,0,$bmp.Size);$bmp.Save((Join-Path $EvidenceDirectory ($name+'.png')))
   $actualSlide=[int]$show.View.Slide.SlideIndex;$row=@{name=$name;expectedSlide=$expectedSlide;actualSlide=$actualSlide;nextExpected=$nextVisible;answerColors=@();passed=($actualSlide -eq $expectedSlide)}
   if($expectedSlide -eq 1){
    foreach($answer in @(4,5,6)){$i=$answer-4;$pixel=$bmp.GetPixel([int]($ox+(3+$i*10.1)*$scale),[int]($oy+9*$scale));$want=HexRgb $colors[$answer];$error=[Math]::Abs([int]$pixel.R-$want[0])+[Math]::Abs([int]$pixel.G-$want[1])+[Math]::Abs([int]$pixel.B-$want[2]);$row.answerColors+=@{answer=$answer;rgb=@([int]$pixel.R,[int]$pixel.G,[int]$pixel.B);expectedHex=$colors[$answer];passed=($error -le 6)};$row.passed=$row.passed -and ($error -le 6)}
    $pixel=$bmp.GetPixel([int]($ox+26.2*$scale),[int]($oy+15.7*$scale));$want=HexRgb $(if($nextVisible){'21645E'}else{'F1EEE5'});$error=[Math]::Abs([int]$pixel.R-$want[0])+[Math]::Abs([int]$pixel.G-$want[1])+[Math]::Abs([int]$pixel.B-$want[2]);$row.nextRgb=@([int]$pixel.R,[int]$pixel.G,[int]$pixel.B);$row.passed=$row.passed -and ($error -le 6)
   }
   $script:states+=$row;if(-not $row.passed){throw ('Native state failed: '+$name)}
  }finally{$g.Dispose();$bmp.Dispose()}
 }
 $physical=New-Object QuizNative+R;[QuizNative]::DwmGetWindowAttribute($hwnd,9,[ref]$physical,16)|Out-Null
 $movie=Join-Path $EvidenceDirectory 'quiz-live.mp4';$argsText='-hide_banner -loglevel error -f gdigrab -framerate 25 -offset_x '+$physical.L+' -offset_y '+$physical.T+' -video_size '+($physical.Rt-$physical.L)+'x'+($physical.B-$physical.T)+' -i desktop -t 8 -vf "scale=1280:-2" -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+[IO.Path]::GetFullPath($movie)+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $argsText -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $EvidenceDirectory 'recorder.txt');Start-Sleep -Milliseconds 400
 Shot 'initial' $false
 Click 28 16.3 'hidden-next-before' 200;Shot 'hidden-next-blocked' $false
 Click 1 17.5 'blank-before' 150;Shot 'blank-blocked' $false
 Key 0x27 'right';Key 0x28 'down';Key 0x25 'left';Key 0x26 'up';Key 0x22 'pagedown';Key 0x20 'space';Shot 'keys-blocked' $false
 $wrong=@(4,5,6)|Where-Object {$_ -ne $CorrectAnswer}
 foreach($answer in $wrong){Click (6.8+($answer-4)*10.1) 10.05 ('answer-'+$answer) 450;$colors[$answer]='AF4E45';Shot ('wrong-'+$answer) $false}
 Click 28 16.3 'hidden-next-after-wrong' 200;Shot 'wrong-does-not-unlock' $false
 Click (6.8+($CorrectAnswer-4)*10.1) 10.05 ('answer-'+$CorrectAnswer) 1050;$colors[$CorrectAnswer]='21645E';Shot 'correct-unlocks' $true
 Key 0x27 'right-after-correct';Shot 'correct-still-needs-next' $true
 Click 28 16.3 'visible-next' 500;Shot 'result-page' $true 2
 $rec.Handle|Out-Null;$rec.WaitForExit();if(-not(Test-Path -LiteralPath $movie) -or (Get-Item -LiteralPath $movie).Length -eq 0){throw 'Desktop movie missing/empty'}
 if($SpeakerControl){
  $show.View.Exit();$show=$null;$p.SlideShowSettings.ShowType=1;$show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 500
  if(-not $show){$show=$app.SlideShowWindows.Item($app.SlideShowWindows.Count)}
  $show.View.PointerType=1;$hwnd=[QuizNative]::Find($p.Name);[QuizNative]::SetForegroundWindow($hwnd)|Out-Null
  $r=New-Object QuizNative+R;[QuizNative]::GetClientRect($hwnd,[ref]$r)|Out-Null;$w=$r.Rt;$h=$r.B;$scale=[Math]::Min($w/33.87,$h/19.05);$ox=($w-33.87*$scale)/2;$oy=($h-19.05*$scale)/2;$origin=New-Object QuizNative+P;[QuizNative]::ClientToScreen($hwnd,[ref]$origin)|Out-Null
  Click 1 17.5 'speaker-blank-negative-control' 300;$actual=[int]$show.View.Slide.SlideIndex;$log+=@{name='speaker-control';expected=2;actual=$actual;passed=($actual -eq 2)};if($actual -ne 2){throw 'Speaker negative control did not advance'}
 }
 $result=@{passed=$true;file=[IO.Path]::GetFullPath($Deck);question=$q;correct=$CorrectAnswer;checks=$log;states=$states;width=$w;height=$h;scale=$scale;offset=@($ox,$oy);speakerControl=[bool]$SpeakerControl;boundary='Only physical clicks/keys drive the tested quiz; Kiosk is presentation navigation gating, not access control.'}
}catch{$result=@{passed=$false;error=$_.Exception.ToString();file=[IO.Path]::GetFullPath($Deck);correct=$CorrectAnswer;checks=$log;states=$states};throw}
finally{
 if($rec -and -not $rec.HasExited){$rec.Kill();$rec.WaitForExit()}
 if($show){try{$show.View.Exit()}catch{}}
 if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'report.json'),($result|ConvertTo-Json -Depth 9),(New-Object Text.UTF8Encoding($false)))
}
