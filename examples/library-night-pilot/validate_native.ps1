param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$Plan,[Parameter(Mandatory=$true)][string]$Spec,[Parameter(Mandatory=$true)][string]$Evidence,[Parameter(Mandatory=$true)][string]$Ffmpeg)
$ErrorActionPreference='Stop'
$Deck=[IO.Path]::GetFullPath($Deck);$Plan=[IO.Path]::GetFullPath($Plan);$Spec=[IO.Path]::GetFullPath($Spec);$Evidence=[IO.Path]::GetFullPath($Evidence)
foreach($f in @($Deck,$Plan,$Spec,$Ffmpeg)){if(-not(Test-Path -LiteralPath $f -PathType Leaf)){throw ('Missing input '+$f)}}
if(Test-Path -LiteralPath $Evidence){throw 'Refusing existing evidence directory'}
[void][IO.Directory]::CreateDirectory($Evidence)
$design=Get-Content -LiteralPath $Spec -Raw -Encoding UTF8|ConvertFrom-Json;$model=Get-Content -LiteralPath $Plan -Raw -Encoding UTF8|ConvertFrom-Json
if($model.pages.Count -ne 8){throw 'Expected eight-page plan'}
$hashes=@{};foreach($entry in @{deck=$Deck;plan=$Plan;spec=$Spec}.GetEnumerator()){$hashes[$entry.Key]=(Get-FileHash -LiteralPath $entry.Value -Algorithm SHA256).Hash}
Add-Type -AssemblyName System.Drawing;Add-Type -AssemblyName System.Windows.Forms;Add-Type -AssemblyName UIAutomationClient;Add-Type -AssemblyName UIAutomationTypes
Add-Type @'
using System;using System.Text;using System.Collections.Generic;using System.Runtime.InteropServices;
public static class PilotWin{
 [StructLayout(LayoutKind.Sequential)]public struct R{public int L,T,Rt,B;}
 [StructLayout(LayoutKind.Sequential)]public struct P{public int X,Y;}
 [DllImport("user32.dll")]public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr v);
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]public static extern bool SetWindowPos(IntPtr h,IntPtr a,int x,int y,int w,int hh,uint f);
 [DllImport("user32.dll")]public static extern bool GetClientRect(IntPtr h,out R r);
 [DllImport("user32.dll")]public static extern bool ClientToScreen(IntPtr h,ref P p);
 [DllImport("user32.dll")]public static extern IntPtr WindowFromPoint(P p);
 [DllImport("user32.dll")]public static extern IntPtr GetAncestor(IntPtr h,uint f);
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")]public static extern bool IsWindowEnabled(IntPtr h);
 [DllImport("user32.dll")]public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")]public static extern bool IsWindow(IntPtr h);
 [DllImport("user32.dll")]public static extern int GetClassName(IntPtr h,StringBuilder s,int n);
 [DllImport("user32.dll")]public static extern void keybd_event(byte k,byte scan,uint f,UIntPtr e);
 [DllImport("dwmapi.dll")]public static extern int DwmGetWindowAttribute(IntPtr h,int a,out R r,int s);
 public delegate bool CB(IntPtr h,IntPtr p);
 [DllImport("user32.dll")]public static extern bool EnumWindows(CB c,IntPtr p);
 public static IntPtr Find(uint pid){var a=new List<IntPtr>();EnumWindows((h,p)=>{uint n;GetWindowThreadProcessId(h,out n);var s=new StringBuilder(256);GetClassName(h,s,256);if(n==pid&&IsWindowVisible(h)&&s.ToString().Equals("screenClass",StringComparison.OrdinalIgnoreCase))a.Add(h);return true;},IntPtr.Zero);if(a.Count!=1)throw new Exception("Own visible screenClass count "+a.Count);return a[0];}
}
'@
$oldDpi=[PilotWin]::SetThreadDpiAwarenessContext([IntPtr](-4));$app=$null;$p=$null;$show=$null;$rec=$null;$h=[IntPtr]::Zero;$normalH=[IntPtr]::Zero;$ownPid=0
$whole=[Diagnostics.Stopwatch]::StartNew();$events=@();$slides=@();$surfaces=0;$failure=$null;$cleanup=@{};$result=@{status='unverified';passed=$false;combination_verified=$false;inputSha256=$hashes;events=@();slides=@();scope='Own eight-page presentation; real Right keys; no ChartData/workbook calls'}
function Capture([string]$name){$b=New-Object Drawing.Bitmap($width,$height);$g=[Drawing.Graphics]::FromImage($b);try{$g.CopyFromScreen($r.L,$r.T,0,0,$b.Size);$b.Save((Join-Path $Evidence ($name+'.png')),[Drawing.Imaging.ImageFormat]::Png)}finally{$g.Dispose();$b.Dispose()}}
function Guard{
 if($whole.Elapsed.TotalSeconds -gt 53){throw '55-second bounded native budget exceeded'}
 if([PilotWin]::GetForegroundWindow() -ne $h){throw 'Own show lost foreground; stopping without input'}
 $rr=New-Object PilotWin+R;[void][PilotWin]::DwmGetWindowAttribute($h,9,[ref]$rr,16)
 if($rr.L -ne $r.L -or $rr.T -ne $r.T -or $rr.Rt -ne $r.Rt -or $rr.B -ne $r.B){throw 'Own DWM rectangle changed'}
 $cc=New-Object PilotWin+R;[void][PilotWin]::GetClientRect($h,[ref]$cc)
 $originNow=New-Object PilotWin+P;[void][PilotWin]::ClientToScreen($h,[ref]$originNow)
 if($cc.Rt -ne $client.Rt -or $cc.B -ne $client.B -or $originNow.X -ne $origin.X -or $originNow.Y -ne $origin.Y){throw 'Client geometry changed'}
 foreach($fy in @(0.04,0.5,0.96)){foreach($fx in @(0.04,0.5,0.96)){$pt=New-Object PilotWin+P;$pt.X=[int]($origin.X+$ox+$fx*$contentW);$pt.Y=[int]($origin.Y+$oy+$fy*$contentH);if([PilotWin]::GetAncestor([PilotWin]::WindowFromPoint($pt),2) -ne $h){throw 'Foreign popup or surface covers own slide'};$script:surfaces++}}
}
function Read-Shape($s){
 $o=@{name=[string]$s.Name;type=[int]$s.Type;leftPt=[double]$s.Left;topPt=[double]$s.Top;widthPt=[double]$s.Width;heightPt=[double]$s.Height;hasTextFrame=[int]$s.HasTextFrame}
 if($s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1){$tr=$s.TextFrame.TextRange;$t2=$s.TextFrame2;$o.text=[string]$tr.Text;$o.textBoundsPt=@([double]$t2.TextRange.BoundLeft,[double]$t2.TextRange.BoundTop,[double]$t2.TextRange.BoundWidth,[double]$t2.TextRange.BoundHeight);$o.font=@{size=[double]$tr.Font.Size;ea=[string]$tr.Font.NameFarEast;latin=[string]$tr.Font.NameAscii;rgb=[int]$tr.Font.Color.RGB};$o.marginsPt=@([double]$s.TextFrame.MarginLeft,[double]$s.TextFrame.MarginTop,[double]$s.TextFrame.MarginRight,[double]$s.TextFrame.MarginBottom)}
 try{$o.fillRgb=[int]$s.Fill.ForeColor.RGB;$o.fillTheme=[int]$s.Fill.ForeColor.ObjectThemeColor}catch{}
 if($s.Type -eq 6){$kids=@();for($k=1;$k -le $s.GroupItems.Count;$k++){$kids+=Read-Shape $s.GroupItems.Item($k)};$o.children=$kids}
 if($s.HasTable -eq -1){$tb=$s.Table;$rows=@();for($ri=1;$ri -le $tb.Rows.Count;$ri++){$cells=@();for($ci=1;$ci -le $tb.Columns.Count;$ci++){$z=$tb.Cell($ri,$ci).Shape;$zf=$z.TextFrame.TextRange.Font;$zt=$z.TextFrame2.TextRange;$cells+=@{text=[string]$z.TextFrame.TextRange.Text;size=[double]$zf.Size;ea=[string]$zf.NameFarEast;latin=[string]$zf.NameAscii;rgb=[int]$zf.Color.RGB;fillRgb=[int]$z.Fill.ForeColor.RGB;textBoundsPt=@([double]$zt.BoundLeft,[double]$zt.BoundTop,[double]$zt.BoundWidth,[double]$zt.BoundHeight);boundsPt=@([double]$z.Left,[double]$z.Top,[double]$z.Width,[double]$z.Height)}};$rows+=,@($cells)};$o.table=@{rows=[int]$tb.Rows.Count;columns=[int]$tb.Columns.Count;cells=$rows}}
 if($s.HasChart -eq -1){$chart=$s.Chart;$sc=$chart.SeriesCollection();$ser=@();for($si=1;$si -le $sc.Count;$si++){$sv=$sc.Item($si);$ser+=@{name=[string]$sv.Name;values=@($sv.Values);categories=@($sv.XValues)}};$axes=@();foreach($axisId in @(1,2)){try{$ax=$chart.Axes($axisId,1);$axes+=@{type=$axisId;minimum=[double]$ax.MinimumScale;maximum=[double]$ax.MaximumScale}}catch{}};$o.chart=@{series=$ser;axes=$axes;chartType=[int]$chart.ChartType}}
 return $o
}
try{
 $app=New-Object -ComObject PowerPoint.Application;$app.Visible=-1
 foreach($q in $app.Presentations){if([string]$q.FullName -eq $Deck){throw 'Input deck already open; refusing to own/close existing document'}}
 $p=$app.Presentations.Open($Deck,-1,0,-1);if([string]$p.FullName -ne $Deck -or $p.Slides.Count -ne 8){throw 'Own FullName/count mismatch'}
 $leaf=[IO.Path]::GetFileName($Deck);$own=@([Windows.Automation.AutomationElement]::RootElement.FindAll([Windows.Automation.TreeScope]::Children,[Windows.Automation.Condition]::TrueCondition)|Where-Object {$_.Current.ClassName -eq 'PPTFrameClass' -and $_.Current.Name -like ('*'+$leaf+'*')})
 if($own.Count -ne 1){throw 'Own normal window not unique'};$normalH=[IntPtr]$own[0].Current.NativeWindowHandle;$ownPid=[uint32]$own[0].Current.ProcessId
 for($sn=1;$sn -le 8;$sn++){$slide=$p.Slides.Item($sn);$slide.Export((Join-Path $Evidence ('native-'+$sn+'.png')),'PNG',1600,900);$shapes=@();foreach($shape in $slide.Shapes){$shapes+=Read-Shape $shape};$effects=@();foreach($eff in $slide.TimeLine.MainSequence){$effects+=@{shape=[string]$eff.Shape.Name;effect=[int]$eff.EffectType;trigger=[int]$eff.Timing.TriggerType;duration=[double]$eff.Timing.Duration;delay=[double]$eff.Timing.TriggerDelayTime;paragraph=[int]$eff.Paragraph;buildLevel=[int]$eff.EffectInformation.BuildByLevelEffect}};$slides+=@{index=$sn;shapes=$shapes;effects=$effects;transition=@{entry=[int]$slide.SlideShowTransition.EntryEffect;duration=[double]$slide.SlideShowTransition.Duration}}}
 $p.Windows.Item(1).ViewType=9;$p.Slides.Item(1).Select();try{$own[0].SetFocus()}catch{};$shell=New-Object -ComObject WScript.Shell;$activated=$shell.AppActivate($own[0].Current.Name);Start-Sleep -Milliseconds 250
 $p.SlideShowSettings.ShowType=1;$p.SlideShowSettings.ShowPresenterView=0;$p.SlideShowSettings.AdvanceMode=1;$p.SlideShowSettings.StartingSlide=1;$p.SlideShowSettings.EndingSlide=8
 $show=$p.SlideShowSettings.Run();Start-Sleep -Milliseconds 650
 if(-not $show){foreach($sw in $app.SlideShowWindows){if([string]$sw.Presentation.FullName -eq $Deck){$show=$sw}}}
 if(-not $show -or [string]$show.Presentation.FullName -ne $Deck -or [int]$show.View.Slide.SlideIndex -ne 1){throw 'Own start show mismatch'}
 $h=[PilotWin]::Find($ownPid);$screen=@([Windows.Forms.Screen]::AllScreens|Where-Object {-not $_.Primary})[0];if(-not $screen){throw 'Second screen needed for guarded recording'};$wa=$screen.WorkingArea
 [void][PilotWin]::SetWindowPos($h,[IntPtr]::Zero,($wa.Left+40),($wa.Top+40),1280,760,68);Start-Sleep -Milliseconds 250;[void][PilotWin]::SetForegroundWindow($h);Start-Sleep -Milliseconds 200
 $r=New-Object PilotWin+R;if([PilotWin]::DwmGetWindowAttribute($h,9,[ref]$r,16) -ne 0){throw 'DWM geometry unavailable'};$width=$r.Rt-$r.L;$height=$r.B-$r.T
 $client=New-Object PilotWin+R;[void][PilotWin]::GetClientRect($h,[ref]$client);$origin=New-Object PilotWin+P;[void][PilotWin]::ClientToScreen($h,[ref]$origin)
 $scale=[Math]::Min($client.Rt/$design.canvas_cm.width,$client.B/$design.canvas_cm.height);$contentW=$design.canvas_cm.width*$scale;$contentH=$design.canvas_cm.height*$scale;$ox=($client.Rt-$contentW)/2;$oy=($client.B-$contentH)/2
 Guard
 $movie=Join-Path $Evidence 'integration-live.mp4';$args='-hide_banner -loglevel error -f gdigrab -framerate 30 -offset_x '+$r.L+' -offset_y '+$r.T+' -video_size '+$width+'x'+$height+' -i desktop -t 23 -c:v libx264 -preset ultrafast -pix_fmt yuv420p "'+$movie+'"'
 $rec=Start-Process -FilePath $Ffmpeg -ArgumentList $args -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $Evidence 'ffmpeg-error.txt');$pin=$rec.Handle
 $clock=[Diagnostics.Stopwatch]::StartNew();$recordStartUtc=[DateTime]::UtcNow.ToString('o');$schedule=@(@{at=2;slide=2;click=0},@{at=4;slide=3;click=0},@{at=5;slide=3;click=1},@{at=6;slide=3;click=2},@{at=7;slide=3;click=3},@{at=8;slide=3;click=4},@{at=9;slide=4;click=0},@{at=11;slide=5;click=0},@{at=13;slide=6;click=0},@{at=15;slide=7;click=0},@{at=17;slide=8;click=0})
 Capture 'state-s01';$next=0;$pending=@();$transitionDone=@{}
 while($clock.Elapsed.TotalSeconds -lt 23){Guard;$now=$clock.Elapsed.TotalSeconds
  if($next -lt $schedule.Count -and $now -ge $schedule[$next].at){$s=$schedule[$next];$beforeSlide=[int]$show.View.Slide.SlideIndex;$beforeClick=[int]$show.View.GetClickIndex();$sent=$clock.Elapsed.TotalSeconds
   [PilotWin]::keybd_event(0x27,0,0,[UIntPtr]::Zero);try{}finally{[PilotWin]::keybd_event(0x27,0,2,[UIntPtr]::Zero)}
   $pending+=@{at=($sent+0.8);sent=$sent;step=$next;beforeSlide=$beforeSlide;beforeClick=$beforeClick;expectedSlide=$s.slide;expectedClick=$s.click};if($s.slide -eq 6){$transitionStart=$sent};$next++
  }
  if($transitionStart){foreach($offset in @(0.14,0.35,0.55)){if(-not $transitionDone.ContainsKey([string]$offset) -and $now -ge $transitionStart+$offset){Capture ('morph-mid-'+$offset);$transitionDone[[string]$offset]=$clock.Elapsed.TotalSeconds}}}
  foreach($q in @($pending|Where-Object {$now -ge $_.at})){Guard;$actual=[int]$show.View.Slide.SlideIndex;$ci=[int]$show.View.GetClickIndex();$name='state-'+('{0:D2}' -f $q.step)+'-s'+$actual+'-c'+$ci;Capture $name;$ev=@{step=$q.step;input='physical-Right';inputTime=$q.sent;observedTime=$clock.Elapsed.TotalSeconds;beforeSlide=$q.beforeSlide;beforeClick=$q.beforeClick;expectedSlide=$q.expectedSlide;expectedClick=$q.expectedClick;actualSlide=$actual;actualClick=$ci;capture=$name+'.png'};$events+=$ev;if($actual -ne $q.expectedSlide -or ($actual -eq 3 -and $ci -ne $q.expectedClick)){throw ('Physical event mismatch '+($ev|ConvertTo-Json -Compress))};$pending=@($pending|Where-Object {$_ -ne $q})}
  Start-Sleep -Milliseconds 40
 }
 if(-not $rec.WaitForExit(6000)){throw 'Recorder exceeded bounded wait'};$rec.Refresh();if($rec.ExitCode -ne 0 -or -not(Test-Path -LiteralPath $movie) -or (Get-Item -LiteralPath $movie).Length -lt 1000){throw 'Own recorder failed'}
 $result.status='recorded-pending-independent-audit';$result.events=$events;$result.slides=$slides;$result.ownPid=$ownPid;$result.powerPointVersion=[string]$app.Version;$result.normalAppActivate=$activated;$result.window=[int64]$h;$result.dwm=@($r.L,$r.T,$r.Rt,$r.B);$result.client=@($client.Rt,$client.B);$result.clientOrigin=@($origin.X,$origin.Y);$result.contentRectDwmPixels=@(($origin.X-$r.L+$ox),($origin.Y-$r.T+$oy),$contentW,$contentH);$result.surfaceChecks=$surfaces;$result.transitionCaptures=$transitionDone;$result.movie='integration-live.mp4';$result.recorderStartUtc=$recordStartUtc;$result.timingNote='InputTime is stopwatch after StartProcess, not exact first decoded video PTS';$result.passed=$true
}catch{$failure=$_.Exception.Message;$result.status='failed';$result.error=$failure;$result.events=$events;$result.slides=$slides}
finally{
 if($rec){$rec.Refresh();if(-not $rec.HasExited){$rec.Kill();[void]$rec.WaitForExit(2000)}}
 if($show){try{$show.View.Exit();$cleanup.showExit=$true}catch{$cleanup.showExit=$false}}
 if($p){try{if([string]$p.FullName -eq $Deck){$p.Close();$cleanup.deckCloseCalled=$true}}catch{$cleanup.deckCloseCalled=$false}}
 if($app){$remaining=@();try{foreach($doc in $app.Presentations){if([string]$doc.FullName -eq $Deck){$remaining+=[string]$doc.FullName}};$cleanup.ownDeckClosed=($remaining.Count -eq 0);$ownShows=0;foreach($sw in $app.SlideShowWindows){if([string]$sw.Presentation.FullName -eq $Deck){$ownShows++}};$cleanup.ownShowCount=$ownShows}catch{$cleanup.readbackError=$_.Exception.Message};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
 $cleanup.ownerEnabledOrDestroyed=($normalH -eq [IntPtr]::Zero -or -not [PilotWin]::IsWindow($normalH) -or [PilotWin]::IsWindowEnabled($normalH));$cleanup.hashUnchanged=@{};foreach($entry in @{deck=$Deck;plan=$Plan;spec=$Spec}.GetEnumerator()){$cleanup.hashUnchanged[$entry.Key]=((Get-FileHash -LiteralPath $entry.Value -Algorithm SHA256).Hash -eq $hashes[$entry.Key])};$result.cleanup=$cleanup;$result.elapsedSeconds=$whole.Elapsed.TotalSeconds;$result.combination_verified=$false
 $result|ConvertTo-Json -Depth 25|Set-Content -LiteralPath (Join-Path $Evidence 'native.json') -Encoding UTF8
 [void][PilotWin]::SetThreadDpiAwarenessContext($oldDpi)
}
if($failure){throw $failure};if(-not $cleanup.ownDeckClosed -or $cleanup.ownShowCount -ne 0 -or -not $cleanup.ownerEnabledOrDestroyed){throw 'Own cleanup readback failed'}
Write-Output 'Native evidence captured; combination remains unverified until independent audit.'
