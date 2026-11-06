param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory,[Parameter(Mandatory=$true)][ValidateSet(4,5,6)][int]$CorrectAnswer)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $EvidenceDirectory)){throw 'Fresh outputs required'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
function Rgb([string]$s){return [Convert]::ToInt32($s.Substring(0,2),16)+([Convert]::ToInt32($s.Substring(2,2),16) -shl 8)+([Convert]::ToInt32($s.Substring(4,2),16) -shl 16)}
function Hold($effect){$effect.Timing.RewindAtEnd=0;for($i=1;$i -le $effect.Behaviors.Count;$i++){$effect.Behaviors.Item($i).Timing.RewindAtEnd=0}}
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$rows=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$s=$p.Slides.Item(1);$s.Select()
 if($p.Slides.Count -ne 2 -or $s.TimeLine.MainSequence.Count -ne 0 -or $s.TimeLine.InteractiveSequences.Count -ne 0){throw 'Requires fresh two-slide CLI quiz baseline'}
 $question=$s.Shapes.Item('question').TextFrame.TextRange.Text
 if($question -notmatch '(\d+)\s*\+\s*(\d+)' -or ([int]$Matches[1]+[int]$Matches[2]) -ne $CorrectAnswer){throw 'CorrectAnswer does not match actual arithmetic question'}
 $next=$s.Shapes.Item('next')
 foreach($answer in @(4,5,6)){
  $button=$s.Shapes.Item('answer-'+$answer);$seq=$s.TimeLine.InteractiveSequences.Add($s.TimeLine.InteractiveSequences.Count+1)
  $fill=$seq.AddTriggerEffect($button,54,4,$button,'',0);$fill.Timing.Duration=[single]0.25;$fill.EffectParameters.Color2.RGB=(Rgb $(if($answer -eq $CorrectAnswer){'21645E'}else{'AF4E45'}));Hold $fill
  $font=$s.TimeLine.MainSequence.AddEffect($button,56,0,2);$font.Timing.Duration=[single]0.25;$font.EffectParameters.Color2.RGB=16777215;$font.MoveAfter($fill);$font.Timing.TriggerType=2;Hold $font
  if($answer -eq $CorrectAnswer){
   $entry=$s.TimeLine.MainSequence.AddEffect($next,2,0,3);$entry.EffectParameters.Direction=11;$entry.Timing.Duration=[single]0.5;$entry.MoveAfter($font);$entry.Timing.TriggerType=3;$entry.Timing.TriggerDelayTime=[single]0;Hold $entry
  }
  $actual=$fill.Parent;$rows+=@{answer=$answer;correct=($answer -eq $CorrectAnswer);count=$actual.Count;triggerShape=$actual.Item(1).Timing.TriggerShape.Name;fillDuration=[single]$fill.Timing.Duration;fontDuration=[single]$font.Timing.Duration}
 }
 $p.SlideShowSettings.ShowType=3;$p.SlideShowSettings.RangeType=1;$p.SlideShowSettings.AdvanceMode=1;$p.SlideShowSettings.ShowPresenterView=0
 foreach($i in @(1,2)){$p.Slides.Item($i).SlideShowTransition.AdvanceOnClick=-1;$p.Slides.Item($i).SlideShowTransition.AdvanceOnTime=0}
 $nextAction=$next.ActionSettings.Item(1)
 $rows+=@{mainSequence=$s.TimeLine.MainSequence.Count;interactiveSequences=$s.TimeLine.InteractiveSequences.Count;showType=[int]$p.SlideShowSettings.ShowType;nextAction=[int]$nextAction.Action;nextSubAddress=$nextAction.Hyperlink.SubAddress;nextExpectedId=[int]$p.Slides.Item(2).SlideID}
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 foreach($i in @(1,2)){$p.Slides.Item($i).Export((Join-Path $EvidenceDirectory ('edit-'+$i+'.png')),'PNG',1600,900)}
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'readback.json'),($rows|ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
