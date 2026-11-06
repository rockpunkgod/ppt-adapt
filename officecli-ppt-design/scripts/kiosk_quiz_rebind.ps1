param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory,[Parameter(Mandatory=$true)][ValidateSet(4,5,6)][int]$CorrectAnswer)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $EvidenceDirectory)){throw 'Fresh outputs required'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
function Rgb([string]$s){return [Convert]::ToInt32($s.Substring(0,2),16)+([Convert]::ToInt32($s.Substring(2,2),16) -shl 8)+([Convert]::ToInt32($s.Substring(4,2),16) -shl 16)}
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$nextEffect=$null;$targetFont=$null;$colors=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$s=$p.Slides.Item(1);$s.Select()
 $q=$s.Shapes.Item('question').TextFrame.TextRange.Text
 if($q -notmatch '(\d+)\s*\+\s*(\d+)' -or ([int]$Matches[1]+[int]$Matches[2]) -ne $CorrectAnswer){throw 'Question and correct answer mismatch'}
 if($s.TimeLine.InteractiveSequences.Count -ne 3){throw 'Expected the tested three-trigger quiz'}
 for($i=1;$i -le $s.TimeLine.InteractiveSequences.Count;$i++){
  $seq=$s.TimeLine.InteractiveSequences.Item($i)
  for($j=1;$j -le $seq.Count;$j++){
   $e=$seq.Item($j);$name=$e.Shape.Name
   if($name -eq 'next'){$nextEffect=$e;continue}
   if($name -notmatch '^answer-(4|5|6)$'){throw 'Unexpected animated object'}
   $answer=[int]$Matches[1]
   if([int]$e.EffectType -eq 54){$e.EffectParameters.Color2.RGB=(Rgb $(if($answer -eq $CorrectAnswer){'21645E'}else{'AF4E45'}));$colors+=@{answer=$answer;isCorrect=($answer -eq $CorrectAnswer)}}
   if([int]$e.EffectType -eq 56 -and $answer -eq $CorrectAnswer){$targetFont=$e}
  }
 }
 if(-not $nextEffect -or -not $targetFont){throw 'Next or destination font effect missing'}
 $before=$nextEffect.Parent.Item(1).Timing.TriggerShape.Name
 $nextEffect.MoveAfter($targetFont);$nextEffect.Timing.TriggerType=3
 $cachedAfter=$nextEffect.Parent.Item(1).Timing.TriggerShape.Name
 # PowerPoint can retain an Effect.Parent RCW after a cross-sequence MoveAfter.
 # Re-enumerate actual sequence contents instead of trusting that stale parent.
 $after=$null;$nextCount=0
 for($i=1;$i -le $s.TimeLine.InteractiveSequences.Count;$i++){$sq=$s.TimeLine.InteractiveSequences.Item($i);for($j=1;$j -le $sq.Count;$j++){if($sq.Item($j).Shape.Name -eq 'next'){$nextCount++;$after=$sq.Item(1).Timing.TriggerShape.Name}}}
 if($nextCount -ne 1){throw 'Expected exactly one existing Next effect after move'}
 if($after -ne ('answer-'+$CorrectAnswer)){throw 'Next failed to move to new correct trigger'}
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'migration.json'),(@{question=$q;correct=$CorrectAnswer;beforeTrigger=$before;afterTrigger=$after;cachedParentAfter=$cachedAfter;colors=$colors;showType=[int]$p.SlideShowSettings.ShowType}|ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
