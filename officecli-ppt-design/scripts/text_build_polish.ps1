param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Evidence)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'Fresh outputs required'}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($Evidence))
$app=New-Object -ComObject PowerPoint.Application;$pres=$null;$log=@()
try{
 $app.Visible=-1;$pres=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 if($pres.Slides.Count -ne 3){throw 'Expected the fresh three-slide text-build foundation'}
 for($i=1;$i -le 3;$i++){
  $s=$pres.Slides.Item($i);$text=$s.Shapes.Item('text-build-target');$seq=$s.TimeLine.MainSequence
  if($seq.Count -ne 0){throw 'Expected no pre-existing animations'}
  if($i -eq 1){
   $effect=$seq.AddEffect($text,10,1,1)
   for($j=1;$j -le $seq.Count;$j++){$item=$seq.Item($j);$item.Timing.TriggerType=1;$item.Timing.Duration=[single]0.5}
  }
  else{
   $effect=$seq.AddEffect($text,19,0,1)
   $unit=if($i -eq 2){1}else{2}
   $effect=$seq.ConvertToTextUnitEffect($effect,$unit)
   $effect.Timing.Duration=[single]0.2
  }
  $s.SlideShowTransition.AdvanceOnClick=-1;$s.SlideShowTransition.AdvanceOnTime=0
  $log+=@{slide=$i;paragraphs=$text.TextFrame.TextRange.Paragraphs().Count;effects=$seq.Count;effectType=[int]$effect.EffectType;unit=[int]$effect.EffectInformation.TextUnitEffect;buildLevel=[int]$effect.EffectInformation.BuildByLevelEffect;duration=[double]$effect.Timing.Duration;text=$text.TextFrame.TextRange.Text}
  $s.Export((Join-Path $Evidence "slide-$i.png"),'PNG',1600,900)
 }
 $pres.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 $log|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $Evidence 'parameters.json') -Encoding utf8
}finally{if($pres){$pres.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
