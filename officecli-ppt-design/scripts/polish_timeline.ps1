param([string]$Deck,[string]$Output)
$ErrorActionPreference='Stop'
if(-not(Test-Path -LiteralPath ([IO.Path]::GetFullPath($Deck)))){throw 'Missing input deck'}
if(Test-Path -LiteralPath ([IO.Path]::GetFullPath($Output))){throw 'Refusing to overwrite output'}
if([IO.Path]::GetFullPath($Deck) -eq [IO.Path]::GetFullPath($Output)){throw 'Input and output must differ'}
$taskApp=New-Object -ComObject PowerPoint.Application
$taskP=$null
try {
 $taskP=$taskApp.Presentations.Open([IO.Path]::GetFullPath($Deck),$false,$false,$false)
 $taskSequence=$taskP.Slides.Item(5).TimeLine.MainSequence
 if($taskSequence.Count -ne 7){throw "Expected 7 motion effects; got $($taskSequence.Count)"}
 for($taskI=1;$taskI -le 7;$taskI++){
   $taskEffect=$taskSequence.Item($taskI)
   $taskEffect.Timing.AutoReverse=-1
   $taskEffect.Timing.RepeatCount=3
   $taskEffect.Timing.TriggerType=2
   $taskJ=[int]($taskEffect.Shape.Name -replace '^trail-','')
   $taskEffect.Timing.Duration=3
   $taskEffect.Timing.TriggerDelayTime=$taskJ*0.07
 }
 $taskP.SaveAs([IO.Path]::GetFullPath($Output))
 Write-Output 'Polished 7 native motion timings with AutoReverse and RepeatCount=3.'
} finally {
 if($taskP){$taskP.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($taskP)|Out-Null}
 [Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)|Out-Null
}
