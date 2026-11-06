param([Parameter(Mandatory=$true)][string]$Root)
$ErrorActionPreference='Stop'
$taskOutput=Join-Path $Root 'refinement-com-v3.pptx'
if(Test-Path -LiteralPath $taskOutput){throw ('Refusing to overwrite '+$taskOutput)}
$taskPpt=New-Object -ComObject PowerPoint.Application
$taskPresentation=$taskPpt.Presentations.Open((Join-Path $Root 'refinement-base-v3.pptx'),$false,$false,$false)
$taskReport=@()
try{
  $taskSlide=$taskPresentation.Slides.Item(2)
  for($taskI=1;$taskI -le $taskSlide.TimeLine.MainSequence.Count;$taskI++){
    $taskEffect=$taskSlide.TimeLine.MainSequence.Item($taskI)
    $taskEffect.Timing.TriggerType=2
    $taskEffect.Timing.Duration=2
    $taskEffect.Timing.TriggerDelayTime=($taskI-1)*0.1
    $taskEffect.Timing.SmoothStart=0
    $taskEffect.Timing.SmoothEnd=0
    $taskEffect.Timing.BounceEnd=0
    $taskEffect.Timing.BounceEndIntensity=0
    $taskReport += @{test='fly-delay-only';shape=$taskEffect.Shape.Name;bounce=$taskEffect.Timing.BounceEnd;intensity=$taskEffect.Timing.BounceEndIntensity;delay=$taskEffect.Timing.TriggerDelayTime;duration=$taskEffect.Timing.Duration}
  }
  $taskWaterSlide=$taskPresentation.Slides.Item(5)
  $taskWaterSlide.TimeLine.MainSequence.Item(1).Delete()
  $taskWater=$taskWaterSlide.TimeLine.MainSequence.AddEffect($taskWaterSlide.Shapes.Item('water'),2,0,2)
  $taskWater.EffectParameters.Direction=11
  $taskWater.Timing.TriggerType=2
  $taskWater.Timing.TriggerDelayTime=0
  $taskWater.Timing.Duration=2
  $taskWater.Timing.SmoothStart=0
  $taskWater.Timing.SmoothEnd=0
  $taskReport += @{test='mask-motion';trigger=[int]$taskWater.Timing.TriggerType;direction=[int]$taskWater.EffectParameters.Direction}
  $taskSlide=$taskPresentation.Slides.Item(6)
  $taskPin=$taskSlide.Shapes.Item('trigger-pin')
  $taskLabel=$taskSlide.Shapes.Item('trigger-label')
  $taskSeq=$taskSlide.TimeLine.InteractiveSequences.Add(1)
  $taskEntrance=$taskSeq.AddTriggerEffect($taskLabel,10,4,$taskPin,'',0)
  $taskEntrance.Timing.Duration=0.5
  $taskActualSeq=$taskEntrance.Parent
  $taskReport += @{test='named-trigger';sequences=$taskSlide.TimeLine.InteractiveSequences.Count;effects=$taskActualSeq.Count;triggerShape=$taskEntrance.Timing.TriggerShape.Name}
  $taskPresentation.SaveAs($taskOutput)
  $taskNative=Join-Path $Root 'refinement-com-v3-native'
  New-Item -ItemType Directory -Path $taskNative -Force | Out-Null
  for($taskI=1;$taskI -le $taskPresentation.Slides.Count;$taskI++){$taskPresentation.Slides.Item($taskI).Export((Join-Path $taskNative ('slide-'+$taskI+'.png')),'PNG',1600,900)}
  $taskVideo=Join-Path $Root 'refinement-com-v3-native.mp4'
  $taskPresentation.CreateVideo($taskVideo,$false,7,720,24,80)
  $taskDeadline=(Get-Date).AddMinutes(2)
  while(($taskPresentation.CreateVideoStatus -eq 1 -or $taskPresentation.CreateVideoStatus -eq 2) -and (Get-Date) -lt $taskDeadline){Start-Sleep -Milliseconds 400}
  $taskReport += @{test='native-export';status=[int]$taskPresentation.CreateVideoStatus}
}finally{$taskPresentation.Close()}
$taskReport | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $Root 'refinement-com-v3-report.json') -Encoding utf8
$taskReport | ConvertTo-Json -Depth 7
