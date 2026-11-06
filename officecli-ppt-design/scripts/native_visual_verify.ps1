param([Parameter(Mandatory=$true)][string]$Root)
$ErrorActionPreference='Stop'
$taskRoot=[IO.Path]::GetFullPath($Root)
$taskPpt=New-Object -ComObject PowerPoint.Application
$taskReport=@()
foreach ($taskName in @('zoom-lab','visual-lab','animation-lab')) {
  $taskPath=Join-Path $taskRoot ($taskName+'.pptx')
  $taskPresentation=$taskPpt.Presentations.Open($taskPath,$true,$false,$false)
  try {
    $taskDir=Join-Path $taskRoot ($taskName+'-native')
    New-Item -ItemType Directory -Path $taskDir -Force | Out-Null
    for ($taskI=1;$taskI -le $taskPresentation.Slides.Count;$taskI++) {
      $taskPresentation.Slides.Item($taskI).Export((Join-Path $taskDir ('slide-'+$taskI+'.png')),'PNG',1600,900)
      $taskEffects=@()
      $taskSlide=$taskPresentation.Slides.Item($taskI)
      for($taskE=1;$taskE -le $taskSlide.TimeLine.MainSequence.Count;$taskE++){
        $taskEffect=$taskSlide.TimeLine.MainSequence.Item($taskE)
        $taskEffects += @{shape=$taskEffect.Shape.Name;effect=[int]$taskEffect.EffectType;duration=$taskEffect.Timing.Duration;delay=$taskEffect.Timing.TriggerDelayTime;repeat=$taskEffect.Timing.RepeatCount;autoReverse=[int]$taskEffect.Timing.AutoReverse}
      }
      $taskReport += @{deck=$taskName;slide=$taskI;effects=$taskEffects;shapes=$taskSlide.Shapes.Count}
    }
    if($taskName -ne 'zoom-lab'){
      $taskVideo=Join-Path $taskRoot ($taskName+'-native.mp4')
      $taskPresentation.CreateVideo($taskVideo,$false,7,720,24,80)
      $taskDeadline=(Get-Date).AddMinutes(3)
      while(($taskPresentation.CreateVideoStatus -eq 1 -or $taskPresentation.CreateVideoStatus -eq 2) -and (Get-Date) -lt $taskDeadline){Start-Sleep -Milliseconds 500}
      $taskReport += @{deck=$taskName;video=$taskVideo;status=[int]$taskPresentation.CreateVideoStatus}
      Write-Output ($taskName+' video status '+$taskPresentation.CreateVideoStatus)
    }
  } finally {$taskPresentation.Close()}
}
$taskReport | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $taskRoot 'native-report.json') -Encoding utf8
# Do not Quit: other agents may share PowerPoint.Application.
