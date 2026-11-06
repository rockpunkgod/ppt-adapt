param(
 [Parameter(Mandatory=$true)][string]$InputPath,
 [Parameter(Mandatory=$true)][string]$OutputPath,
 [Parameter(Mandatory=$true)][string]$PlanPath
)
$ErrorActionPreference='Stop'
$taskInput=[IO.Path]::GetFullPath($InputPath)
$taskOutput=[IO.Path]::GetFullPath($OutputPath)
if(-not(Test-Path -LiteralPath $taskInput)){throw "Missing input: $taskInput"}
if(Test-Path -LiteralPath $taskOutput){throw "Refusing to overwrite: $taskOutput"}
if($taskInput -eq $taskOutput){throw 'Input and output must differ'}
$taskPlan=@([IO.File]::ReadAllText([IO.Path]::GetFullPath($PlanPath))|ConvertFrom-Json)
if($taskPlan.Count -eq 0){throw 'Empty timing plan'}
$taskOutputDir=Split-Path -Parent $taskOutput
if(-not(Test-Path -LiteralPath $taskOutputDir)){New-Item -ItemType Directory -Path $taskOutputDir | Out-Null}
Copy-Item -LiteralPath $taskInput -Destination $taskOutput
$taskApp=$null;$taskPresentation=$null;$taskResults=@()
try {
 try{$taskApp=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')}
 catch{$taskApp=New-Object -ComObject PowerPoint.Application}
 $taskPresentation=$taskApp.Presentations.Open($taskOutput,0,0,0)
 foreach($taskStep in $taskPlan){
  $taskSlide=$taskPresentation.Slides.Item([int]$taskStep.slide)
  $taskMatching=@()
  for($taskIdx=1;$taskIdx -le $taskSlide.TimeLine.MainSequence.Count;$taskIdx++){
   $taskEffect=$taskSlide.TimeLine.MainSequence.Item($taskIdx)
   if($taskEffect.Shape.Name -eq [string]$taskStep.shape){$taskMatching+=,$taskEffect}
  }
  $taskOrdinal=if($null -ne $taskStep.effect){[int]$taskStep.effect}else{1}
  if($taskOrdinal -lt 1 -or $taskOrdinal -gt $taskMatching.Count){throw "Effect not found: slide $($taskStep.slide), shape $($taskStep.shape), ordinal $taskOrdinal"}
  $taskTiming=$taskMatching[$taskOrdinal-1].Timing
  foreach($taskKey in @('duration','delay','accelerate','decelerate')){
   if($null -eq $taskStep.$taskKey){continue}
   $taskValue=[double]$taskStep.$taskKey
   switch($taskKey){
    'duration'{if($taskValue -le 0){throw 'Duration must be positive seconds'};$taskTiming.Duration=$taskValue}
    'delay'{if($taskValue -lt 0){throw 'Delay cannot be negative'};$taskTiming.TriggerDelayTime=$taskValue}
    'accelerate'{if($taskValue -lt 0 -or $taskValue -gt 1){throw 'Accelerate must be 0..1'};$taskTiming.Accelerate=$taskValue}
    'decelerate'{if($taskValue -lt 0 -or $taskValue -gt 1){throw 'Decelerate must be 0..1'};$taskTiming.Decelerate=$taskValue}
   }
  }
  if($null -ne $taskStep.autoReverse){$taskTiming.AutoReverse=if([bool]$taskStep.autoReverse){-1}else{0}}
  if($null -ne $taskStep.repeatCount){
   if([double]$taskStep.repeatCount -lt 1){throw 'repeatCount must be >= 1'}
   $taskTiming.RepeatCount=[double]$taskStep.repeatCount
  }
  if($null -ne $taskStep.trigger){
   $taskTrigger=@{onClick=1;withPrevious=2;afterPrevious=3}[[string]$taskStep.trigger]
   if(-not $taskTrigger){throw "Unknown trigger: $($taskStep.trigger)"}
   $taskTiming.TriggerType=$taskTrigger
  }
  $taskResults+=@{slide=[int]$taskStep.slide;shape=[string]$taskStep.shape;effect=$taskOrdinal;
    duration=$taskTiming.Duration;delay=$taskTiming.TriggerDelayTime;
    accelerate=$taskTiming.Accelerate;decelerate=$taskTiming.Decelerate;
    autoReverse=$taskTiming.AutoReverse;repeatCount=$taskTiming.RepeatCount;trigger=$taskTiming.TriggerType}
 }
 $taskPresentation.Save()
 [IO.File]::WriteAllText($taskOutput+'.timing-readback.json',($taskResults|ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
 Write-Output "Saved timing refinement: $taskOutput"
}
finally{
 if($null -ne $taskPresentation){$taskPresentation.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($taskPresentation)}
 if($null -ne $taskApp){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)}
 # Do not Quit the shared PowerPoint application or close other presentations.
}
