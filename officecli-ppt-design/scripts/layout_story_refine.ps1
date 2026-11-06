param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$PatternImage,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
$taskOut=[IO.Path]::GetFullPath($OutputDeck)
if(Test-Path -LiteralPath $EvidenceDirectory){throw ('Refusing to overwrite evidence directory '+$EvidenceDirectory)}
if(Test-Path -LiteralPath $taskOut){throw ('Refusing to overwrite '+$taskOut)}
[void][IO.Directory]::CreateDirectory($EvidenceDirectory)
$taskApp=New-Object -ComObject PowerPoint.Application
$taskP=$taskApp.Presentations.Open(([IO.Path]::GetFullPath($InputDeck)),$false,$false,$true)
$taskLog=@()
function Merge-Names($taskSlide,[object[]]$taskNames,[int]$taskMode,$taskPrimary){
 $taskSlide.Select()
 $taskOld=@($taskSlide.Shapes|ForEach-Object{$_.Id})
 $taskRange=$taskSlide.Shapes.Range($taskNames)
 $taskRange.MergeShapes($taskMode,$taskPrimary)
 return @($taskSlide.Shapes|Where-Object{$taskOld -notcontains $_.Id})
}
try{
 $taskS=$taskP.Slides.Item(1)
 $taskPrimary=$taskS.Shapes.Item('ball-left')
 $taskMerged=@(Merge-Names $taskS @('ball-left','ball-right','ball-top','cone') 1 $taskPrimary)
 if($taskMerged.Count -ne 1){throw 'Union did not yield one new shape'}
 $taskShape=$taskMerged[0];$taskShape.Name='custom-union'
 $taskPic=$taskS.Shapes.Item('source-image')
 $taskResult=@(Merge-Names $taskS @('source-image','custom-union') 3 $taskPic)
 $taskLog+=@{operation='Union then picture Intersect';slide=1;outputs=$taskResult.Count;shapeTypes=@($taskResult|ForEach-Object{[int]$_.Type})}
 foreach($taskRow in @(@(2,3,'text-picture-intersect'),@(3,4,'text-picture-subtract'),@(4,5,'text-picture-fragment'))){
  $taskS=$taskP.Slides.Item($taskRow[0]);$taskPic=$taskS.Shapes.Item('source-image')
  $taskResult=@(Merge-Names $taskS @('source-image','mask-text') $taskRow[1] $taskPic)
  $taskLog+=@{operation=$taskRow[2];slide=$taskRow[0];outputs=$taskResult.Count;shapeTypes=@($taskResult|ForEach-Object{[int]$_.Type})}
 }
 $taskS=$taskP.Slides.Item(5)
 $taskLayout=$taskP.SlideMaster.CustomLayouts.Add($taskP.SlideMaster.CustomLayouts.Count+1)
 $taskLayout.Name='Reusable picture frame'
 $taskPh=$taskLayout.Shapes.AddPlaceholder(18,430,155,440,290)
 $taskPh.Name='layout-picture-frame'
 $taskS.CustomLayout=$taskLayout
 $taskSlidePh=@($taskS.Shapes|Where-Object{$_.Type -eq 14 -and $_.PlaceholderFormat.Type -eq 18})[0]
 if(-not $taskSlidePh){$taskSlidePh=$taskS.Shapes.AddPlaceholder(18,430,155,440,290)}
 $taskSlidePh.Name='editable-picture-placeholder'
 $taskSlidePh.Fill.UserPicture(([IO.Path]::GetFullPath($PatternImage)))
 $taskLog+=@{operation='native picture placeholder';slide=5;placeholderType=[int]$taskSlidePh.PlaceholderFormat.Type;layout=$taskS.CustomLayout.Name;placeholderCount=$taskS.Shapes.Placeholders.Count}
 $taskS=$taskP.Slides.Item(6)
 for($taskI=0;$taskI -lt 4;$taskI++){
  $taskSh=$taskS.Shapes.Item('photo-'+$taskI)
  $taskEffect=$taskS.TimeLine.MainSequence.AddEffect($taskSh,33,0,$(if($taskI -eq 0){2}else{3}))
  $taskEffect.Timing.Duration=1
  $taskLog+=@{operation='Pinwheel';shape=$taskSh.Name;effect=[int]$taskEffect.EffectType;trigger=[int]$taskEffect.Timing.TriggerType;duration=$taskEffect.Timing.Duration}
 }
 $taskP.SaveAs($taskOut)
 $taskNative=Join-Path $EvidenceDirectory 'native';New-Item -ItemType Directory -Path $taskNative -Force|Out-Null
 for($taskI=1;$taskI -le $taskP.Slides.Count;$taskI++){$taskP.Slides.Item($taskI).Export((Join-Path $taskNative ('slide-'+$taskI+'.png')),'PNG',1600,900)}
 # Export only motion slide to keep a bounded real native animation check.
 $taskCopy=Join-Path $EvidenceDirectory 'pinwheel-native.pptx';if(Test-Path -LiteralPath $taskCopy){throw 'Refusing to overwrite video deck'}
 $taskP.SaveCopyAs($taskCopy)
 $taskMovieP=$taskApp.Presentations.Open($taskCopy,$false,$false,$false)
 try{
  for($taskI=$taskMovieP.Slides.Count;$taskI -ge 1;$taskI--){if($taskI -ne 6){$taskMovieP.Slides.Item($taskI).Delete()}}
  $taskMovieP.Save()
  $taskMovieP.CreateVideo((Join-Path $EvidenceDirectory 'pinwheel-native.mp4'),$false,7,720,24,80)
  $taskDeadline=(Get-Date).AddMinutes(1)
  while(($taskMovieP.CreateVideoStatus -eq 1 -or $taskMovieP.CreateVideoStatus -eq 2) -and (Get-Date) -lt $taskDeadline){Start-Sleep -Milliseconds 300}
  $taskLog+=@{operation='native Pinwheel video';status=[int]$taskMovieP.CreateVideoStatus}
 }finally{$taskMovieP.Close()}
}finally{$taskP.Close();$taskLog|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $EvidenceDirectory 'com-report.json') -Encoding utf8}
$taskLog|ConvertTo-Json -Depth 8
# Close own presentations only; never Quit shared PowerPoint.
