param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $OutputDirectory){throw ('Refusing to overwrite '+$OutputDirectory)}
$taskInput=[IO.Path]::GetFullPath($InputDeck)
if(-not(Test-Path -LiteralPath $taskInput)){throw ('Input missing: '+$taskInput)}
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$taskApp=New-Object -ComObject PowerPoint.Application
$taskP=$null;$taskReport=@()
try{
 $taskP=$taskApp.Presentations.Open($taskInput,$true,$false,$false)
 for($taskI=1;$taskI -le $taskP.Slides.Count;$taskI++){
  $taskS=$taskP.Slides.Item($taskI)
  $taskS.Export((Join-Path $OutputDirectory ('slide-'+$taskI+'.png')),'PNG',1600,900)
  $taskOverflow=@()
  foreach($taskSh in $taskS.Shapes){
   if($taskSh.HasTextFrame -and $taskSh.TextFrame.HasText){if($taskSh.TextFrame2.TextRange.BoundHeight -gt $taskSh.Height+2){$taskOverflow+=@{name=$taskSh.Name;bound=$taskSh.TextFrame2.TextRange.BoundHeight;height=$taskSh.Height}}}
  }
  $taskReport+=@{slide=$taskI;shapes=$taskS.Shapes.Count;overflow=$taskOverflow}
 }
}finally{
 if($taskP){$taskP.Close()}
 $taskReport|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $OutputDirectory 'native-report.json') -Encoding utf8
}
# Intentionally do not Quit the shared PowerPoint application.
