param([Parameter(Mandatory=$true)][string]$DeckPath,[Parameter(Mandatory=$true)][string]$OutDir,[switch]$Video)
$ErrorActionPreference='Stop'
$taskFullDeck=(Resolve-Path -LiteralPath $DeckPath).Path
if(Test-Path -LiteralPath $OutDir){throw 'Choose a new output directory; overwrite is disabled'}
[void][IO.Directory]::CreateDirectory($OutDir)
$taskOut=(Resolve-Path -LiteralPath $OutDir).Path
$taskApp=New-Object -ComObject PowerPoint.Application
$taskPres=$null
try {
 $taskPres=$taskApp.Presentations.Open($taskFullDeck,-1,0,0)
 $taskSlides=$taskPres.Slides.Count
 for($taskI=1;$taskI -le $taskSlides;$taskI++) {
  $taskPres.Slides.Item($taskI).Export((Join-Path $taskOut ('slide-{0:D2}.png' -f $taskI)),'PNG',1600,900)
 }
 if($Video) {
  $taskVideo=Join-Path $taskOut 'playback.mp4'
  $taskPres.CreateVideo($taskVideo,$false,3,900,30,85)
  $taskDeadline=[DateTime]::UtcNow.AddMinutes(6)
  do {
   Start-Sleep -Milliseconds 500
   $taskStatus=[int]$taskPres.CreateVideoStatus
   if([DateTime]::UtcNow -gt $taskDeadline){throw 'PowerPoint video export timed out after 6 minutes'}
  } while($taskStatus -in @(1,2))
  if($taskStatus -ne 3){throw "PowerPoint CreateVideoStatus=$taskStatus"}
 }
 [pscustomobject]@{deck=$taskFullDeck;slides=$taskSlides;nativeExport='passed';videoExport=if($Video){'passed'}else{'not_requested'}} | ConvertTo-Json | Out-File -LiteralPath (Join-Path $taskOut 'native-result.json') -Encoding utf8
} finally {
 if($null -ne $taskPres){$taskPres.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($taskPres)}
 [void][Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)
}
