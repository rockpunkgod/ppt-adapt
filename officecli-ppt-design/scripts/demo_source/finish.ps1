param([string]$Work,[string]$OutputDeck)
$ErrorActionPreference='Stop'
$OutputDeck=[IO.Path]::GetFullPath($OutputDeck)
$app=New-Object -ComObject PowerPoint.Application
$p=$null
try {
 $p=$app.Presentations.Open((Join-Path $Work 'main.pptx'),0,0,0)
 foreach($s in $p.Slides){
  $s.SlideShowTransition.AdvanceOnClick=-1
  $s.SlideShowTransition.AdvanceOnTime=0
 }
 $p.SaveAs($OutputDeck,24)
} finally {
 if($p){$p.Saved=-1;$p.Close()}
 # Preserve the shared PowerPoint application and all user documents.
}
