param([string]$Work,[string]$OutputDeck)
$ErrorActionPreference='Stop'
$app=New-Object -ComObject PowerPoint.Application
$p=$null
try {
 $p=$app.Presentations.Open((Join-Path $Work 'main.pptx'),0,0,0)
 foreach($s in $p.Slides){
  foreach($v in $s.Shapes){
   if($v.Name.TrimStart('!') -eq 'paper-facet-3' -or $v.Name.TrimStart('!') -eq 'knowledge-doc0'){
    $v.Shadow.Visible=-1;$v.Shadow.ForeColor.RGB=0
    $v.Shadow.Transparency=[single].83;$v.Shadow.Blur=[single]18
    $v.Shadow.OffsetX=[single]0;$v.Shadow.OffsetY=[single]8
   }
  }
  $s.SlideShowTransition.AdvanceOnClick=-1
  $s.SlideShowTransition.AdvanceOnTime=0
 }
 $p.SaveAs($OutputDeck,24)
} finally {
 if($p){$p.Saved=-1;$p.Close()}
 # Preserve the shared PowerPoint application and all user documents.
}
