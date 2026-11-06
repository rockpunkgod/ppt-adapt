param([string]$Deck,[string]$Output,[string]$Landscape,[string]$Portrait,[string]$RenderDirectory)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Output){throw 'Refusing overwrite of polished presentation'}
if(Test-Path -LiteralPath $RenderDirectory){throw 'Choose a new render directory'}
New-Item -ItemType Directory -Path $RenderDirectory|Out-Null
function Color-RGB([string]$hex){$r=[Convert]::ToInt32($hex.Substring(0,2),16);$g=[Convert]::ToInt32($hex.Substring(2,2),16);$b=[Convert]::ToInt32($hex.Substring(4,2),16);return $r -bor ($g -shl 8) -bor ($b -shl 16)}
$taskApp=New-Object -ComObject PowerPoint.Application;$taskP=$null
try{
 $taskP=$taskApp.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,0)
 foreach($s in $taskP.Slides.Item(3).Shapes){
  if($s.Name -like 'outline-*'){
   $s.TextFrame2.TextRange.Font.Fill.Visible=0
   $s.TextFrame2.TextRange.Font.Line.Visible=-1
   $s.TextFrame2.TextRange.Font.Line.ForeColor.RGB=Color-RGB '47646B'
   $s.TextFrame2.TextRange.Font.Line.Weight=.7
  }
 }
 $table=$taskP.Slides.Item(7).Shapes.Item('editorial-grid').Table
 for($i=1;$i -le 5;$i++){
  $fill=$table.Cell(1,$i).Shape.Fill;$fill.Solid();$fill.ForeColor.RGB=Color-RGB '15212B'
  $fill.Transparency=if($i -in 2,4){.55}else{.45}
  $s=$taskP.Slides.Item(7).Shapes.Item('ed-num-'+($i-1))
  $s.TextFrame2.TextRange.Font.Fill.Visible=0
  $s.TextFrame2.TextRange.Font.Line.Visible=-1
  $s.TextFrame2.TextRange.Font.Line.ForeColor.RGB=Color-RGB 'FFFFFF'
  $s.TextFrame2.TextRange.Font.Line.Weight=.8
  $text=$taskP.Slides.Item(7).Shapes.Item('ed-text-'+($i-1)).TextFrame2.TextRange.Paragraphs(1,1)
  $text.Font.Size=28;$text.Font.Bold=-1
 }
 # OfficeCLI 1.0.160 labels preset 12 as float, but native PIA defines 12=Peek, 30=Float.
 foreach($effect in $taskP.Slides.Item(7).TimeLine.MainSequence){
  $name=$effect.Shape.Name
  if($name -like 'ed-num-*'){
   $effect.EffectType=30;$effect.Timing.Duration=.4;$effect.Timing.TriggerDelayTime=0
   $effect.Timing.TriggerType=if($name -eq 'ed-num-0'){2}else{3}
  }elseif($name -like 'ed-text-*'){
   $effect.Timing.Duration=.4;$effect.Timing.TriggerDelayTime=0;$effect.Timing.TriggerType=3
  }
 }
 $assets=@([IO.Path]::GetFullPath($Landscape),[IO.Path]::GetFullPath($Portrait))
 for($i=0;$i -lt 4;$i++){
  $s=$taskP.Slides.Item(13).Shapes.Item('glyph-'+$i)
  $s.TextFrame2.TextRange.Font.Fill.UserPicture($assets[$i%2])
 }
 $taskP.SaveAs([IO.Path]::GetFullPath($Output))
 for($i=1;$i -le 13;$i++){
  $taskP.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($RenderDirectory)) ('slide-'+$i+'.png')),'PNG',1600,900)
  $t=$taskP.Slides.Item($i).SlideShowTransition;$t.AdvanceOnTime=-1;$t.AdvanceTime=2
  if($i -eq 7){$t.AdvanceTime=6}
 }
 $taskP.CreateVideo((Join-Path ([IO.Path]::GetFullPath($RenderDirectory)) 'native.mp4'),-1,2,720,24,80)
 $deadline=[DateTime]::UtcNow.AddMinutes(3)
 do{Start-Sleep -Milliseconds 500;$status=$taskP.CreateVideoStatus;if([DateTime]::UtcNow -gt $deadline){throw 'Movie export deadline'}}while($status -in 1,2)
 if($status -ne 3){throw "CreateVideoStatus=$status"}
 Write-Output 'Native text outlines, native glyph pictures, cell opacity and all 13 pages exported.'
}finally{if($taskP){$taskP.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($taskP)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)|Out-Null}
