param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $OutputDeck){throw 'Fresh output required'}
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,0)
 if($p.Slides.Count -ne 5){throw 'Expected existing five-page Section Zoom foundation'}
 foreach($idx in @(1,4,7)){$copy=$p.Slides.Item($idx).Duplicate();$copy.Item(1).MoveTo($idx+1)}
 if($p.Slides.Count -ne 8){throw 'Page plan mismatch'}
 foreach($i in @(1,2)){
  $s=$p.Slides.Item($i)
  $s.Shapes.Item('title').TextFrame.TextRange.Text='FIELD / STUDIO'
  $s.Shapes.Item('subtitle').TextFrame.TextRange.Text='Explore the evidence. Build a shared language.'
  $s.Shapes.Item('hub-footer').TextFrame.TextRange.Text='01  FIELD NOTES                       02  DESIGN SYSTEM'
  $s.SlideShowTransition.AdvanceOnTime=0;$s.SlideShowTransition.AdvanceOnClick=0
 }
 foreach($i in @(4,5,7,8)){$s=$p.Slides.Item($i);$s.SlideShowTransition.AdvanceOnTime=0;$s.SlideShowTransition.AdvanceOnClick=-1}
 $p.Slides.Item(4).Shapes.Item('!!body').TextFrame.TextRange.Text="Capture one observation.`rSeparate evidence from assumptions.`rName the decision it informs."
 $p.Slides.Item(5).Shapes.Item('!!chapter-title').TextFrame.TextRange.Text='FROM NOTES TO ACTION'
 $p.Slides.Item(5).Shapes.Item('!!body').TextFrame.TextRange.Text="Observe  >  Compare  >  Decide`rKeep the original evidence close.`rMake the next action explicit."
 $p.Slides.Item(7).Shapes.Item('!!body').TextFrame.TextRange.Text="Set a clear type hierarchy.`rReuse meaningful color roles.`rLet each page serve its message."
 $p.Slides.Item(8).Shapes.Item('!!chapter-title').TextFrame.TextRange.Text='CONSISTENCY WITH VARIETY'
 $p.Slides.Item(8).Shapes.Item('!!body').TextFrame.TextRange.Text="A shared visual language.`rDifferent compositions for different ideas.`rMotion only where it helps understanding."
 foreach($i in @(4,7)){$p.Slides.Item($i).Shapes.Item('!!section-end').TextFrame.TextRange.Text='Continue to the second page of this chapter'}
 foreach($i in @(5,8)){$p.Slides.Item($i).Shapes.Item('!!section-end').TextFrame.TextRange.Text='Chapter complete / click to return to the studio'}
 $intro=$p.Slides.Item(1)
 foreach($name in @('section-card-1','section-card-2')){$e=$intro.TimeLine.MainSequence.AddEffect($intro.Shapes.Item($name),23,0,2);$e.Timing.Duration=1}
 $intro.SlideShowTransition.AdvanceOnTime=-1;$intro.SlideShowTransition.AdvanceTime=0;$intro.SlideShowTransition.AdvanceOnClick=-1
 $p.SlideShowSettings.RangeType=1;$p.SlideShowSettings.ShowWithAnimation=-1;$p.SlideShowSettings.AdvanceMode=2
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 $rows=@();for($i=1;$i -le $p.SectionProperties.Count;$i++){$rows+=@{name=$p.SectionProperties.Name($i);first=$p.SectionProperties.FirstSlide($i);count=$p.SectionProperties.SlidesCount($i);id=$p.SectionProperties.SectionID($i)}}
 @{slides=$p.Slides.Count;sections=$rows;introEffects=$intro.TimeLine.MainSequence.Count;staticHomeEffects=$p.Slides.Item(2).TimeLine.MainSequence.Count;autoTime=$intro.SlideShowTransition.AdvanceTime;autoOn=[int]$intro.SlideShowTransition.AdvanceOnTime}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath ($OutputDeck+'.json') -Encoding UTF8
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
