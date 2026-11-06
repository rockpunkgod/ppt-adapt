param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Sections,[Parameter(Mandatory=$true)][string]$Posters)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Sections) -or (Test-Path -LiteralPath $Posters)){throw 'Fresh outputs required'}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($Posters))
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,0)
 if($p.Slides.Count -ne 5){throw 'Expected existing five-page Hover Zoom foundation'}
 foreach($idx in @(4,6)){$copy=$p.Slides.Item($idx).Duplicate();$copy.Item(1).MoveTo($idx+1)}
 if($p.Slides.Count -ne 7){throw 'Page plan mismatch'}
 foreach($i in @(1,2,3)){$p.Slides.Item($i).Shapes.Item('!!hint').TextFrame.TextRange.Text='Hover to preview. Explore both pages, then return.';$p.Slides.Item($i).SlideShowTransition.Duration=[single]0.35}
 $p.Slides.Item(5).Shapes.Item('!!detail-title').TextFrame.TextRange.Text='FROM OBSERVATION TO DECISION'
 $p.Slides.Item(5).Shapes.Item('!!detail-subtitle').TextFrame.TextRange.Text='Turn evidence into a clear next step.'
 $p.Slides.Item(7).Shapes.Item('!!detail-title').TextFrame.TextRange.Text='MAKE DELIVERY REPEATABLE'
 $p.Slides.Item(7).Shapes.Item('!!detail-subtitle').TextFrame.TextRange.Text='Keep the result editable and easy to maintain.'
 foreach($i in @(4,6)){$p.Slides.Item($i).Shapes.Item('!!detail-return').TextFrame.TextRange.Text='Click blank space to continue to the second page.'}
 foreach($i in @(5,7)){$p.Slides.Item($i).Shapes.Item('!!detail-return').TextFrame.TextRange.Text='Chapter complete / click blank space to return to the preview.'}
 foreach($i in @(4,5,6,7)){$s=$p.Slides.Item($i);$s.SlideShowTransition.AdvanceOnTime=0;$s.SlideShowTransition.AdvanceOnClick=-1;$s.SlideShowTransition.EntryEffect=2051;$s.SlideShowTransition.Duration=[single]0.5}
 [void]$p.SectionProperties.AddBeforeSlide(1,'Menu states');[void]$p.SectionProperties.AddBeforeSlide(4,'Research');[void]$p.SectionProperties.AddBeforeSlide(6,'Delivery')
 $rows=@();for($i=1;$i -le $p.SectionProperties.Count;$i++){$rows+=@{name=$p.SectionProperties.Name($i);id=$p.SectionProperties.SectionID($i);first=$p.SectionProperties.FirstSlide($i);count=$p.SectionProperties.SlidesCount($i)}}
 $p.SlideShowSettings.RangeType=1;$p.SlideShowSettings.AdvanceMode=1;$p.SlideShowSettings.ShowWithAnimation=-1
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 foreach($i in @(4,6)){$p.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($Posters)) ('entry-'+$i+'.png')),'PNG',1600,900)}
 [IO.File]::WriteAllText([IO.Path]::GetFullPath($Sections),($rows|ConvertTo-Json -Depth 5),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
