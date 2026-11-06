param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
function RGB([string]$hex){return [Convert]::ToInt32($hex.Substring(0,2),16) -bor ([Convert]::ToInt32($hex.Substring(2,2),16) -shl 8) -bor ([Convert]::ToInt32($hex.Substring(4,2),16) -shl 16)}
function Plain([string]$text){return $text.Replace("`r`n","`n").Replace("`r","`n").TrimEnd([char]10)}
function Font-Set($font,[double]$size,[string]$color){$font.Name='Arial';$font.Size=$size;$font.Color.RGB=RGB $color;$font.Bold=0}
function PH($slide,[int]$type){$found=@($slide.Shapes|Where-Object {$_.Type -eq 14 -and $_.PlaceholderFormat.Type -eq $type});if($found.Count -ne 1){throw ('Expected one true placeholder type '+$type+'; found '+$found.Count)};return $found[0]}
function Snapshot([string]$phase){
 $rows=@()
 for($i=1;$i -le 4;$i++){
  $s=$taskP.Slides.Item($i);$t=PH $s 1;$b=PH $s 2;$body=@()
  for($j=1;$j -le $b.TextFrame.TextRange.Paragraphs().Count;$j++){$q=$b.TextFrame.TextRange.Paragraphs($j,1);$body+=@{text=(Plain $q.Text);level=[int]$q.IndentLevel;size=[double]$q.Font.Size;color=[int]$q.Font.Color.RGB;font=$q.Font.Name}}
  if((Plain $t.TextFrame.TextRange.Text) -ne $script:original[$i-1].title -or (Plain $b.TextFrame.TextRange.Text) -ne $script:original[$i-1].body){throw 'Local placeholder text changed'}
  $rows+=@{slide=$i;layout=$s.CustomLayout.Name;shapeCount=$s.Shapes.Count;title=@{text=(Plain $t.TextFrame.TextRange.Text);placeholderType=[int]$t.PlaceholderFormat.Type;size=[double]$t.TextFrame.TextRange.Font.Size;color=[int]$t.TextFrame.TextRange.Font.Color.RGB;font=$t.TextFrame.TextRange.Font.Name;left=[double]$t.Left;top=[double]$t.Top;width=[double]$t.Width};body=$body}
 }
 $script:states+=@{phase=$phase;slides=$rows}
 $script:states|ConvertTo-Json -Depth 16|Set-Content -LiteralPath (Join-Path $Evidence 'states.json') -Encoding UTF8
}
function Role-Readback([string]$moment){
 $row=@{moment=$moment;masterShapes=$master.Shapes.Count;levels=@()}
 for($j=1;$j -le 2;$j++){$ff=$master.TextStyles.Item(3).Levels.Item($j).Font;$row.levels+=@{level=$j;type=$ff.GetType().FullName;size=$ff.Size;color=$ff.Color.RGB}}
 $script:roleLogs+=,$row;$script:roleLogs|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $Evidence 'role-readback.json') -Encoding UTF8
}
function Save-State([string]$phase,[string]$filename){
 $taskP.SaveAs((Join-Path ([IO.Path]::GetFullPath($Evidence)) $filename),24)
 Role-Readback ('after-save-'+$phase)
 Snapshot $phase
 for($i=1;$i -le 4;$i++){$taskP.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ($phase+'-'+$i+'.png')),'PNG',1600,900)}
}
$taskApp=New-Object -ComObject PowerPoint.Application;$taskP=$null;$script:states=@();$script:original=@();$script:roleLogs=@()
try{
 $taskP=$taskApp.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,-1)
 if($taskP.Slides.Count -ne 4){throw 'Expected the four-page OfficeCLI content fixture'}
 for($i=1;$i -le 4;$i++){$s=$taskP.Slides.Item($i);$script:original+=@{title=(Plain $s.Shapes.Item('source-title').TextFrame.TextRange.Text);body=(Plain $s.Shapes.Item('source-body').TextFrame.TextRange.Text)}}
 $master=$taskP.SlideMaster
 # Different enum families: ppPlaceholderBody=2; ppBodyStyle=3.
 $masterBody=$master.Shapes.AddPlaceholder(2,60,180,800,260);$masterBody.Name='master-body-role'
 Font-Set ($master.TextStyles.Item(3).Levels.Item(1).Font) 24 '17324E'
 Font-Set ($master.TextStyles.Item(3).Levels.Item(2).Font) 20 '527086'
 $master.TextStyles.Item(3).Levels.Item(1).ParagraphFormat.Bullet.Visible=0
 $master.TextStyles.Item(3).Levels.Item(2).ParagraphFormat.Bullet.Visible=-1
 $master.TextStyles.Item(3).Levels.Item(2).ParagraphFormat.Bullet.Character=8226
 Role-Readback 'immediate-initial-style'
 $a=$master.CustomLayouts.Add($master.CustomLayouts.Count+1);$a.Name='A / Research';$b=$master.CustomLayouts.Add($master.CustomLayouts.Count+1);$b.Name='B / Delivery'
 foreach($layout in @($a,$b)){for($j=$layout.Shapes.Count;$j -ge 1;$j--){$layout.Shapes.Item($j).Delete()};$layout.DisplayMasterShapes=-1}
 $ta=$a.Shapes.AddPlaceholder(1,50,65,850,70);$ta.Name='layout-title-A';Font-Set ($ta.TextFrame.TextRange.Font) 30 '3867A8'
 $ba=$a.Shapes.AddPlaceholder(2,60,180,800,260);$ba.Name='layout-body-A'
 $tb=$b.Shapes.AddPlaceholder(1,80,85,800,70);$tb.Name='layout-title-B';Font-Set ($tb.TextFrame.TextRange.Font) 28 '467665'
 $bb=$b.Shapes.AddPlaceholder(2,80,220,790,250);$bb.Name='layout-body-B'
 for($i=1;$i -le 4;$i++){
  $s=$taskP.Slides.Item($i);$s.CustomLayout=$(if($i -le 3){$a}else{$b})
  $t=PH $s 1;$body=PH $s 2;$t.Name='role-title';$body.Name='role-body'
  $t.TextFrame.TextRange.Text=$script:original[$i-1].title
  $body.TextFrame.TextRange.Text=$script:original[$i-1].body.Replace("`n","`r")
  $body.TextFrame.TextRange.Paragraphs(1,1).IndentLevel=1;$body.TextFrame.TextRange.Paragraphs(2,1).IndentLevel=2
  $s.Shapes.Item('source-title').Delete();$s.Shapes.Item('source-body').Delete()
 }
 $directT=PH ($taskP.Slides.Item(2)) 1;$directT.Left=120;$directT.Width=700;Font-Set ($directT.TextFrame.TextRange.Font) 40 'B7633C'
 $directB=PH ($taskP.Slides.Item(2)) 2;Font-Set ($directB.TextFrame.TextRange.Paragraphs(1,1).Font) 24 'B7633C'
 Save-State 'before' '01-before.pptx'
 $ta.Left=65;$ta.Top=75;$ta.Width=820;Font-Set ($ta.TextFrame.TextRange.Font) 36 '1B7775'
 Save-State 'layout' '02-layout-update.pptx'
 Font-Set ($master.TextStyles.Item(3).Levels.Item(1).Font) 26 '1B7775';Font-Set ($master.TextStyles.Item(3).Levels.Item(2).Font) 18 '2D6972'
 Role-Readback 'immediate-updated-style'
 Save-State 'body' '03-body-role-update.pptx'
 $taskP.Slides.Item(1).CustomLayout=$b;$taskP.Slides.Item(2).CustomLayout=$b
 Save-State 'swap' '04-layout-swap.pptx'
 $taskP.Windows.Item(1).Activate();$taskP.Slides.Item(2).Select()
 if($taskApp.ActivePresentation.FullName -ne $taskP.FullName){throw 'Own presentation is not active'}
 if(-not $taskApp.CommandBars.GetEnabledMso('SlideReset')){throw 'Native SlideReset unavailable for own slide'}
 $taskApp.CommandBars.ExecuteMso('SlideReset')
 Save-State 'reset' 'placeholder-final.pptx'
 $final=Join-Path ([IO.Path]::GetFullPath($Evidence)) 'placeholder-final.pptx';$taskP.Close();$taskP=$null
 $taskP=$taskApp.Presentations.Open($final,-1,0,-1);Snapshot 'reopen'
 for($i=1;$i -le 4;$i++){$taskP.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('reopen-'+$i+'.png')),'PNG',1600,900)}
 Write-Output 'Native placeholder states created; run independent structure and pixel checks.'
}catch{@{error=$_.Exception.Message;capturedPhases=$script:states.Count}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'failure.json') -Encoding UTF8;throw}
finally{if($taskP){try{$taskP.Close()}catch{}};[Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)|Out-Null}
