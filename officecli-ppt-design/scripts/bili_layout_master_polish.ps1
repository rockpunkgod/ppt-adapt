param([string]$Deck,[string]$Output,[string]$RenderDirectory)
$ErrorActionPreference='Stop'
$final=[IO.Path]::GetFullPath($Output);$folder=[IO.Path]::GetDirectoryName($final);$stem=[IO.Path]::GetFileNameWithoutExtension($final)
$before=Join-Path $folder ($stem+'.before.pptx');$masterUpdate=Join-Path $folder ($stem+'.master-update.pptx')
foreach($path in @($final,$before,$masterUpdate,$RenderDirectory)){if(Test-Path -LiteralPath $path){throw ('Refusing overwrite: '+$path)}}
New-Item -ItemType Directory -Path $RenderDirectory|Out-Null
function RGB-Value([string]$hex){return [Convert]::ToInt32($hex.Substring(0,2),16) -bor ([Convert]::ToInt32($hex.Substring(2,2),16) -shl 8) -bor ([Convert]::ToInt32($hex.Substring(4,2),16) -shl 16)}
function Assert-That($value,[string]$message){if(-not $value){throw $message}}
function Add-Label($shapes,[string]$name,[string]$text,[double]$x,[double]$y,[double]$w,[double]$h,[string]$fill,[int]$fontSize){
 $shape=$shapes.AddShape(1,$x,$y,$w,$h);$shape.Name=$name;$shape.Line.Visible=0
 $shape.Fill.Solid();$shape.Fill.ForeColor.RGB=RGB-Value $fill
 $range=$shape.TextFrame2.TextRange;$range.Text=$text;$range.Font.Name='Microsoft YaHei';$range.Font.NameFarEast='Microsoft YaHei';$range.Font.Size=$fontSize;$range.Font.Fill.ForeColor.RGB=RGB-Value 'FFFFFF'
 $shape.TextFrame2.MarginLeft=12;$shape.TextFrame2.MarginRight=8;$shape.TextFrame2.MarginTop=3;$shape.TextFrame2.MarginBottom=3;$shape.TextFrame2.VerticalAnchor=3
 return $shape
}
function Local-Snapshot($p){$rows=@();for($i=1;$i -le 3;$i++){$texts=@{};foreach($s in $p.Slides.Item($i).Shapes){if($s.Name -like 'local-*'){$texts[$s.Name]=$s.TextFrame2.TextRange.Text}};$rows+=@{slide=$i;texts=$texts;localShapeCount=$p.Slides.Item($i).Shapes.Count;layout=$p.Slides.Item($i).CustomLayout.Name}};return $rows}
function Export-State($p,[string]$phase){for($i=1;$i -le 3;$i++){$p.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($RenderDirectory)) ($phase+'-'+$i+'.png')),'PNG',1600,900)}}
$taskApp=New-Object -ComObject PowerPoint.Application;$taskP=$null
try{
 $taskP=$taskApp.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,0);$master=$taskP.SlideMaster
 $initialTexts=Local-Snapshot $taskP
 $shared=Add-Label $master.Shapes 'master-shared-brand' 'NOVA / STUDIO' 45 24 870 34 '3867A8' 16
 $layoutA=$master.CustomLayouts.Add($master.CustomLayouts.Count+1);$layoutA.Name='Branch A / Research';$layoutA.DisplayMasterShapes=-1
 $layoutB=$master.CustomLayouts.Add($master.CustomLayouts.Count+1);$layoutB.Name='Branch B / Delivery';$layoutB.DisplayMasterShapes=-1
 foreach($layout in @($layoutA,$layoutB)){for($j=$layout.Shapes.Count;$j -ge 1;$j--){$layout.Shapes.Item($j).Delete()}}
 $aLabel=Add-Label $layoutA.Shapes 'layout-A-strip' 'A / RESEARCH' 45 92 870 40 '3867A8' 18
 $bLabel=Add-Label $layoutB.Shapes 'layout-B-strip' 'B / DELIVERY' 45 92 870 40 '467665' 18
 for($i=1;$i -le 3;$i++){$slide=$taskP.Slides.Item($i);$slide.CustomLayout=if($i -le 2){$layoutA}else{$layoutB};$slide.DisplayMasterShapes=-1}
 $baseline=Local-Snapshot $taskP
 Assert-That (($baseline[0].localShapeCount -eq 3) -and ($baseline[1].localShapeCount -eq 3) -and ($baseline[2].localShapeCount -eq 3)) 'Inherited labels must not be copied into slide shapes'
 $taskP.SaveAs($before);Export-State $taskP 'before'
 $shared.TextFrame2.TextRange.Text='NOVA / SYSTEM 2026';$shared.Fill.ForeColor.RGB=RGB-Value '1B7775'
 $taskP.SaveAs($masterUpdate);Export-State $taskP 'master'
 $bState=@{text=$bLabel.TextFrame2.TextRange.Text;color=$bLabel.Fill.ForeColor.RGB}
 $aLabel.TextFrame2.TextRange.Text='A / RESEARCH UPDATED';$aLabel.Fill.ForeColor.RGB=RGB-Value 'B7633C'
 Assert-That ($bLabel.TextFrame2.TextRange.Text -eq $bState.text -and $bLabel.Fill.ForeColor.RGB -eq $bState.color) 'Changing A contaminated B layout'
 $finalTexts=Local-Snapshot $taskP
 for($i=0;$i -lt 3;$i++){
  foreach($key in @('local-title','local-body','local-key')){Assert-That ($initialTexts[$i].texts[$key] -eq $finalTexts[$i].texts[$key]) 'Local slide content changed'}
 }
 $taskP.SaveAs($final);$taskP.Close();$taskP=$null
 $taskP=$taskApp.Presentations.Open($final,-1,0,0);Export-State $taskP 'after'
 $reopened=Local-Snapshot $taskP
 for($i=0;$i -lt 3;$i++){foreach($key in @('local-title','local-body','local-key')){Assert-That ($initialTexts[$i].texts[$key] -eq $reopened[$i].texts[$key]) 'Reopen changed local content'}}
 @{passed=$true;initial=$initialTexts;before=$baseline;after=$finalTexts;reopened=$reopened;layouts=@($reopened[0].layout,$reopened[2].layout);masterShape='master-shared-brand';scope='One shared master, A pages 1/2, B page 3. Master update then A-only update, retaining local text.'}|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $RenderDirectory 'com-audit.json') -Encoding UTF8
 Write-Output 'Native inheritance before/master/after states exported.'
}finally{if($taskP){$taskP.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($taskP)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)|Out-Null}
