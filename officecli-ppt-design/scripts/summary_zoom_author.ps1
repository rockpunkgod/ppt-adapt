param([Parameter(Mandatory)][string]$Deck,[Parameter(Mandatory)][string]$Out)
$ErrorActionPreference='Stop'
$Out=[IO.Path]::GetFullPath($Out)
foreach($name in @('summary-sectioned.pptx','summary-edited.pptx','resize-probe','author.json','author-progress.json','summary-edited.png')){if(Test-Path -LiteralPath (Join-Path $Out $name)){throw "Refusing overwrite: $name"}}
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
. (Join-Path $PSScriptRoot 'summary_zoom_ready.ps1')
$D=[Windows.Automation.TreeScope]::Descendants
function Property($prop,$value){[Windows.Automation.PropertyCondition]::new($prop,$value)}
function Named($root,$name){$root.FindFirst($D,(Property ([Windows.Automation.AutomationElement]::NameProperty) $name))}
function Dialog($root,$name){$root.FindFirst($D,[Windows.Automation.AndCondition]::new((Property ([Windows.Automation.AutomationElement]::NameProperty) $name),(Property ([Windows.Automation.AutomationElement]::ControlTypeProperty) ([Windows.Automation.ControlType]::Window))))}
function Identified($root,$id){$root.FindFirst($D,(Property ([Windows.Automation.AutomationElement]::AutomationIdProperty) $id))}
function WaitNode([scriptblock]$find,[string]$label){$sw=[Diagnostics.Stopwatch]::StartNew();do{$node=&$find;if($node){return $node};Start-Sleep -Milliseconds 75}while($sw.Elapsed.TotalSeconds -lt 3);throw "Own UI control absent: $label"}
function InvokeNode($e){$e.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()}
function Dump($root){@($root.FindAll($D,[Windows.Automation.Condition]::TrueCondition)|ForEach-Object {$c=$_.Current;if(-not $c.IsOffscreen -and $c.Name){@{name=$c.Name;id=$c.AutomationId;type=$c.ControlType.ProgrammaticName;patterns=@($_.GetSupportedPatterns()|ForEach-Object ProgrammaticName)}}})}
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$w=$null;$rows=@{}
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,-1)
 if($p.Slides.Count -ne 6){throw 'Expected six fresh OfficeCLI foundation slides'}
 foreach($item in @(@(1,'Field Notes'),@(3,'Signal Lab'),@(5,'Next Route'))){[void]$p.SectionProperties.AddBeforeSlide([int]$item[0],[string]$item[1])}
 $path=Join-Path $Out 'summary-sectioned.pptx';$p.SaveAs($path,24)
 $leaf=[IO.Path]::GetFileNameWithoutExtension($path)
 $w=WaitNode { $candidates=@([Windows.Automation.AutomationElement]::RootElement.FindAll([Windows.Automation.TreeScope]::Children,[Windows.Automation.Condition]::TrueCondition)|Where-Object {$_.Current.ClassName -eq 'PPTFrameClass' -and $_.Current.Name.Contains($leaf)});if($candidates.Count -gt 1){throw 'Ambiguous own Normal window'};if($candidates.Count -eq 1){$candidates[0]} } 'own Normal window'
 $tab=WaitNode {Identified $w 'TabInsert'} 'Insert tab';$tab.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
 $menu=WaitNode {Identified $w 'InsertZoomMenu'} 'Zoom menu';$menu.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Expand()
 InvokeNode (WaitNode {Named $w '摘要缩放定位'} 'visible Summary Zoom menu')
 $dialog=WaitNode {Dialog $w '插入摘要缩放定位'} 'Insert Summary dialog';$rows.insertDialog=Dump $dialog
 $choices=@();foreach($e in $dialog.FindAll($D,(Property ([Windows.Automation.AutomationElement]::ControlTypeProperty) ([Windows.Automation.ControlType]::CheckBox)))){
  $name=$e.Current.Name;$t=$e.GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern)
  if($name -match '^幻灯片 ([0-9]+)'){$wanted=[int]$Matches[1] -in @(1,3)}else{$wanted=$true}
  if(($t.Current.ToggleState -eq [Windows.Automation.ToggleState]::On) -ne $wanted){$t.Toggle()};$choices+=@{name=$name;on=[int]$t.Current.ToggleState}
 };$rows.insertChoices=$choices;InvokeNode (Named $dialog '插入')
 $sw=[Diagnostics.Stopwatch]::StartNew();while($p.Slides.Count -ne 7 -and $sw.Elapsed.TotalSeconds -lt 3){Start-Sleep -Milliseconds 75};if($p.Slides.Count -ne 7){throw 'Native Summary insertion did not add a slide'}
 $rows.initialReady=Wait-SummaryPosters $p.Slides.Item(1) $Out 'initial-ready' 2;$p.Save();$s=$p.Slides.Item(1).Shapes.Item(1);$resize=Join-Path $Out 'resize-probe';[IO.Directory]::CreateDirectory($resize)|Out-Null
 $rows.nativeObject=@{type=[int]$s.Type;name=$s.Name;slideCount=$p.Slides.Count;sections=$p.SectionProperties.Count}
 $originalWidth=$s.Width;$originalLeft=$s.Left
 foreach($phase in @('initial','wide','narrow')){
  if($phase -ne 'initial'){$s.LockAspectRatio=0;$s.Width=[single]($p.PageSetup.SlideWidth*$(if($phase -eq 'wide'){0.90}else{0.38}));$s.Left=[single](($p.PageSetup.SlideWidth-$s.Width)/2)}
  $p.Slides.Item(1).Export((Join-Path $resize ($phase+'.png')),'PNG',1600,900);$p.SaveCopyAs((Join-Path $resize ($phase+'.pptx')),24)
 }
 $s.Width=[single]($p.PageSetup.SlideWidth*0.90);$s.Left=[single](($p.PageSetup.SlideWidth-$s.Width)/2)
 $p.Slides.Item(1).Select();$s.Select()
 $zoom=WaitNode {Named $w '缩放'} 'native contextual Zoom tab';$zoom.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
 $edit=WaitNode {Identified $w 'UpdateInteractiveIndex'} 'Edit Summary';$rows.editButton=@{name=$edit.Current.Name;id=$edit.Current.AutomationId};InvokeNode $edit
 $dialog=WaitNode {Dialog $w '编辑摘要缩放定位'} 'Edit Summary dialog';$rows.editDialog=Dump $dialog;$rows|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $Out 'author-progress.json') -Encoding UTF8
 $last=WaitNode {$dialog.FindAll($D,(Property ([Windows.Automation.AutomationElement]::ControlTypeProperty) ([Windows.Automation.ControlType]::CheckBox)))|Where-Object {$_.Current.Name -match 'Next Route$'}|Select-Object -First 1} 'Next Route section'
 $toggle=$last.GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern);if($toggle.Current.ToggleState -ne [Windows.Automation.ToggleState]::On){$toggle.Toggle()}
 InvokeNode (Named $dialog '更新');Start-Sleep -Milliseconds 200
 $rows.editedReady=Wait-SummaryPosters $p.Slides.Item(1) $Out 'edited-ready' 3;$p.SaveAs((Join-Path $Out 'summary-edited.pptx'),24);[IO.File]::Copy($rows.editedReady.finalImage,(Join-Path $Out 'summary-edited.png'))
 $rows.completed=$true;$rows.nativeUiCreation=$true;$rows|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $Out 'author.json') -Encoding UTF8
}finally{
 if($w){foreach($name in @('插入摘要缩放定位','编辑摘要缩放定位')){try{$modal=Dialog $w $name;if($modal){$cancel=Named $modal '取消';if($cancel){InvokeNode $cancel}}}catch{}}}
 if($p){$p.Saved=-1;$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
