param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
$before=(Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
$ev=[IO.Path]::GetFullPath($Evidence);[IO.Directory]::CreateDirectory($ev)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$rows=@()
function Style($shape){
 $shape.TextFrame.MarginLeft=0;$shape.TextFrame.MarginRight=0;$shape.TextFrame.MarginTop=0;$shape.TextFrame.MarginBottom=0
 $shape.TextFrame.VerticalAnchor=1;$shape.TextFrame.AutoSize=0
 $shape.TextFrame.TextRange.Font.Name='Arial';$shape.TextFrame.TextRange.Font.Size=42;$shape.TextFrame.TextRange.Font.Color.RGB=5124375
 $shape.TextFrame.TextRange.ParagraphFormat.Alignment=1
}
function Snapshot([string]$phase){
 for($i=1;$i -le $p.Slides.Count;$i++){
  $s=$p.Slides.Item($i);$s.HeadersFooters.SlideNumber.Visible=-1
  try{$s.Shapes.Item('independent-number').Delete()}catch{}
  $t=$s.Shapes.AddTextbox(1,[single]($p.PageSetup.SlideWidth*1420/1600),[single]($p.PageSetup.SlideHeight*500/900),80,80)
  $t.Name='independent-number';$t.TextFrame.TextRange.Text=[string]$i;Style $t
  $s.Export((Join-Path $ev ($phase+'-'+$i+'.png')),'PNG',1600,900)
  $script:rows+=@{phase=$phase;slide=$i;slideId=$s.SlideID;visible=[int]$s.HeadersFooters.SlideNumber.Visible}
 }
}
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,0)
 if($p.Slides.Count -ne 3){throw 'Expected the three-slide CLI fixture'}
 $field=$p.SlideMaster.Shapes.AddPlaceholder(13,[single]($p.PageSetup.SlideWidth*1420/1600),[single]($p.PageSetup.SlideHeight*700/900),80,80)
 $field.Name='master-dynamic-slide-number';$field.TextFrame.TextRange.Text='';$field.TextFrame.TextRange.InsertSlideNumber()|Out-Null;Style $field
 $p.SlideMaster.HeadersFooters.SlideNumber.Visible=-1
 for($j=1;$j -le $p.SlideMaster.CustomLayouts.Count;$j++){
  $layout=$p.SlideMaster.CustomLayouts.Item($j);$layout.DisplayMasterShapes=-1
  $layoutField=$layout.Shapes.AddPlaceholder(13,[single]$field.Left,[single]$field.Top,[single]$field.Width,[single]$field.Height)
  $layoutField.Name='layout-slide-number';$layout.HeadersFooters.SlideNumber.Visible=-1
 }
 Snapshot 'initial'
 $copy=$p.Slides.Item(2).Duplicate();$copy.Item(1).MoveTo(3)
 Snapshot 'duplicated'
 $p.Slides.Item(4).MoveTo(1)
 Snapshot 'reordered'
 $saved=Join-Path $ev 'field-final.pptx';$p.SaveAs($saved,24);$p.Close();$p=$null
 $p=$app.Presentations.Open($saved,-1,0,0)
 for($i=1;$i -le $p.Slides.Count;$i++){$p.Slides.Item($i).Export((Join-Path $ev ('reopened-'+$i+'.png')),'PNG',1600,900)}
 $rows|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $ev 'states.json') -Encoding UTF8
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
if((Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash -ne $before){throw 'Source file changed'}
