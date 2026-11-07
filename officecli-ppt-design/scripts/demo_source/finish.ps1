param([string]$Work,[string]$OutputDeck,[string]$ChartTemplate)
$ErrorActionPreference='Stop'
function ColorValue([string]$h){return [int]('0x'+$h.Substring(4,2)+$h.Substring(2,2)+$h.Substring(0,2))}
$app=New-Object -ComObject PowerPoint.Application
$p=$null;$template=$null
try {
 $p=$app.Presentations.Open((Join-Path $Work 'main.pptx'),0,0,0)
 foreach($item in @(@{slide=1;name='cover-back';angle=7},@{slide=1;name='cover-front';angle=-7},@{slide=10;name='closing-work';angle=-6})){
  $v=$p.Slides.Item($item.slide).Shapes.Item($item.name)
  $v.Rotation=[single](($item.angle+360)%360)
  $v.Shadow.Visible=-1;$v.Shadow.ForeColor.RGB=0;$v.Shadow.Transparency=[single].3
  $v.Shadow.Blur=[single]24;$v.Shadow.OffsetX=[single]0;$v.Shadow.OffsetY=[single]12
 }
 $p.Slides.Item(3).Shapes.Item('!!case-wordmark').Height=[single](4.6*72/2.54)
 foreach($v in @($p.Slides.Item(5).Shapes)){if($v.Name.TrimStart('!') -eq 'aura-stamp'){$v.Delete()}}
 $s=$p.Slides.Item(7)
 $s.Shapes.Item('workbook-message').TextFrame.TextRange.Text="原生图表`r真实数据`r内嵌工作簿"
 $template=$app.Presentations.Open($ChartTemplate,-1,0,0)
 $s.Shapes.Item('editable-chart').Delete()
 $template.Slides.Item(1).Shapes.Item('editable-chart').Copy()
 $chartShape=$s.Shapes.Paste().Item(1);$chartShape.Name='editable-chart'
 $chartShape.Left=[single](1.3*72/2.54);$chartShape.Top=[single](6*72/2.54)
 $chartShape.Width=[single](22*72/2.54);$chartShape.Height=[single](10.1*72/2.54)
 while($s.TimeLine.MainSequence.Count -gt 0){$s.TimeLine.MainSequence.Item(1).Delete()}
 $series=$chartShape.Chart.SeriesCollection(1)
 for($i=1;$i -le 4;$i++){
  $fill=$series.Points($i).Format.Fill;$fill.Solid()
  $fill.ForeColor.RGB=ColorValue $(if($i -eq 3){'FF5B2A'}else{'8792A1'})
  $series.Points($i).Format.Line.Visible=0
 }
 $series.ApplyDataLabels();$series.DataLabels().Font.Size=19
 $series.DataLabels().Font.Color=ColorValue '111318';$series.DataLabels().Position=2
 foreach($axis in @(1,2)){$chartShape.Chart.Axes($axis).TickLabels.Font.Size=18;$chartShape.Chart.Axes($axis).TickLabels.Font.Color=ColorValue '111318'}
 $p.Slides.Item(8).Shapes.Item('pipe-body0').TextFrame.TextRange.Text="读懂受众与内容，`r确定叙事、视觉与版式。"
 $p.SaveAs($OutputDeck,24)
} finally {
 if($template){$template.Close()}
 if($p){$p.Saved=-1;$p.Close()}
 # Shared PowerPoint remains open; never close user documents.
}
