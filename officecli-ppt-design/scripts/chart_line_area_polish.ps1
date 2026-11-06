param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Evidence,[switch]$DemoEdit)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'New output paths required; overwrite disabled'}
[void][IO.Directory]::CreateDirectory($Evidence)
$a=New-Object -ComObject PowerPoint.Application;$p=$null;$w=$null
try{
$p=$a.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$s=$p.Slides.Item(1);$s.Select()
$s.Export((Join-Path $Evidence 'before.png'),'PNG',1600,900)
$c=$s.Shapes.Item('revenue-line-area').Chart
$fill=$c.SeriesCollection(1);$line=$c.SeriesCollection(2)
if($fill.ChartType -ne 1 -or $line.ChartType -notin @(4,65)){throw "Expected area1/line4-or65; got $($fill.ChartType),$($line.ChartType)"}
$fill.Format.Fill.Visible=-1;$fill.Format.Fill.ForeColor.RGB=12417062;$fill.Format.Fill.BackColor.RGB=12417062
$fill.Format.Fill.TwoColorGradient(1,1)
$fill.Format.Fill.GradientStops.Item(1).Transparency=0.26
$fill.Format.Fill.GradientStops.Item(2).Transparency=1
$fill.Format.Line.Visible=0
$line.Format.Line.Visible=-1;$line.Format.Line.ForeColor.RGB=12417062;$line.Format.Line.Weight=3
$line.Smooth=0;$line.MarkerStyle=8;$line.MarkerSize=9;$line.MarkerBackgroundColor=16777215;$line.MarkerForegroundColor=12417062
$c.Legend.LegendEntries(1).Delete()
# Activate materializes an embedded workbook from OfficeCLI literal series.
$c.ChartData.Activate();$w=$c.ChartData.Workbook;$ws=$w.Worksheets.Item(1)
$fill.Values="='$($ws.Name)'!"+'$B$2:$B$7';$fill.XValues="='$($ws.Name)'!"+'$A$2:$A$7'
$line.Values="='$($ws.Name)'!"+'$D$2:$D$7';$line.XValues="='$($ws.Name)'!"+'$C$2:$C$7'
# Optional fixture-only edit probe. Default polishing preserves all input values.
if($DemoEdit){
 $ws.Cells.Item(3,2).Value2=[double]31;$ws.Cells.Item(3,4).Value2=[double]31;$c.Refresh()
 if(-not (@($line.Values) -contains 31) -or -not (@($fill.Values) -contains 31)){throw 'Native workbook edit not visible in both series'}
}
$p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
$s.Export((Join-Path $Evidence 'after.png'),'PNG',1600,900)
@{passed=$true;nativeWorkbookEditTo31=[bool]$DemoEdit;seriesCount=$c.SeriesCollection().Count;types=@([int]$fill.ChartType,[int]$line.ChartType);samePrimaryAxes=($fill.AxisGroup -eq 1 -and $line.AxisGroup -eq 1);min=$c.Axes(2).MinimumScale;max=$c.Axes(2).MaximumScale;legendEntries=$c.Legend.LegendEntries().Count;smooth=[bool]$line.Smooth}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'result.json') -Encoding utf8
}finally{if($w){$w.Close($true)};if($p){$p.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)}
