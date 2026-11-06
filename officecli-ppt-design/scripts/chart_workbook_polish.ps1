param([string]$InputDeck,[string]$OutputDeck,[string]$Plan,[string]$Evidence,[string]$EditsPath='')
$ErrorActionPreference='Stop'
foreach($path in @($OutputDeck,$Evidence)){if(Test-Path -LiteralPath $path){throw ('Refusing overwrite: '+$path)}}
$model=Get-Content -LiteralPath $Plan -Raw -Encoding UTF8|ConvertFrom-Json
if(@($model.charts).Count -ne 1){throw 'Workbook materialization is validated for one chart per document only; multi-chart lifecycle needs separate validation'}
New-Item -ItemType Directory -Path $Evidence|Out-Null
if([IO.Path]::GetFullPath($InputDeck) -ne [IO.Path]::GetFullPath($model.source)){throw 'Plan does not match the input path'}
if((Get-FileHash -LiteralPath $InputDeck -Algorithm SHA256).Hash.ToLower() -ne $model.sourceSha256){throw 'Input changed after the plan was prepared'}
$edits=@();if($EditsPath){$edits=@(Get-Content -LiteralPath $EditsPath -Raw -Encoding UTF8|ConvertFrom-Json)}
foreach($edit in $edits){
 $matches=@($model.charts|Where-Object {$_.slide -eq $edit.slide -and $_.shape -eq $edit.shape})
 if($matches.Count -ne 1){throw 'An explicit edit does not match exactly one planned chart'}
 $entry=$matches[0]
 if([int]$edit.series -lt 1 -or [int]$edit.series -gt $entry.series.Count){throw 'Edit series is out of range'}
 if([int]$edit.point -lt 1 -or [int]$edit.point -gt $entry.series[[int]$edit.series-1].values.Count){throw 'Edit point is out of range'}
 if([double]::IsNaN([double]$edit.value) -or [double]::IsInfinity([double]$edit.value)){throw 'Edit value must be finite'}
}
function Col([int]$n){$s='';while($n -gt 0){$n--;$s=[string][char](65+$n%26)+$s;$n=[int][Math]::Floor($n/26)};return $s}
function Style-Snapshot($chart){
 $series=@();for($k=1;$k -le $chart.SeriesCollection().Count;$k++){$s=$chart.SeriesCollection($k);$series+=@{type=[int]$s.ChartType;axis=[int]$s.AxisGroup;fillType=[int]$s.Format.Fill.Type;fillColor=$s.Format.Fill.ForeColor.RGB;lineColor=$s.Format.Line.ForeColor.RGB}}
 $axes=@();foreach($axisType in @(1,2)){try{$ax=$chart.Axes($axisType,1);$axes+=@{type=$axisType;minimum=$ax.MinimumScale;maximum=$ax.MaximumScale}}catch{}}
 return @{series=$series;axes=$axes}
}
$taskApp=New-Object -ComObject PowerPoint.Application;$taskP=$null;$taskW=$null;$log=@()
try{
 $taskP=$taskApp.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 foreach($entry in $model.charts){
  $slide=$taskP.Slides.Item([int]$entry.slide);$chart=$slide.Shapes.Item($entry.shape).Chart
  $slide.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('before-'+$entry.slide+'.png')),'PNG',1600,900)
  $styleBefore=Style-Snapshot $chart;$chart.ChartData.Activate();$taskW=$chart.ChartData.Workbook;$ws=$taskW.Worksheets.Item(1)
  $ws.Cells.Clear()|Out-Null;$ws.Cells.Item(1,1).Value2='Category'
  $sheet=$ws.Name.Replace("'","''");$n=$entry.series[0].values.Count
  for($row=0;$row -lt $n;$row++){$cell=$ws.Cells.Item($row+2,1);$cell.NumberFormat='@';$cell.Value2=[string]$entry.series[0].categories[$row]}
  for($j=0;$j -lt $entry.series.Count;$j++){
   $spec=$entry.series[$j];$column=Col ($j+2);$cell=$ws.Cells.Item(1,$j+2);$cell.NumberFormat='@';$cell.Value2=[string]$spec.name
   for($row=0;$row -lt $n;$row++){$ws.Cells.Item($row+2,$j+2).Value2=[double]$spec.values[$row]}
   $ser=$chart.SeriesCollection($j+1)
   $ser.Name="='$sheet'!"+'$'+$column+'$1'
   $ser.XValues="='$sheet'!"+'$A$2:$A$'+($n+1)
   $ser.Values="='$sheet'!"+'$'+$column+'$2:$'+$column+'$'+($n+1)
  }
  $applied=@()
  foreach($edit in $edits){if($edit.slide -eq $entry.slide -and $edit.shape -eq $entry.shape){
   $ws.Cells.Item([int]$edit.point+1,[int]$edit.series+1).Value2=[double]$edit.value;$applied+=,$edit
  }}
  $chart.Refresh()
  for($j=0;$j -lt $entry.series.Count;$j++){
   $vals=@($chart.SeriesCollection($j+1).Values);$want=@($entry.series[$j].values)
   foreach($edit in $applied){if($edit.series -eq $j+1){$want[[int]$edit.point-1]=[double]$edit.value}}
   if($vals.Count -ne $want.Count){throw 'Native series point count mismatch'}
   for($k=0;$k -lt $vals.Count;$k++){if([double]$vals[$k] -ne [double]$want[$k]){throw 'Workbook cell change did not update native series values'}}
  }
  $styleAfter=Style-Snapshot $chart
  if(($styleBefore|ConvertTo-Json -Depth 8 -Compress) -ne ($styleAfter|ConvertTo-Json -Depth 8 -Compress)){throw 'Workbook binding altered chart type, fill or axis range'}
  $log+=@{slide=$entry.slide;shape=$entry.shape;stylesBefore=$styleBefore;stylesAfter=$styleAfter;edits=$applied;sheet=$ws.Name;seriesCount=$chart.SeriesCollection().Count}
  $taskW.Close($true);$taskW=$null
 }
 $taskP.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 foreach($entry in $model.charts){$taskP.Slides.Item([int]$entry.slide).Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('after-'+$entry.slide+'.png')),'PNG',1600,900)}
 $log|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $Evidence 'com-audit.json') -Encoding UTF8
 Write-Output 'Real embedded workbooks bound; native values and formats checked.'
}finally{if($taskW){$taskW.Close($true)};if($taskP){$taskP.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($taskP)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)|Out-Null}
