param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite of evidence directory'}
New-Item -ItemType Directory -Path $Evidence|Out-Null
$taskApp=New-Object -ComObject PowerPoint.Application;$taskP=$null
try{
 $taskP=$taskApp.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,0)
 $shape=$taskP.Slides.Item(1).Shapes.Item('quarter-service-table')
 if($shape.HasTable -ne -1){throw 'PowerPoint did not recognize a real table'}
 $table=$shape.Table
 if($table.Rows.Count -ne 6 -or $table.Columns.Count -ne 5){throw 'Native table row/column mismatch'}
 $horizontal=$table.Cell(1,2).Shape.Width/$table.Columns.Item(2).Width
 $vertical=$table.Cell(1,1).Shape.Height/$table.Rows.Item(1).Height
 if([Math]::Abs($horizontal-4) -gt 0.01 -or [Math]::Abs($vertical-2) -gt 0.01){throw 'Native merge geometry does not match real 4-column / 2-row spans'}
 $taskP.Slides.Item(1).Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'table.png'),'PNG',1600,900)
 @{passed=$true;rows=$table.Rows.Count;columns=$table.Columns.Count;horizontalSpanRatio=$horizontal;verticalSpanRatio=$vertical;readOnlyNativeInspection=$true;dataRefinedWithCom=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'native-audit.json') -Encoding UTF8
 Write-Output 'Native real table and merge geometry verified; PNG exported without modifying the deck.'
}finally{if($taskP){$taskP.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($taskP)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)|Out-Null}
