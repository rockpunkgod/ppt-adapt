param([string]$OutputDirectory,[string]$OfficeCli,[string]$PythonExe='python')
$ErrorActionPreference='Stop'
$taskRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
if(-not $OutputDirectory){$OutputDirectory=Join-Path $taskRoot ('output/robomaster-'+(Get-Date -Format 'yyyyMMdd-HHmmss'))}
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(-not $OfficeCli){$OfficeCli=Join-Path $taskRoot 'tools/officecli.exe'}
& $PythonExe -X utf8 (Join-Path $PSScriptRoot 'build.py') --out $OutputDirectory --officecli $OfficeCli
if($LASTEXITCODE -ne 0){throw 'RoboMaster build failed'}
$taskDeck=Join-Path $OutputDirectory 'robomaster.pptx'
$taskNative=Join-Path $OutputDirectory '.native-save.pptx'
$taskApp=New-Object -ComObject PowerPoint.Application
$taskP=$null
try{
 $taskP=$taskApp.Presentations.Open($taskDeck,0,0,0)
 $taskP.SaveAs($taskNative,24)
}finally{if($taskP){$taskP.Close()}}
foreach($taskFile in @($taskDeck,$taskNative)){
 if(-not([IO.Path]::GetFullPath($taskFile).StartsWith($OutputDirectory.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase))){throw 'Native save path outside output directory'}
}
Copy-Item -LiteralPath $taskNative -Destination $taskDeck -Force
Remove-Item -LiteralPath $taskNative
$taskRender=Join-Path $taskRoot 'officecli-ppt-design/scripts/render_preview.ps1'
& $taskRender -InputDeck (Join-Path $OutputDirectory 'robomaster.pptx') -OutputDirectory $OutputDirectory -Video
$taskPages=[IO.Path]::GetFullPath((Join-Path $OutputDirectory '.preview-pages'))
if(-not $taskPages.StartsWith($OutputDirectory.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Preview cleanup path outside output directory'}
if(Test-Path -LiteralPath $taskPages){Remove-Item -LiteralPath $taskPages -Recurse -Force}
Write-Output ('Ready: '+$OutputDirectory)
