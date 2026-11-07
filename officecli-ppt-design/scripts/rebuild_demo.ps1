param(
 [string]$OutputDirectory,
 [string]$OfficeCli,
 [string]$PythonExe,
 [switch]$SkipVideo,
 [switch]$KeepWork
)
$ErrorActionPreference='Stop'
if(-not $OutputDirectory){$OutputDirectory=Join-Path (Get-Location) ('output/ppt-adapt-'+(Get-Date -Format 'yyyyMMdd-HHmmss'))}
if(-not $PythonExe){
 foreach($candidate in @('python','python3','py')){
  $command=Get-Command $candidate -ErrorAction SilentlyContinue
  if($command -and $command.Source -notlike '*WindowsApps*'){
   & $command.Source -c 'import sys; assert sys.version_info >= (3,10)' 2>$null
   if($LASTEXITCODE -eq 0){$PythonExe=$command.Source;break}
  }
 }
}
if(-not $PythonExe){throw 'Python 3.10+ is required. Ask your Agent to prepare the local runtime, then retry.'}
$params=@('-X','utf8',(Join-Path $PSScriptRoot 'rebuild_demo.py'),'--output',$OutputDirectory)
if($OfficeCli){$params+=@('--officecli',$OfficeCli)}
if($SkipVideo){$params+='--skip-video'}
if($KeepWork){$params+='--keep-work'}
& $PythonExe @params
if($LASTEXITCODE -ne 0){throw 'Rebuild failed. See the error above; existing presentations were not overwritten.'}
