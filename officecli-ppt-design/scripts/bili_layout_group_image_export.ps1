param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite of evidence directory'}
New-Item -ItemType Directory -Path $Evidence|Out-Null
$taskApp=New-Object -ComObject PowerPoint.Application;$taskP=$null
try{
 $taskP=$taskApp.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,0)
 $proof=@{slides=$taskP.Slides.Count;readOnlyInspection=$true;groups=@()}
 for($i=2;$i -le 3;$i++){$g=$taskP.Slides.Item($i).Shapes.Item($(if($i -eq 2){'mapped-adjacent'}else{'mapped-separated'}));if($g.Type -ne 6 -or $g.GroupItems.Count -ne 4){throw 'Expected four real grouped pieces'};$proof.groups+=@{slide=$i;type=$g.Type;children=$g.GroupItems.Count}}
 for($i=1;$i -le $taskP.Slides.Count;$i++){$taskP.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('slide-'+$i+'.png')),'PNG',1600,900)}
 $proof|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $Evidence 'native-audit.json') -Encoding UTF8
 Write-Output 'Native inherited-image groups exported read-only.'
}finally{if($taskP){$taskP.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($taskP)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)|Out-Null}
