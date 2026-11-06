param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Fresh evidence required'}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($Evidence))
$app=New-Object -ComObject PowerPoint.Application;$pres=$null
try{
 $pres=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,-1)
 if($pres.Slides.Count -ne 1){throw 'Expected one grid page'}
 $s=$pres.Slides.Item(1);$group=$s.Shapes.Item('seamless-grid')
 if($group.Type -ne 6 -or $group.GroupItems.Count -ne 40){throw 'Expected editable native grid group'}
 $s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'grid.png'),'PNG',1600,900)
 @{groupType=[int]$group.Type;members=[int]$group.GroupItems.Count;effects=[int]$s.TimeLine.MainSequence.Count}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'native.json') -Encoding utf8
}finally{if($pres){$pres.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
