param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
function Snapshot($owner,[string]$name){$effects=@();foreach($e in $owner.TimeLine.MainSequence){$effects+=@{shape=$e.Shape.Name;duration=$e.Timing.Duration;autoReverse=[int]$e.Timing.AutoReverse;repeatCount=$e.Timing.RepeatCount;repeatDuration=$e.Timing.RepeatDuration;trigger=[int]$e.Timing.TriggerType}};return @{owner=$name;shapeNames=@($owner.Shapes|ForEach-Object Name);effects=$effects}}
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),-1,0,0);$master=$p.SlideMaster
 $rows=@((Snapshot $master 'master'),(Snapshot $p.Slides.Item(1).CustomLayout 'layout-A'),(Snapshot $p.Slides.Item(3).CustomLayout 'layout-B'))
 for($i=1;$i -le 3;$i++){$s=$p.Slides.Item($i);$rows+=Snapshot $s ('slide-'+$i);$s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('slide-'+$i+'.png')),'PNG',1600,900)}
 $rows|ConvertTo-Json -Depth 15|Set-Content -LiteralPath (Join-Path $Evidence 'objects.json') -Encoding UTF8
}finally{if($p){$p.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
