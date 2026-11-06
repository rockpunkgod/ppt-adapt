param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $EvidenceDirectory){throw 'Fresh evidence required'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$before=(Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash;$app=New-Object -ComObject PowerPoint.Application;$p=$null;$rows=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,-1)
 foreach($i in @(1,2,3)){$s=$p.Slides.Item($i);$s.Export((Join-Path $EvidenceDirectory ('native-'+$i+'.png')),'PNG',1920,1080);$rows+=@{slide=$i;shapeCount=$s.Shapes.Count;originalText=$s.Shapes.Item('original-text').TextFrame2.TextRange.Text}}
 $saved=Join-Path $EvidenceDirectory 'roundtrip.pptx';$p.SaveAs([IO.Path]::GetFullPath($saved),24);$p.Close();$p=$null
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($saved),0,0,-1)
 foreach($i in @(1,2,3)){$p.Slides.Item($i).Export((Join-Path $EvidenceDirectory ('reopened-'+$i+'.png')),'PNG',1920,1080)}
 $after=(Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'native-readback.json'),(@{slides=$rows;before=$before;after=$after;unchanged=($before -eq $after)}|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
