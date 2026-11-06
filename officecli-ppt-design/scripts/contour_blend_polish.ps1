param([string]$Deck,[string]$Output,[string]$Evidence)
$ErrorActionPreference='Stop'
foreach($x in @($Output,$Evidence)){if(Test-Path -LiteralPath $x){throw ('Refusing overwrite '+$x)}}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
function Merge-Names($slide,[object[]]$names,[int]$mode,$primary){$slide.Select();$old=@($slide.Shapes|ForEach-Object Id);$range=$slide.Shapes.Range($names);$range.MergeShapes($mode,$primary);return @($slide.Shapes|Where-Object {$_.Id -notin $old})}
function Snapshot($p,[string]$phase){
 $rows=@()
 foreach($spec in @(@(1,'endpoint-near'),@(1,'solid-100'),@(1,'solid-000'),@(1,'solid-front-cap'),@(2,'long-shadow'),@(3,'iso-dense-100'),@(3,'iso-sparse-015'))){
  $s=$p.Slides.Item([int]$spec[0]).Shapes.Item([string]$spec[1]);$nodes=[int]$s.Nodes.Count
  $rows+=@{slide=[int]$spec[0];name=$s.Name;type=[int]$s.Type;nodes=$nodes;box=@($s.Left,$s.Top,$s.Width,$s.Height)}
 }
 return @{phase=$phase;shapes=$rows;slideShapeCounts=@($p.Slides|ForEach-Object {$_.Shapes.Count})}
}
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$log=@();$phase='open'
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,-1);$s=$p.Slides.Item(2);$s.Select()
 $shadowNames=@($s.Shapes|Where-Object {$_.Name -match '^shadow-(start|end|edge-)'}|ForEach-Object Name);if($shadowNames.Count -ne 106){throw 'Expected106 exact sweep primitives'}
 $phase='union';$primary=$s.Shapes.Item('shadow-start');$r=@(Merge-Names $s $shadowNames 1 $primary);if($r.Count -ne 1){throw 'Union did not yield one editable vector'};$r[0].Name='long-shadow-unclipped'
 $before=@{type=[int]$r[0].Type;nodes=[int]$r[0].Nodes.Count;box=@($r[0].Left,$r[0].Top,$r[0].Width,$r[0].Height)}
 $phase='clip';$primary=$r[0];$r=@(Merge-Names $s @('long-shadow-unclipped','shadow-clip') 3 $primary);if($r.Count -ne 1){throw 'Intersect did not yield one vector'};$r[0].Name='long-shadow'
 $r[0].Fill.Solid();$r[0].Fill.ForeColor.RGB=0x4E3217;$r[0].Fill.Transparency=[single]0.82;$r[0].Line.Visible=0;$s.Shapes.Item('shadow-front-cap').ZOrder(0)
 $phase='save';$log+=Snapshot $p 'before-save';$p.SaveAs([IO.Path]::GetFullPath($Output),24);$p.Close();$p=$null
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Output),-1,0,0);$log+=Snapshot $p 'reopen'
 for($i=1;$i -le 3;$i++){$p.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('slide-'+$i+'.png')),'PNG',1600,900)}
 @{passed=$true;unionInputCount=$shadowNames.Count;unclipped=$before;states=$log;scope='True native Union/Intersect; no Multiply claim'}|ConvertTo-Json -Depth 15|Set-Content -LiteralPath (Join-Path $Evidence 'native-objects.json') -Encoding UTF8
}catch{@{passed=$false;phase=$phase;error=$_.Exception.Message}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'failure.json') -Encoding UTF8;throw}
finally{if($p){$p.Saved=-1;$p.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
