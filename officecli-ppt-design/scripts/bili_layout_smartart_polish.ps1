param([string]$Deck,[string]$Output,[string]$RenderDirectory,[string]$Landscape,[string]$Portrait)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Output){throw 'Refusing overwrite'}
if(Test-Path -LiteralPath $RenderDirectory){throw 'Choose a new render directory'}
New-Item -ItemType Directory -Path $RenderDirectory|Out-Null
$taskApp=New-Object -ComObject PowerPoint.Application;$taskP=$null
function Assert-That($condition,[string]$message){if(-not $condition){throw $message}}
function Cover-Jpeg([string]$source,[double]$ratio,[string]$target){
 Add-Type -AssemblyName System.Drawing
 $img=[Drawing.Image]::FromFile($source);$canvas=$null;$g=$null
 try{
  $srcRatio=$img.Width/$img.Height;$sw=[double]$img.Width;$sh=[double]$img.Height
  if($srcRatio -gt $ratio){$sw=$sh*$ratio}else{$sh=$sw/$ratio}
  $canvas=New-Object Drawing.Bitmap(1000,[int][Math]::Round(1000/$ratio));$g=[Drawing.Graphics]::FromImage($canvas)
  $g.InterpolationMode=[Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.DrawImage($img,[Drawing.Rectangle]::new(0,0,$canvas.Width,$canvas.Height),[single](($img.Width-$sw)/2),[single](($img.Height-$sh)/2),[single]$sw,[single]$sh,[Drawing.GraphicsUnit]::Pixel)
  $canvas.Save($target,[Drawing.Imaging.ImageFormat]::Jpeg)
 }finally{if($g){$g.Dispose()};if($canvas){$canvas.Dispose()};$img.Dispose()}
}
function Snapshot($art){$rows=@();foreach($node in $art.AllNodes){$rows+=@{text=$node.TextFrame2.TextRange.Text;level=$node.Level}};return $rows}
function Find-Node($art,[string]$text){for($j=1;$j -le $art.AllNodes.Count;$j++){$n=$art.AllNodes.Item($j);if($n.TextFrame2.TextRange.Text -eq $text){return $n}};throw 'Expected node text not found'}
function Populate($art){
 while($art.AllNodes.Count -gt 1){$art.AllNodes.Item($art.AllNodes.Count).Delete()}
 $root=$art.AllNodes.Item(1);$root.TextFrame2.TextRange.Text='产品研究'
 $one=$root.AddNode(5,1);$one.TextFrame2.TextRange.Text='需求洞察'
 $two=$one.AddNode(2,1);$two.TextFrame2.TextRange.Text='原型测试'
 $leaf=$two.AddNode(5,1);$leaf.TextFrame2.TextRange.Text='反馈归纳'
 Assert-That ($root.Level -eq 1 -and $one.Level -eq 2 -and $two.Level -eq 2 -and $leaf.Level -eq 3) 'Native levels are incorrect'
 $temp=$one.AddNode(2,1);$temp.TextFrame2.TextRange.Text='临时节点'
 Assert-That ($art.AllNodes.Count -eq 5) 'AddNode did not create the fifth node'
 $temp.Delete();Assert-That ($art.AllNodes.Count -eq 4) 'Delete did not remove only the temporary node'
 $leaf=Find-Node $art '反馈归纳';$prePromote=Snapshot $art
 $leaf.Promote();Start-Sleep -Milliseconds 100;$leaf=Find-Node $art '反馈归纳'
 $postPromote=Snapshot $art
 @{before=$prePromote;after=$postPromote;level=$leaf.Level}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $RenderDirectory 'promote-trace.json') -Encoding UTF8
 Assert-That ($leaf.Level -eq 2) 'Promote failed'
 $leaf.Demote();Start-Sleep -Milliseconds 100;$leaf=Find-Node $art '反馈归纳'
 Assert-That ($leaf.Level -eq 3 -and $leaf.ParentNode.TextFrame2.TextRange.Text -eq '原型测试') 'Demote changed parent relation'
 foreach($n in $art.AllNodes){$n.TextFrame2.TextRange.Font.Name='Microsoft YaHei';$n.TextFrame2.TextRange.Font.NameFarEast='Microsoft YaHei';$n.TextFrame2.TextRange.Font.Size=22}
}
try{
 # This build requires a real visible document window for SmartArtLayouts initialization.
 $taskApp.Visible=-1;$taskP=$taskApp.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,-1)
 $collection=$taskApp.SmartArtLayouts;$all=@();$layoutObjects=@()
 for($index=1;$index -le $collection.Count;$index++){$l=$collection.Item($index);$all+=@{name=$l.Name;id=$l.Id};$layoutObjects+=,$l}
 ConvertTo-Json -InputObject @{count=$collection.Count;layouts=$all} -Depth 5|Set-Content -LiteralPath (Join-Path $RenderDirectory 'layouts.json') -Encoding UTF8
 # IDs below were read from this installed native catalog; do not invent layout GUIDs.
 $vertical=@($layoutObjects|Where-Object {$_.Id -eq 'urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1'})[0]
 $horizontal=@($layoutObjects|Where-Object {$_.Id -eq 'urn:microsoft.com/office/officeart/2005/8/layout/hierarchy2'})[0]
 Assert-That ($null -ne $vertical -and $null -ne $horizontal) 'Required hierarchy layouts not found; inspect layouts.json'
 $a=$taskP.Slides.Item(1).Shapes.AddSmartArt($vertical,55,152,850,310);$a.Name='native-hierarchy-before';Populate $a.SmartArt
 $before=Snapshot $a.SmartArt
 $b=$taskP.Slides.Item(2).Shapes.AddSmartArt($vertical,55,152,850,310);$b.Name='native-hierarchy-after';Populate $b.SmartArt
 $old=Snapshot $b.SmartArt;$b.SmartArt.Layout=$horizontal;$after=Snapshot $b.SmartArt
 Assert-That (($old|ConvertTo-Json -Compress) -eq ($after|ConvertTo-Json -Compress)) 'Layout change altered text or levels'
 $pictureStatus=@{passed=$false;reason='Picture layout not available'}
 $picLayout=@($layoutObjects|Where-Object {$_.Id -eq 'urn:microsoft.com/office/officeart/2005/8/layout/pList1'})[0]
 if($null -ne $picLayout){
  $c=$taskP.Slides.Item(3).Shapes.AddSmartArt($picLayout,75,152,810,310);$c.Name='native-picture-pairs'
  while($c.SmartArt.AllNodes.Count -gt 2){$c.SmartArt.AllNodes.Item($c.SmartArt.AllNodes.Count).Delete()}
  $assets=@([IO.Path]::GetFullPath($Landscape),[IO.Path]::GetFullPath($Portrait));$labels=@('日落','步道');$audit=@();$ok=$true
  for($i=1;$i -le 2;$i++){
   $n=$c.SmartArt.AllNodes.Item($i);$n.TextFrame2.TextRange.Text=$labels[$i-1]
   $n.TextFrame2.TextRange.Font.Name='Microsoft YaHei';$n.TextFrame2.TextRange.Font.NameFarEast='Microsoft YaHei';$n.TextFrame2.TextRange.Font.Size=22
   $candidates=@()
   foreach($shape in $n.Shapes){if($shape.HasTextFrame -eq 0 -or $shape.TextFrame2.HasText -eq 0){$candidates+=,$shape}}
   if($candidates.Count -eq 1){
    $picture=$candidates[0];$prepared=Join-Path ([IO.Path]::GetFullPath($RenderDirectory)) ('picture-'+$i+'.jpg')
    Cover-Jpeg $assets[$i-1] ($picture.Width/$picture.Height) $prepared
    $picture.Fill.UserPicture($prepared)
    $audit+=@{text=$labels[$i-1];fillType=$picture.Fill.Type;source=[IO.Path]::GetFileName($assets[$i-1]);prepared=[IO.Path]::GetFileName($prepared);width=$picture.Width;height=$picture.Height}
   }
   else{$ok=$false;$audit+=@{text=$labels[$i-1];emptyShapeCandidates=$candidates.Count}}
  }
  $pictureStatus=@{passed=$ok;layout=$picLayout.Id;nodes=$audit;reason='Requires independent native visual and embedded-image pairing check'}
 }
 $proof=@{nodeOperationsPassed=$true;vertical=$vertical.Id;horizontal=$horizontal.Id;before=$before;after=$after;picturePairing=$pictureStatus}
 $taskP.SaveAs([IO.Path]::GetFullPath($Output))
 $taskP.Close();$taskP=$null
 $taskP=$taskApp.Presentations.Open([IO.Path]::GetFullPath($Output),-1,0,0)
 Assert-That ($taskP.Slides.Item(1).Shapes.Item('native-hierarchy-before').HasSmartArt -eq -1) 'Reopened shape is not native SmartArt'
 $reopened=Snapshot $taskP.Slides.Item(2).Shapes.Item('native-hierarchy-after').SmartArt
 Assert-That (($after|ConvertTo-Json -Compress) -eq ($reopened|ConvertTo-Json -Compress)) 'Reopen altered node text or levels'
 $proof.reopenedNodes=$reopened
 for($i=1;$i -le 3;$i++){$taskP.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($RenderDirectory)) ('slide-'+$i+'.png')),'PNG',1600,900)}
 $proof|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $RenderDirectory 'com-audit.json') -Encoding UTF8
 Write-Output 'Native SmartArt operations exported for independent inspection.'
}finally{if($taskP){$taskP.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($taskP)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp)|Out-Null}
