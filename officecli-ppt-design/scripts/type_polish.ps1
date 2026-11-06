param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Choose a new evidence directory; overwrite is disabled'}
[void][IO.Directory]::CreateDirectory($Evidence)
if(Test-Path -LiteralPath $OutputDeck){throw 'Choose a new output file; overwrite is disabled'}
$a=New-Object -ComObject PowerPoint.Application;$p=$null
function Rgb([string]$Hex){return [int]::Parse($Hex.Substring(0,2),'HexNumber')+256*[int]::Parse($Hex.Substring(2,2),'HexNumber')+65536*[int]::Parse($Hex.Substring(4,2),'HexNumber')}
function Solid($shape,[string]$Hex){$shape.Fill.Visible=-1;$shape.Fill.Solid();$shape.Fill.ForeColor.RGB=(Rgb $Hex);$shape.Line.Visible=0}
function Copy-Shape($shape,[string]$Name){$q=$shape.Duplicate().Item(1);$q.Name=$Name;$q.Left=[single]$shape.Left;$q.Top=[single]$shape.Top;return $q}
function Merge-Native($slide,[string[]]$Names,[int]$Cmd,[string]$Primary,[string]$Name){
 $ids=@{};for($i=1;$i -le $slide.Shapes.Count;$i++){$ids[[int]$slide.Shapes.Item($i).Id]=$true}
 $primaryShape=$slide.Shapes.Item($Primary)
 $slide.Shapes.Range([object[]]$Names).MergeShapes($Cmd,$primaryShape)
 $result=$null;for($i=1;$i -le $slide.Shapes.Count;$i++){$q=$slide.Shapes.Item($i);if(-not $ids.ContainsKey([int]$q.Id)){$result=$q}}
 if($null -eq $result){throw "No new native geometry after merge: $Name"}
 $result.Name=$Name;return $result
}
try {
 $p=$a.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$s=$p.Slides.Item(1);$s.Select()
 $s.Export((Join-Path $Evidence 'before.png'),'PNG',1600,900)
 $seed=$s.Shapes.Item('word-seed')
 $cover=$s.Shapes.AddShape(1,0,0,$p.PageSetup.SlideWidth,$p.PageSetup.SlideHeight);$cover.Name='conversion-cover';Solid $cover '888888';$cover.ZOrder(1)
 $glyph=Merge-Native $s @('word-seed','conversion-cover') 3 'word-seed' 'glyph-outline'
 if($glyph.HasTextFrame -eq -1 -and $glyph.TextFrame.HasText -eq -1){throw 'Text conversion failed: result still contains editable text'}
 Solid $glyph '171620'
 # Three source layers: brown back, white-edged black middle, clean black front.
 $brown=Copy-Shape $glyph 'shadow-outline';$brown.Left=$glyph.Left+8;$brown.Top=$glyph.Top+2;Solid $brown 'AA7F5F';$brown.ZOrder(1)
 $shadow=Copy-Shape $brown 'shadow-clip-base'
 $middle=Copy-Shape $glyph 'white-edge';Solid $middle '171620';$middle.Line.Visible=-1;$middle.Line.ForeColor.RGB=(Rgb 'FFFFFF');$middle.Line.Weight=4.5
 $front=Copy-Shape $glyph 'black-front';Solid $front '171620'
 # Highlight = brown glyph minus a right-shifted duplicate; actual Boolean subtraction.
 $hi=Copy-Shape $glyph 'highlight-positive';Solid $hi 'AA7F5F'
 $cut=Copy-Shape $glyph 'highlight-cutter';$cut.Left=$cut.Left+3
 $highlight=Merge-Native $s @('highlight-positive','highlight-cutter') 4 'highlight-positive' 'highlight-subtract';Solid $highlight 'AA7F5F';$highlight.Left=$highlight.Left+1.4
 # Thin rotated rectangles form a native union, then intersect with the shadow glyph.
 $stripeNames=@();$sw=$p.PageSetup.SlideWidth;$sh=$p.PageSetup.SlideHeight
 for($i=-20;$i -le 170;$i++){
  $q=$s.Shapes.AddShape(1,-$sw*.3,$i*4,$sw*1.6,1.1);$q.Name="stripe-$i";Solid $q 'AA7F5F';$q.Rotation=-12;$stripeNames+=$q.Name
 }
 $stripes=Merge-Native $s $stripeNames 1 $stripeNames[0] 'stripe-union'
 $clipped=Merge-Native $s @('shadow-clip-base','stripe-union') 3 'shadow-clip-base' 'stripe-shadow-intersect';Solid $clipped 'AA7F5F'
 $brown.Fill.Visible=0;$brown.Line.Visible=-1;$brown.Line.ForeColor.RGB=(Rgb 'AA7F5F');$brown.Line.Weight=.7
 # Establish final stacking explicitly; helpers do not depend on duplicate z-order.
 $glyph.Delete();$clipped.ZOrder(1);$brown.ZOrder(0);$middle.ZOrder(0);$front.ZOrder(0);$highlight.ZOrder(0)
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 $s.Export((Join-Path $Evidence 'after.png'),'PNG',1600,900)
 $items=@();for($i=1;$i -le $s.Shapes.Count;$i++){$q=$s.Shapes.Item($i);$items+=@{name=$q.Name;type=[int]$q.Type;nodes=try{$q.Nodes.Count}catch{0};hasText=[int]$q.HasTextFrame}}
 @{passed=$true;input=$InputDeck;output=$OutputDeck;font='Script MT Bold substitutes absent Milkshake';shapes=$items;operations=@('intersect-text-and-cover','subtract-highlight','union-stripes','intersect-stripes-and-shadow')}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $Evidence 'polish-result.json') -Encoding utf8
}finally{if($p){$p.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)}
