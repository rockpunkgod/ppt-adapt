param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Evidence)
$ErrorActionPreference='Stop'
$InputDeck=[IO.Path]::GetFullPath($InputDeck);$OutputDeck=[IO.Path]::GetFullPath($OutputDeck);$Evidence=[IO.Path]::GetFullPath($Evidence)
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'Fresh outputs required'}
[void][IO.Directory]::CreateDirectory($Evidence);[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($OutputDeck))
function Solid($q,[int]$rgb){$q.Fill.Visible=-1;$q.Fill.Solid();$q.Fill.ForeColor.RGB=$rgb;$q.Line.Visible=0}
function Merge($s,[string[]]$names,[int]$mode,[string]$primary,[string]$name){
 $ids=@{};for($k=1;$k -le $s.Shapes.Count;$k++){$ids[[int]$s.Shapes.Item($k).Id]=$true}
 $primaryShape=$s.Shapes.Item($primary)
 $range=$s.Shapes.Range([object[]]$names)
 $range.Select()
 $range.MergeShapes($mode,$primaryShape)
 $r=$null;for($k=1;$k -le $s.Shapes.Count;$k++){$q=$s.Shapes.Item($k);if(-not $ids.ContainsKey([int]$q.Id)){$r=$q}}
 if(-not $r){throw 'No new geometry after Boolean'};$r.Name=$name;return $r
}
$app=New-Object -ComObject PowerPoint.Application;$pres=$null
try{
 $app.Visible=-1;$pres=$app.Presentations.Open($InputDeck,0,0,-1)
 $pres.Windows.Item(1).Activate();$pres.Windows.Item(1).ViewType=9
 $s=$pres.Slides.Item(1);$s.Select();$seed=$s.Shapes.Item('number-seed')
 $s.Export((Join-Path $Evidence 'before-1.png'),'PNG',1600,900)
 $copy=$seed.Duplicate().Item(1);$copy.Name='number-copy';$copy.Left=[single]$seed.Left;$copy.Top=[single]$seed.Top
 $clip=$s.Shapes.AddShape(1,0,[single](12.1*72/2.54),$pres.PageSetup.SlideWidth,$pres.PageSetup.SlideHeight);$clip.Name='lower-clip'
 $part=Merge $s @('number-copy','lower-clip') 3 'number-copy' 'partial-gold-glyph';Solid $part 1667493
 $s.Export((Join-Path $Evidence 'slide-1.png'),'PNG',1600,900)
 $s=$pres.Slides.Item(2);$s.Select()
 $s.Export((Join-Path $Evidence 'before-2.png'),'PNG',1600,900)
 $cover=$s.Shapes.AddShape(1,0,0,$pres.PageSetup.SlideWidth,$pres.PageSetup.SlideHeight);$cover.Name='outline-cover'
 $glyph=Merge $s @('number-seed','outline-cover') 3 'number-seed' 'full-glyph';Solid $glyph 1667493
 $right=[double]($glyph.Left+$glyph.Width);$bottom=[double]($glyph.Top+$glyph.Height)
 $side=[Math]::Min([double]$glyph.Width*0.30,[double]$glyph.Height*0.40)
 $tri=$s.Shapes.AddShape(8,[single]($right-$side),[single]($bottom-$side),[single]($side+1),[single]($side+1));$tri.Name='corner-cut';$tri.Flip(0)
 $cut=Merge $s @('full-glyph','corner-cut') 4 'full-glyph' 'corner-subtracted-glyph';Solid $cut 1667493
 $line=$s.Shapes.AddLine([single]($right-$side*1.10),[single]($bottom+$side*0.10),[single]($right+$side*0.10),[single]($bottom-$side*1.10));$line.Name='independent-slash';$line.Line.ForeColor.RGB=1667493;$line.Line.Weight=1.2
 $s.Export((Join-Path $Evidence 'slide-2.png'),'PNG',1600,900)
 $pres.SaveAs($OutputDeck,24)
 @{boundaryCm=12.1;cornerRightPt=$right;cornerBottomPt=$bottom;cutSidePt=$side;operations=@('duplicate-reset-position','intersect-lower-glyph','intersect-full-glyph','subtract-real-triangle')}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'operations.json') -Encoding utf8
}finally{if($pres){$pres.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
