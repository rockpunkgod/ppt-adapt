param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck)
$ErrorActionPreference='Stop'
$InputDeck=[IO.Path]::GetFullPath($InputDeck)
$OutputDeck=[IO.Path]::GetFullPath($OutputDeck)
if(-not(Test-Path -LiteralPath $InputDeck)){throw 'Input file missing'}
if(Test-Path -LiteralPath $OutputDeck){throw 'Output exists'}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($OutputDeck))
$a=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$a.Presentations.Open($InputDeck,0,0,-1)
 for($j=1;$j -le 2;$j++){
  $s=$p.Slides.Item($j);$s.Select()
  for($k=1;$k -le $s.Shapes.Count;$k++){$q=$s.Shapes.Item($k);$q.Name=$q.Name -replace '^!!',''}
  $s.Shapes.Item('shade').Fill.Transparency=.65
  $text=$s.Shapes.Item('window-seed')
  $cover=$s.Shapes.AddShape(1,0,0,$p.PageSetup.SlideWidth,$p.PageSetup.SlideHeight);$cover.Name='cover'
  $ids=@{};for($k=1;$k -le $s.Shapes.Count;$k++){$ids[[int]$s.Shapes.Item($k).Id]=$true}
  $s.Shapes.Range([object[]]@('window-seed','cover')).MergeShapes(3,$text)
  $g=$null
  for($k=1;$k -le $s.Shapes.Count;$k++){$q=$s.Shapes.Item($k);if(-not $ids.ContainsKey([int]$q.Id)){$g=$q}}
  if(-not $g){throw 'No glyph outline'}
  $g.Name='!!background-window';$g.Line.Visible=0;$g.Fill.Background()
  if($j -eq 2){$g.Left=$g.Left+65;$g.Top=$g.Top+90}
 }
 $p.SaveAs($OutputDeck,24)
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)}
