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
  for($i=1;$i -le 4;$i++){
   $ids=@{};for($k=1;$k -le $s.Shapes.Count;$k++){$ids[[int]$s.Shapes.Item($k).Id]=$true}
   $primary=$s.Shapes.Item("card-$i")
   $s.Shapes.Range([object[]]@("card-$i","digit-$i")).MergeShapes(1,$primary)
   $merged=$null
   for($k=1;$k -le $s.Shapes.Count;$k++){$q=$s.Shapes.Item($k);if(-not $ids.ContainsKey([int]$q.Id)){$merged=$q}}
   if(-not $merged){throw 'No native union result'}
   $merged.Name="union-$i"
   $merged.Line.Visible=0
   $merged.Shadow.Visible=-1;$merged.Shadow.Transparency=.65;$merged.Shadow.Blur=6;$merged.Shadow.OffsetX=-3;$merged.Shadow.OffsetY=1
   $g=$s.Shapes.Range([object[]]@("union-$i","title-$i","body-$i")).Group();$g.Name="!!true-card-$i"
   if($j -eq 1){$g.Left=810+($i-1)*9}
   $g.ZOrder(0)
  }
 }
 $p.SaveAs($OutputDeck,24)
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)}
