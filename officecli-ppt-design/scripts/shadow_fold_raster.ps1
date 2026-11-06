param([Parameter(Mandatory=$true)][string]$Deck,[Parameter(Mandatory=$true)][string]$Png)
$ErrorActionPreference='Stop'
$Deck=[IO.Path]::GetFullPath($Deck);$Png=[IO.Path]::GetFullPath($Png)
if(-not(Test-Path -LiteralPath $Deck)){throw 'Input missing'}
if(Test-Path -LiteralPath $Png){throw 'Output exists'}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Png))
$a=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$a.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,-1)
 # ScaleWidth/Height are relative to the whole slide, not the shape's pixels.
 # Use one common scale on both slide axes to preserve ellipse proportions.
 $p.Slides.Item(1).Shapes.Item('soft-ellipse').Export([IO.Path]::GetFullPath($Png),2,[int]($p.PageSetup.SlideWidth*2),[int]($p.PageSetup.SlideHeight*2),1)
 if(-not(Test-Path -LiteralPath $Png)){throw 'Native raster missing'}
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)}
