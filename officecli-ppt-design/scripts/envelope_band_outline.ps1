param([string]$InputDeck,[string]$OutputDeck,[string]$Evidence)
$ErrorActionPreference='Stop';if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'Fresh outputs required'};[void][IO.Directory]::CreateDirectory($Evidence)
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $app.Visible=-1;$p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1);$s=$p.Slides.Item(2);$s.Select();$seed=$s.Shapes.Item('glyph-seed');$cover=$s.Shapes.AddShape(1,0,0,$p.PageSetup.SlideWidth,$p.PageSetup.SlideHeight);$cover.Name='outline-cover';$ids=@{};for($i=1;$i -le $s.Shapes.Count;$i++){$ids[[int]$s.Shapes.Item($i).Id]=$true}
 $s.Shapes.Range([object[]]@('glyph-seed','outline-cover')).MergeShapes(3,$seed);$glyph=$null;for($i=1;$i -le $s.Shapes.Count;$i++){$q=$s.Shapes.Item($i);if(-not $ids.ContainsKey([int]$q.Id)){$glyph=$q}}
 if(-not $glyph){throw 'Native outline conversion produced no shape'};$glyph.Name='source-glyph';$glyph.Line.Visible=0
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24);@{nodes=$glyph.Nodes.Count;type=[int]$glyph.Type;name=$glyph.Name;font='Arial Bold';textBackup=$p.Slides.Item(1).Shapes.Item('editable-backup').TextFrame.TextRange.Text}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'outline.json') -Encoding UTF8
}finally{if($p){try{$p.Close()}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
