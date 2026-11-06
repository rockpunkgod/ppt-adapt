param([string]$Deck,[string]$Evidence,[string]$SourceFill='1B7775',[string]$EditedFill='3867A8')
$ErrorActionPreference='Stop'
if($SourceFill -notmatch '^[0-9A-Fa-f]{6}$' -or $EditedFill -notmatch '^[0-9A-Fa-f]{6}$' -or $SourceFill -eq $EditedFill){throw 'Two different six-digit RGB fill colors required'}
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
function RGB([string]$hex){return [Convert]::ToInt32($hex.Substring(0,2),16) -bor ([Convert]::ToInt32($hex.Substring(2,2),16) -shl 8) -bor ([Convert]::ToInt32($hex.Substring(4,2),16) -shl 16)}
function Leaves($shape){if($shape.Type -eq 6){for($j=1;$j -le $shape.GroupItems.Count;$j++){Leaves ($shape.GroupItems.Item($j))}}else{return $shape}}
function Targets($slide){return @($slide.Shapes|Where-Object {$_.Id -notin $script:untouchedIds})}
function Snapshot($slide,[string]$phase){
 $parts=@();foreach($top in @(Targets $slide)){foreach($leaf in @(Leaves $top)){
  $nodes=$null;if($leaf.Type -eq 5){$nodes=[int]$leaf.Nodes.Count}
  $parts+=@{topName=$top.Name;name=$leaf.Name;id=[int]$leaf.Id;type=[int]$leaf.Type;fillType=[int]$leaf.Fill.Type;color=[int]$leaf.Fill.ForeColor.RGB;nodes=$nodes}
 }}
 $script:log+=@{phase=$phase;leaves=$parts};$script:log|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $Evidence 'native-objects.json') -Encoding UTF8
 return $parts
}
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$script:log=@();$script:untouchedIds=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,-1)
 if($p.Slides.Count -ne 1){throw 'One SVG foundation slide required'}
 $s=$p.Slides.Item(1);$s.Select();$target=$s.Shapes.Item('svg-target')
 $script:untouchedIds=@($s.Shapes|Where-Object {$_.Id -ne $target.Id}|ForEach-Object {[int]$_.Id})
 $s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'before.png'),'PNG',1600,900)
 $target.Select(-1)
 if($app.ActivePresentation.FullName -ne $p.FullName){throw 'Own presentation is not active'}
 $enabled=$app.CommandBars.GetEnabledMso('SVGEdit')
 @{command='SVGEdit';enabled=[bool]$enabled;sourceShapeType=[int]$target.Type}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'command.json') -Encoding UTF8
 if(-not $enabled){throw 'Native SVG Convert to Shape command is unavailable for this source picture'}
 $app.CommandBars.ExecuteMso('SVGEdit')
 $converted=@(Targets $s)
 if($converted.Count -eq 0){throw 'Native conversion produced no target geometry'}
 $parts=Snapshot $s 'converted'
 if(@($parts|Where-Object {$_.type -notin @(1,5)}).Count -gt 0){throw 'Converted leaves are still image/unsupported objects'}
 $convertedPath=Join-Path ([IO.Path]::GetFullPath($Evidence)) 'converted.pptx';$p.SaveAs($convertedPath,24)
 $s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'converted.png'),'PNG',1600,900)
 $teal=@();foreach($top in @(Targets $s)){foreach($leaf in @(Leaves $top)){if($leaf.Fill.Type -eq 1 -and $leaf.Fill.ForeColor.RGB -eq (RGB $SourceFill)){$teal+=,$leaf}}}
 if($teal.Count -ne 1){throw 'Expected one editable native leaf matching SourceFill'}
 $teal[0].Fill.Solid();$teal[0].Fill.ForeColor.RGB=RGB $EditedFill
 $editedPath=Join-Path ([IO.Path]::GetFullPath($Evidence)) 'edited.pptx';$p.SaveAs($editedPath,24)
 [void](Snapshot $s 'edited');$s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'edited.png'),'PNG',1600,900)
 $p.Close();$p=$null;$p=$app.Presentations.Open($editedPath,-1,0,-1);$s=$p.Slides.Item(1)
 [void](Snapshot $s 'reopen');$s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'reopen.png'),'PNG',1600,900)
 Write-Output 'Native SVG conversion and actual leaf-fill edit captured; independent geometry/pixel checks required.'
}catch{@{error=$_.Exception.Message;phases=$script:log.Count}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'failure.json') -Encoding UTF8;throw}
finally{if($p){try{$p.Close()}catch{}};[Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null}
