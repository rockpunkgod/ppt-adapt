param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $EvidenceDirectory)){throw 'Refusing existing outputs'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 $p.SectionProperties.AddBeforeSlide(1,'Studio')|Out-Null
 $p.SectionProperties.AddBeforeSlide(2,'Field Notes')|Out-Null
 $p.SectionProperties.AddBeforeSlide(4,'Design System')|Out-Null
 $sections=@();for($i=1;$i -le $p.SectionProperties.Count;$i++){$sections+=@{index=$i;id=$p.SectionProperties.SectionID($i);first=$p.SectionProperties.FirstSlide($i);count=$p.SectionProperties.SlidesCount($i);name=$p.SectionProperties.Name($i)}}
 foreach($i in @(2,4)){$p.Slides.Item($i).Export((Join-Path $EvidenceDirectory "entry-$i.png"),'PNG',1600,900)}
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'sections.json'),($sections|ConvertTo-Json -Depth 5),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close();[Runtime.InteropServices.Marshal]::ReleaseComObject($p)|Out-Null};[Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null}
