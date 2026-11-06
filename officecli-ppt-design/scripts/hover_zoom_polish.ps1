param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $EvidenceDirectory)){throw 'Fresh outputs required'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$p=$null;$log=@()
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 for($i=1;$i -le 5;$i++){
  $s=$p.Slides.Item($i);$s.SlideShowTransition.AdvanceOnClick=$(if($i -gt 3){-1}else{0});$s.SlideShowTransition.AdvanceOnTime=0
  if($i -le 3){
   $s.SlideShowTransition.Duration=[single]0.3
   foreach($j in @(1,2)){
    $card=$s.Shapes.Item('!!card-'+$j);$hover=$card.ActionSettings.Item(2)
    if($i -eq $j+1){$hover.Action=0}else{$target=$p.Slides.Item($j+1);$hover.Action=7;$hover.Hyperlink.SubAddress=([string]$target.SlideID)+','+([string]$target.SlideIndex)+',Hover preview';$hover.AnimateAction=0}
    # Headers/labels sit above the card; route their hover identically instead of leaving dead regions.
    foreach($kind in @('icon-','label-')){$h=$s.Shapes.Item('!!'+$kind+$j).ActionSettings.Item(2);$h.Action=$hover.Action;if($h.Action -eq 7){$h.Hyperlink.SubAddress=$hover.Hyperlink.SubAddress;$h.AnimateAction=0}}
    $log+=@{slide=$i;card=$j;hoverAction=[int]$hover.Action;target=$(if($hover.Action -eq 7){$j+1}else{$i})}
   }
   $mask=$s.Shapes.Item('!!return-mask');$back=$mask.ActionSettings.Item(2)
   if($i -eq 1){$back.Action=0}else{$back.Action=7;$back.Hyperlink.SubAddress=([string]$p.Slides.Item(1).SlideID)+',1,Home';$back.AnimateAction=0}
   $log+=@{slide=$i;maskReturnAction=[int]$back.Action;maskType=[int]$mask.Type}
  }
  $s.Export((Join-Path $EvidenceDirectory ('slide-'+$i+'.png')),'PNG',1600,900)
 }
 $p.SlideShowSettings.ShowPresenterView=0;$p.SlideShowSettings.RangeType=1;$p.SlideShowSettings.AdvanceMode=1
 $p.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'actions.json'),($log|ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
}finally{if($p){$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
