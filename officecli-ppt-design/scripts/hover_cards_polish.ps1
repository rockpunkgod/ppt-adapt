param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$Evidence)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $Evidence)){throw 'Fresh outputs required'}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($Evidence))
$app=New-Object -ComObject PowerPoint.Application;$pres=$null;$log=@();$bounds=@()
try{
 $app.Visible=-1;$pres=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 if($pres.Slides.Count -ne 5){throw 'Expected five fresh foundation states'}
 for($i=1;$i -le 5;$i++){
  $s=$pres.Slides.Item($i);$s.SlideShowTransition.AdvanceOnClick=0;$s.SlideShowTransition.AdvanceOnTime=0
  for($j=1;$j -le 4;$j++){
   $q=$s.Shapes.Item("hover-card-$j");$q.ActionSettings.Item(1).Action=0
   $hover=$q.ActionSettings.Item(2)
   if($i -eq $j+1){$hover.Action=0;continue}
   $target=$pres.Slides.Item($j+1);$hover.Action=7;$hover.Hyperlink.Address='';$hover.Hyperlink.SubAddress=([string]$target.SlideID)+','+([string]$target.SlideIndex)+',Team state'
   $hover.AnimateAction=0
   $log+=@{slide=$i;shape=$q.Name;target=$j+1;action=[int]$hover.Action;subAddress=$hover.Hyperlink.SubAddress}
  }
  if($i -gt 1){
   $bio=$s.Shapes.Item("bio-$($i-1)");$range=$bio.TextFrame2.TextRange
   $bottom=[double]$range.BoundTop+[double]$range.BoundHeight
   $allowed=[double]$bio.Top+[double]$bio.Height-[double]$bio.TextFrame2.MarginBottom
   if($bottom -gt $allowed+1){throw "Profile text overflows state $i"}
   $bounds+=@{slide=$i;bottom=$bottom;allowedBottom=$allowed;passed=($bottom -le $allowed+1)}
  }
  $s.Export((Join-Path $Evidence "slide-$i.png"),'PNG',1600,900)
 }
 $pres.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 $log|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $Evidence 'actions.json') -Encoding utf8
 $bounds|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $Evidence 'text-bounds.json') -Encoding utf8
}finally{if($pres){$pres.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
