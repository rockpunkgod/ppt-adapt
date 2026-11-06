param([Parameter(Mandatory=$true)][string]$InputDeck,[Parameter(Mandatory=$true)][string]$OutputDeck,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
if((Test-Path -LiteralPath $OutputDeck) -or (Test-Path -LiteralPath $EvidenceDirectory)){throw 'Refusing existing outputs'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$app=New-Object -ComObject PowerPoint.Application;$pres=$null
try{
 $pres=$app.Presentations.Open([IO.Path]::GetFullPath($InputDeck),0,0,-1)
 $s=$pres.Slides.Item(1);$s.Select();$pin=$s.Shapes.Item('trigger-pin');$label=$s.Shapes.Item('trigger-label')
 $seq=$s.TimeLine.InteractiveSequences.Add(1)
 $entry=$seq.AddTriggerEffect($label,10,4,$pin,'',0);$entry.Timing.Duration=[single]0.5
 $actual=$entry.Parent
 $exit=$s.TimeLine.MainSequence.AddEffect($label,10,0,1);$exit.Exit=-1;$exit.Timing.Duration=[single]0.5;$exit.MoveAfter($entry)
 $s.SlideShowTransition.AdvanceOnClick=0;$s.SlideShowTransition.AdvanceOnTime=0
 $pres.SaveAs([IO.Path]::GetFullPath($OutputDeck),24)
 $s.Export((Join-Path $EvidenceDirectory 'edit-state.png'),'PNG',1600,900)
 $log=@{sequences=$s.TimeLine.InteractiveSequences.Count;effects=$actual.Count;entryTrigger=[int]$entry.Timing.TriggerType;entryTriggerShape=$entry.Timing.TriggerShape.Name;exitTrigger=[int]$exit.Timing.TriggerType;exit=[int]$exit.Exit}
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'readback.json'),($log|ConvertTo-Json -Depth 5),(New-Object Text.UTF8Encoding($false)))
}finally{if($pres){$pres.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
