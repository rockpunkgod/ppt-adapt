param([string]$Deck,[string]$Evidence)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Evidence){throw 'Refusing overwrite'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
$before=(Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash
$app=New-Object -ComObject PowerPoint.Application;$p=$null
try{
 $p=$app.Presentations.Open([IO.Path]::GetFullPath($Deck),0,0,-1)
 if($p.Slides.Count -ne 1){throw 'Expected one original gradient chart'}
 $s=$p.Slides.Item(1);$g=$s.Shapes.Item('styled-series')
 if($g.HasChart -ne -1){throw 'Native chart required'}
 $s.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'gradient.png'),'PNG',1600,900)
 $dup=$s.Duplicate();$patternPage=$dup.Item(1)
 $patternPage.Shapes.Item('heading-1').TextFrame.TextRange.Text='SERIES PATTERN'
 $fill=$patternPage.Shapes.Item('styled-series').Chart.SeriesCollection(1).Format.Fill
 $fill.Patterned(30)
 $fill.ForeColor.RGB=0x65DB78
 $fill.BackColor.RGB=0x382310
 if([int]$fill.Type -ne 2 -or [int]$fill.Pattern -ne 30){throw 'Native pattern readback mismatch'}
 $patternPage.Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'pattern.png'),'PNG',1600,900)
 $p.SaveAs((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'styled.pptx'),24)
 @{pattern=30;fillType=[int]$fill.Type;foreColor=[int]$fill.ForeColor.RGB;backColor=[int]$fill.BackColor.RGB;chartDataApiCalled=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Evidence 'native.json') -Encoding UTF8
 $p.Close();$p=$null
 $p=$app.Presentations.Open((Join-Path ([IO.Path]::GetFullPath($Evidence)) 'styled.pptx'),-1,0,0)
 for($i=1;$i -le 2;$i++){$p.Slides.Item($i).Export((Join-Path ([IO.Path]::GetFullPath($Evidence)) ('reopen-'+$i+'.png')),'PNG',1600,900)}
}finally{if($p){$p.Saved=-1;$p.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
if((Get-FileHash -LiteralPath $Deck -Algorithm SHA256).Hash -ne $before){throw 'Original base changed'}
