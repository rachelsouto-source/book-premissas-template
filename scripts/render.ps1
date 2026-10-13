param($pptx,$out)
$ppt = New-Object -ComObject PowerPoint.Application
$p = $ppt.Presentations.Open($pptx, $true, $false, $false)
New-Item -ItemType Directory -Force $out | Out-Null
$i=1; foreach($s in $p.Slides){ $s.Export("$out\s$('{0:D2}' -f $i).png","PNG",1280,720); $i++ }
$p.Close(); $ppt.Quit()
