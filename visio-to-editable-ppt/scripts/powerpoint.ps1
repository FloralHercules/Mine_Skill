param(
 [Parameter(Mandatory=$true)][ValidateSet('template','verify')][string]$Mode,
 [Parameter(Mandatory=$true)][string]$WorkDir
)
$ErrorActionPreference='Stop'
$work=[IO.Path]::GetFullPath($WorkDir)
$manifest=Get-Content -LiteralPath (Join-Path $work 'manifest.json') -Encoding UTF8 -Raw | ConvertFrom-Json
$app=$null; $deck=$null
try {
 # A prior automation server can still be shutting down after Quit returns.
 # Retry only startup RPC-disconnection errors, before modifying slide content.
 $openTime=0.0
 for($attempt=1;$attempt -le 6;$attempt++){
  try{
   $app=New-Object -ComObject PowerPoint.Application
   $timer=[Diagnostics.Stopwatch]::StartNew()
   if($Mode -eq 'template'){$deck=$app.Presentations.Add(0)}
   else{$deck=$app.Presentations.Open((Join-Path $work 'candidate.pptx'),0,0,0)}
   $openTime=$timer.Elapsed.TotalSeconds
   break
  }catch{
   $retryable=$false
   for($cause=$_.Exception;$null -ne $cause;$cause=$cause.InnerException){
    if($cause.HResult -in @(-2147417848,-2147023174)){$retryable=$true}
   }
   if(-not $retryable -or $attempt -eq 6){throw}
   if($null -ne $app -and [Runtime.InteropServices.Marshal]::IsComObject($app)){
    try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}catch{}
   }
   $app=$null;$deck=$null
   Start-Sleep -Milliseconds 500
  }
 }
 if($Mode -eq 'template'){
  $deck.PageSetup.SlideWidth=[single]$manifest.slide_width_pt
  $deck.PageSetup.SlideHeight=[single]$manifest.slide_height_pt
  $null=$deck.Slides.Add(1,12)
  $deck.SaveAs((Join-Path $work 'base.pptx'),24)
 }else{
  if($deck.Slides.Count -ne 1){throw 'Expected exactly one slide'}
  $timer.Restart()
  $previewWidth=1800
  $previewHeight=[int][Math]::Round($previewWidth*$manifest.slide_height_pt/$manifest.slide_width_pt)
  $deck.Slides.Item(1).Export((Join-Path $work 'preview.png'),'PNG',$previewWidth,$previewHeight)
  $renderTime=$timer.Elapsed.TotalSeconds
  $ole=0
  foreach($s in $deck.Slides.Item(1).Shapes){if($s.Type -eq 7){$ole++}}
  $expectedOle=@($manifest.objects | Where-Object { $_.ole_binary }).Count
  if($ole -ne $expectedOle){throw "OLE count mismatch: expected $expectedOle, got $ole"}
  $deck.SaveAs((Join-Path $work 'verified.pptx'),24)
  $report=[ordered]@{slides=1;ole_count=$ole;open_seconds=$openTime;render_seconds=$renderTime;top_level_objects=$deck.Slides.Item(1).Shapes.Count}
  $report | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $work 'powerpoint-check.json') -Encoding UTF8
  $report | ConvertTo-Json
 }
}finally{
 if($null -ne $deck){try{$deck.Close()}catch{}}
 # PowerPoint can reuse a running automation server. Never close other decks.
 if($null -ne $app){
  try{if($app.Presentations.Count -eq 0){$app.Quit()}}catch{}
 }
 foreach($value in @($deck,$app)){
  if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){
   try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}catch{}
  }
 }
}
