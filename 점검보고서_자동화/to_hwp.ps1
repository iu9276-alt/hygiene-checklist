# HTML 보고서를 한글(HWP)과 PDF로 변환한다. 한글 2020의 HWPFrame.HwpObject 사용.
# 사용: powershell -NoProfile -File to_hwp.ps1 -Html <html> -Hwp <hwp> -Pdf <pdf>
param(
  [Parameter(Mandatory = $true)][string]$Html,
  [Parameter(Mandatory = $true)][string]$Hwp,
  [string]$Pdf = ""
)
$job = Start-Job -ScriptBlock {
  param($Html, $Hwp, $Pdf)
  $h = New-Object -ComObject HWPFrame.HwpObject
  try { $h.XHwpWindows.Item(0).Visible = $false } catch {}
  try {
    if (-not $h.Open($Html, "HTML", "forceopen:true")) { throw "열기 실패: $Html" }
    if (-not $h.SaveAs($Hwp, "HWP", "")) { throw "HWP 저장 실패: $Hwp" }
    if ($Pdf -ne "") { if (-not $h.SaveAs($Pdf, "PDF", "")) { throw "PDF 저장 실패: $Pdf" } }
    "OK"
  } finally { $h.Quit() }
} -ArgumentList $Html, $Hwp, $Pdf

if (Wait-Job $job -Timeout 120) {
  Receive-Job $job
  if ($job.State -eq "Failed") { exit 1 }
} else {
  Stop-Job $job
  Write-Output "TIMEOUT: 한글 창에 확인 대화상자가 떠 있는지 확인하세요."
  exit 2
}
