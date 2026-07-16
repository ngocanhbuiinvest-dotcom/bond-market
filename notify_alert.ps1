# Hiển thị cảnh báo desktop (balloon khay hệ thống) — dùng cho "mã chậm trả mới".
# Gọi: powershell -ExecutionPolicy Bypass -File notify_alert.ps1 -Title "..." -Message "..."
# Không cần cài module ngoài (dùng System.Windows.Forms.NotifyIcon). Best-effort: nếu phiên không
# tương tác (task chạy nền) có thể không hiện — khi đó vẫn còn banner dashboard + file alert.
param(
  [string]$Title   = "Cảnh báo TPDN",
  [string]$Message = ""
)
try {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  $ni = New-Object System.Windows.Forms.NotifyIcon
  $ni.Icon = [System.Drawing.SystemIcons]::Warning
  $ni.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]::Warning
  $ni.BalloonTipTitle = $Title
  $ni.BalloonTipText  = $Message
  $ni.Visible = $true
  $ni.ShowBalloonTip(20000)
  Start-Sleep -Seconds 12
  $ni.Dispose()
} catch {
  Write-Host "notify_alert lỗi: $_"
}
