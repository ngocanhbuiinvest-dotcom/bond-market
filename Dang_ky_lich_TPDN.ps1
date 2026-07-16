# Đăng ký Windows Task Scheduler chạy cập nhật TPDN hàng ngày.
# Chạy MỘT LẦN (PowerShell thường, KHÔNG cần Administrator vì task chạy dưới user hiện tại):
#     powershell -ExecutionPolicy Bypass -File ".\Dang_ky_lich_TPDN.ps1"
# Gỡ lịch:  Unregister-ScheduledTask -TaskName "TPDN_CapNhatHangNgay" -Confirm:$false

$dir      = "E:\NGOC ANH_BACK UP 20220921\6. AI\CLAUDE\3. Bond Market"
$bat      = Join-Path $dir "run_update.bat"
$taskName = "TPDN_CapNhatHangNgay"
$runAt    = "16:30"   # 16:30 các NGÀY LÀM VIỆC (T2-T6); đổi nếu muốn

if (-not (Test-Path $bat)) { Write-Error "Không thấy $bat"; exit 1 }

$action   = New-ScheduledTaskAction -Execute $bat -WorkingDirectory $dir
$trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At $runAt
# StartWhenAvailable: nếu máy tắt lúc 16:30 thì chạy bù khi bật lại (không bỏ lỡ ngày)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
              -MultipleInstances IgnoreNew
$desc     = "Cập nhật dữ liệu TPDN riêng lẻ (HNX) 16:30 ngày làm việc: full re-scrape nguồn hay đổi + incremental GD thứ cấp + nhật ký thay đổi + cảnh báo mã chậm trả mới + rebuild dashboard."

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
    -Settings $settings -Description $desc -Force | Out-Null

Write-Host "Đã đăng ký task '$taskName' chạy 16:30 các ngày làm việc (T2-T6)."
Write-Host "Chạy thử ngay:  Start-ScheduledTask -TaskName '$taskName'"
Write-Host "Xem log:        Get-Content '$dir\update_log.txt' -Tail 40"
