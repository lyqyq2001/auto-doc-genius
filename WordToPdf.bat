@echo off
chcp 65001 > nul
:: 记录开始时间（PowerShell的Get-Date获取时间戳，精确到毫秒）
for /f "tokens=*" %%a in ('powershell -Command "(Get-Date).ToFileTimeUtc()"') do set "startTime=%%a"

echo 开始批量转换docx到PDF...
echo 开始时间：%date% %time%
echo.

:: 调用PowerShell脚本执行转换，同时传递开始时间用于计算耗时
powershell -Command "$startTime = [DateTime]::FromFileTimeUtc(%startTime%); $word = New-Object -ComObject Word.Application; $word.Visible = $false; $word.DisplayAlerts = 0; $docxFiles = Get-ChildItem -Path . -Filter *.docx -File; if ($docxFiles.Count -eq 0) { Write-Host '未找到docx文件！' -ForegroundColor Red; $word.Quit(); exit; } $successCount = 0; $failCount = 0; foreach ($file in $docxFiles) { try { $pdfPath = $file.FullName -replace '\.docx$', '.pdf'; $doc = $word.Documents.Open($file.FullName); $doc.SaveAs2($pdfPath, 17); $doc.Close(); $successCount++; Write-Host '转换成功：' $file.Name -ForegroundColor Green; } catch { $failCount++; Write-Host '转换失败：' $file.Name '，错误：' $_ -ForegroundColor Red; if ($doc) { $doc.Close(); } } } $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word); [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers(); $endTime = Get-Date; $duration = $endTime - $startTime; Write-Host ''; Write-Host '========== 转换统计 ==========' -ForegroundColor Cyan; Write-Host '总文件数：' $docxFiles.Count -ForegroundColor White; Write-Host '成功数：' $successCount -ForegroundColor Green; Write-Host '失败数：' $failCount -ForegroundColor Red; Write-Host '开始时间：' $startTime.ToString('yyyy-MM-dd HH:mm:ss.fff') -ForegroundColor White; Write-Host '结束时间：' $endTime.ToString('yyyy-MM-dd HH:mm:ss.fff') -ForegroundColor White; Write-Host '总耗时：' $duration.Hours '小时' $duration.Minutes '分钟' $duration.Seconds '秒' $duration.Milliseconds '毫秒' -ForegroundColor Yellow;"

echo.
echo 转换结束，按任意键退出...
pause > nul