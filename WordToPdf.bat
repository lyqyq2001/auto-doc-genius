@echo off
chcp 65001 > nul
echo 开始批量转换docx到PDF...
echo.

:: 调用PowerShell脚本（核心逻辑和方法一一致）
powershell -Command "$word = New-Object -ComObject Word.Application; $word.Visible = $false; $word.DisplayAlerts = 0; $docxFiles = Get-ChildItem -Path . -Filter *.docx -File; if ($docxFiles.Count -eq 0) { Write-Host '未找到docx文件！' -ForegroundColor Red; $word.Quit(); exit; } foreach ($file in $docxFiles) { try { $pdfPath = $file.FullName -replace '\.docx$', '.pdf'; $doc = $word.Documents.Open($file.FullName); $doc.SaveAs2($pdfPath, 17); $doc.Close(); Write-Host '转换成功：' $file.Name -ForegroundColor Green; } catch { Write-Host '转换失败：' $file.Name '，错误：' $_ -ForegroundColor Red; if ($doc) { $doc.Close(); } } } $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word); [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers(); Write-Host '批量转换完成！' -ForegroundColor Cyan;"

echo.
echo 转换结束，按任意键退出...
pause > nul