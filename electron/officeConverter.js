const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

async function checkWordInstallation() {
  // 简单检查：只要能调起 Word 或 WPS 就算已安装
  return new Promise(resolve => {
    const checkCmd = `
        try { 
            $a = New-Object -ComObject Word.Application; $a.Quit() 
        } catch { 
            $b = New-Object -ComObject Kwps.Application; $b.Quit() 
        }
        `;
    exec(
      `powershell -Command "${checkCmd}"`,
      {
        stdio: 'ignore',
        timeout: 10000,
      },
      error => {
        if (error) {
          resolve(false);
        } else {
          resolve(true);
        }
      }
    );
  });
}

async function convertBatchWordToPdf(inputOutputPairs, progressCallback) {
  return new Promise((resolve, reject) => {
    let tempScriptPath = null;
    let tempDir = null;

    try {
      if (!inputOutputPairs || inputOutputPairs.length === 0) {
        resolve({ success: true, results: [] });
        return;
      }

      tempDir = path.join(os.tmpdir(), 'office_batch_convert_' + Date.now());
      if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir, { recursive: true });
      }

      inputOutputPairs.forEach(pair => {
        const outputDir = path.dirname(pair.output);
        if (!fs.existsSync(outputDir)) {
          fs.mkdirSync(outputDir, { recursive: true });
        }
      });

      const psContent = `
$ErrorActionPreference = "Stop"
$inputOutputPairs = @(
${inputOutputPairs
  .map(p => `    @{Input="${p.input}"; Output="${p.output}"}`)
  .join('\n')}
)

Write-Host "Input/Output pairs:"
$inputOutputPairs | ForEach-Object { Write-Host "  Input: $($_.Input), Output: $($_.Output)" }

$app = $null
$appName = ""
$results = @()

try {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
        $appName = "MS Word (Active)"
    } catch {
        try {
            $app = New-Object -ComObject Word.Application
            $appName = "MS Word (New)"
        } catch {
            try {
                $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Kwps.Application")
                $appName = "WPS (Active)"
            } catch {
                try {
                    $app = New-Object -ComObject Kwps.Application
                    $appName = "WPS (New)"
                } catch {
                    throw "Could not launch Microsoft Word or WPS Writer."
                }
            }
        }
    }

    Write-Host "Using application: $appName"

    if (-not $app) { throw "Application object is null." }

    $app.Visible = $false
    $app.DisplayAlerts = 0
    $app.ScreenUpdating = $false
    $app.EnableCancelKey = 0

    $total = $inputOutputPairs.Count
    $current = 0
    $successCount = 0

    foreach ($pair in $inputOutputPairs) {
        $current++
        $percent = [math]::Round(($current / $total) * 100)
        Write-Host "Processing [$current/$total] ($percent%): $($pair.Input)"

        $doc = $null
        try {
            if (-not (Test-Path -LiteralPath $pair.Input)) {
                Write-Host "ERROR: Input file not found: $($pair.Input)"
                $results += @{"Success"=$false;"Input"=$pair.Input;"Error"="File not found"}
                continue
            }

            Write-Host "Opening document: $($pair.Input)"
            $doc = $app.Documents.Open($pair.Input, $false, $true, $false)

            if (-not $doc) { 
                Write-Host "ERROR: Failed to open document"
                $results += @{"Success"=$false;"Input"=$pair.Input;"Error"="Failed to open document"}
                continue
            }

            Write-Host "Exporting to PDF: $($pair.Output)"
            $doc.ExportAsFixedFormat($pair.Output, 17)
            $doc.Close(0)

            if (Test-Path -LiteralPath $pair.Output) {
                Write-Host "SUCCESS: PDF created: $($pair.Output)"
                $results += @{"Success"=$true;"Input"=$pair.Input;"Output"=$pair.Output}
                $successCount++
            } else {
                Write-Host "ERROR: Output file not created: $($pair.Output)"
                $results += @{"Success"=$false;"Input"=$pair.Input;"Error"="Output file not created"}
            }

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            $doc = $null
        } catch {
            $errorMsg = $_.Exception.Message
            Write-Host "ERROR: Exception occurred: $errorMsg"
            Write-Host $_.ScriptStackTrace
            $results += @{"Success"=$false;"Input"=$pair.Input;"Error"=$errorMsg}
            if ($doc) { 
                try { $doc.Close(0) } catch {}
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
                $doc = $null
            }
        }
    }

    Write-Host "Conversion completed: $successCount/$total succeeded"

    $app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
    $app = $null

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "BATCH_CONVERSION_SUCCESS"
    $results | ConvertTo-Json -Compress
} catch {
    Write-Host "ERROR_OCCURRED"
    Write-Host $_.Exception.Message
    Write-Host $_.ScriptStackTrace
    if ($doc) { 
        try { $doc.Close(0) } catch {}
    }
    if ($app) { 
        try { $app.Quit() } catch {}
    }
    exit 1
}
`;

      tempScriptPath = path.join(tempDir, 'batch_convert.ps1');
      fs.writeFileSync(tempScriptPath, psContent, { encoding: 'utf8' });

      const cmd = `powershell -NoProfile -ExecutionPolicy Bypass -File "${tempScriptPath}"`;

      exec(
        cmd,
        {
          encoding: 'utf8',
          timeout: 300000,
          windowsHide: true,
        },
        (error, stdout, stderr) => {
          console.log(`[Office 批量转换] PowerShell输出:\n${stdout}`);
          if (stderr) {
            console.log(`[Office 批量转换] PowerShell错误:\n${stderr}`);
          }

          if (error) {
            console.error(`[Office 批量转换] 失败: ${error.message}`);
            if (fs.existsSync(tempDir)) {
              fs.rmSync(tempDir, { recursive: true, force: true });
            }
            resolve({ success: false, error: error.message });
            return;
          }

          if (stdout.includes('BATCH_CONVERSION_SUCCESS')) {
            console.log(`[Office 批量转换] 成功，开始读取PDF文件`);

            const pdfResults = [];

            for (const pair of inputOutputPairs) {
              const pdfPath = pair.output;
              console.log(`[Office 批量转换] 检查PDF文件: ${pdfPath}`);
              if (fs.existsSync(pdfPath)) {
                const pdfBuffer = fs.readFileSync(pdfPath);
                pdfResults.push({
                  name: path.basename(pdfPath),
                  data: pdfBuffer,
                });
                console.log(
                  `[Office 批量转换] 成功读取PDF: ${pdfPath} (${pdfBuffer.length} bytes)`
                );
              } else {
                console.warn(`[Office 批量转换] PDF文件不存在: ${pdfPath}`);
                console.log(`[Office 批量转换] 临时目录内容:`);
                const files = fs.readdirSync(tempDir);
                files.forEach(file => {
                  const filePath = path.join(tempDir, file);
                  const stats = fs.statSync(filePath);
                  console.log(`  ${file} (${stats.size} bytes)`);
                });
              }
            }

            if (fs.existsSync(tempDir)) {
              fs.rmSync(tempDir, { recursive: true, force: true });
            }

            resolve({ success: true, results: pdfResults });
          } else {
            console.error(`[Office 批量转换] 输出:\n${stdout}`);
            if (fs.existsSync(tempDir)) {
              fs.rmSync(tempDir, { recursive: true, force: true });
            }
            resolve({ success: false, error: '转换失败' });
          }
        }
      );
    } catch (error) {
      console.error(`[Office 批量转换] 失败: ${error.message}`);
      if (tempDir && fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, {
          recursive: true,
          force: true,
        });
      }
      resolve({ success: false, error: error.message });
    }
  });
}

module.exports = {
  checkWordInstallation,
  convertBatchWordToPdf,
};
