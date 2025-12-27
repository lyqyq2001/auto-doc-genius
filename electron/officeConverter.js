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

async function convertBatchWordToPdf(inputOutputPairs) {
  return new Promise(resolve => {
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

$app = $null
$appName = ""
$results = @()

function ReleaseComObject($obj) {
    if ($obj) {
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null
        } catch {}
        $obj = $null
    }
}

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
                    throw "Could not launch Microsoft Word or WPS Writer. Please ensure one of them is installed."
                }
            }
        }
    }

    if (-not $app) { throw "Application object is null" }

    $app.Visible = $false
    $app.DisplayAlerts = 0
    $app.ScreenUpdating = $false
    $app.EnableCancelKey = 0

    $total = $inputOutputPairs.Count
    $successCount = 0

    foreach ($pair in $inputOutputPairs) {
        $doc = $null
        try {
            if (-not (Test-Path -LiteralPath $pair.Input)) {
                $results += @{"Success"=$false;"Input"=$pair.Input;"Error"="File not found"}
                continue
            }

            $doc = $app.Documents.Open($pair.Input, $false, $true, $false)

            if (-not $doc) { 
                $results += @{"Success"=$false;"Input"=$pair.Input;"Error"="Failed to open document"}
                continue
            }

            $doc.ExportAsFixedFormat($pair.Output, 17)
            $doc.Close(0)
            ReleaseComObject $doc
            $doc = $null

            if (Test-Path -LiteralPath $pair.Output) {
                $results += @{"Success"=$true;"Input"=$pair.Input;"Output"=$pair.Output}
                $successCount++
            } else {
                $results += @{"Success"=$false;"Input"=$pair.Input;"Error"="PDF file not generated"}
            }

        } catch {
            $errorMsg = $_.Exception.Message
            $results += @{"Success"=$false;"Input"=$pair.Input;"Error"=$errorMsg}
            if ($doc) { 
                try { $doc.Close(0) } catch {}
                ReleaseComObject $doc
                $doc = $null
            }
        }
    }

} catch {
    if ($doc) { 
        try { $doc.Close(0) } catch {}
        ReleaseComObject $doc
    }
} finally {
    if ($app) { 
        try {
            $app.Quit()
            ReleaseComObject $app
        } catch {}
        $app = $null
    }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

if ($successCount -eq $inputOutputPairs.Count) {
    Write-Host "BATCH_CONVERSION_SUCCESS"
} else {
    Write-Host "BATCH_CONVERSION_PARTIAL_SUCCESS"
}
$results | ConvertTo-Json -Compress
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
          if (stderr) {
            console.error(`[Office 批量转换] PowerShell错误: ${stderr}`);
          }

          if (error) {
            console.error(`[Office 批量转换] 执行失败: ${error.message}`);

            if (fs.existsSync(tempDir)) {
              fs.rmSync(tempDir, { recursive: true, force: true });
            }
            resolve({ success: false, error: error.message });

            return;
          }

          const isSuccess =
            stdout.includes('BATCH_CONVERSION_SUCCESS') ||
            stdout.includes('BATCH_CONVERSION_PARTIAL_SUCCESS');

          if (isSuccess) {
            const pdfResults = [];

            for (let i = 0; i < inputOutputPairs.length; i++) {
              const pdfPath = inputOutputPairs[i].output;

              if (fs.existsSync(pdfPath)) {
                const pdfBuffer = fs.readFileSync(pdfPath);
                pdfResults.push({
                  name: path.basename(pdfPath),
                  data: pdfBuffer,
                });
              } else {
                console.error(`[Office 批量转换] PDF文件不存在: ${pdfPath}`);
              }
            }

            if (fs.existsSync(tempDir)) {
              fs.rmSync(tempDir, { recursive: true, force: true });
            }

            console.log(
              `[Office 批量转换] 完成，成功转换 ${pdfResults.length}/${inputOutputPairs.length} 个文件`
            );
            resolve({ success: true, results: pdfResults });
          } else {
            console.error(`[Office 批量转换] 转换失败`);

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
        try {
          fs.rmSync(tempDir, {
            recursive: true,
            force: true,
          });
        } catch (cleanupError) {}
      }

      resolve({ success: false, error: error.message });
    }
  });
}

module.exports = {
  checkWordInstallation,
  convertBatchWordToPdf,
};
