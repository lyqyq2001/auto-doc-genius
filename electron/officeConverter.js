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
    const MAX_RETRIES = 2;
    const RETRY_INTERVAL = 1000; // 1秒
    let retryCount = 0;

    // 执行转换的函数
    const executeConversion = () => {
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

Write-Host "File count: $($inputOutputPairs.Count)"
$inputOutputPairs | ForEach-Object { Write-Host "  - $($_.Input) -> $($_.Output)" }

$app = $null
$appName = ""
$results = @()

# Function: Release COM object
function ReleaseComObject($obj) {
    if ($obj) {
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null
        } catch {}
        Remove-Variable obj -Force -ErrorAction SilentlyContinue
    }
}

try {
    # Try to get or create Office app object
    Write-Host "Starting Office application..."
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
        $appName = "MS Word (Active)"
    } catch {
        Write-Host "No active Word instance found, trying to create new instance..."
        try {
            $app = New-Object -ComObject Word.Application
            $appName = "MS Word (New)"
        } catch {
            Write-Host "Word startup failed, trying WPS..."
            try {
                $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Kwps.Application")
                $appName = "WPS (Active)"
            } catch {
                Write-Host "No active WPS instance found, trying to create new instance..."
                try {
                    $app = New-Object -ComObject Kwps.Application
                    $appName = "WPS (New)"
                } catch {
                    throw "Could not launch Microsoft Word or WPS Writer. Please ensure one of them is installed."
                }
            }
        }
    }

    Write-Host "Successfully started: $appName"

    if (-not $app) { throw "Application object is null" }

    # Configure app settings
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
        Write-Host "[$current/$total] ($percent %%) Processing: $($pair.Input)"

        $doc = $null
        try {
            # Check if input file exists
            if (-not (Test-Path -LiteralPath $pair.Input)) {
                Write-Host "ERROR: Input file not found: $($pair.Input)"
                $results += @{"Success"=$false;"Input"=$pair.Input;"Error"="File not found"}
                continue
            }

            # Open document
            Write-Host "  Opening document..."
            $doc = $app.Documents.Open($pair.Input, $false, $true, $false)

            if (-not $doc) { 
                Write-Host "ERROR: Failed to open document"
                $results += @{"Success"=$false;"Input"=$pair.Input;"Error"="Failed to open document"}
                continue
            }

            # Convert to PDF
            Write-Host "  Converting to PDF..."
            $doc.ExportAsFixedFormat($pair.Output, 17)
            
            # Close document
            Write-Host "  Closing document..."
            $doc.Close(0)
            
            # Release document object
            ReleaseComObject $doc
            $doc = $null

            # Check output file
            if (Test-Path -LiteralPath $pair.Output) {
                $fileSize = (Get-Item $pair.Output).Length
                Write-Host "SUCCESS: PDF created, size: $fileSize bytes"
                $results += @{"Success"=$true;"Input"=$pair.Input;"Output"=$pair.Output}
                $successCount++
            } else {
                Write-Host "ERROR: PDF file not generated"
                $results += @{"Success"=$false;"Input"=$pair.Input;"Error"="PDF file not generated"}
            }

        } catch {
            $errorMsg = $_.Exception.Message
            Write-Host "ERROR: Exception occurred: $errorMsg"
            $results += @{"Success"=$false;"Input"=$pair.Input;"Error"=$errorMsg}
            if ($doc) { 
                try { 
                    Write-Host "  Trying to close document..."
                    $doc.Close(0) 
                } catch {
                    Write-Host "  Failed to close document: $($_.Exception.Message)"
                }
                ReleaseComObject $doc
                $doc = $null
            }
        }
    }

    Write-Host "=== CONVERSION COMPLETED ==="
    Write-Host "Success: $successCount/$total files"

} catch {
    Write-Host "CRITICAL ERROR: $($_.Exception.Message)"
    if ($doc) { 
        try { $doc.Close(0) } catch {}
        ReleaseComObject $doc
    }
} finally {
    # Ensure app exits and all resources are released
    Write-Host "Cleaning up resources..."
    if ($app) { 
        try {
            Write-Host "  Quitting application..."
            $app.Quit()
            ReleaseComObject $app
        } catch {
            Write-Host "  Failed to quit application: $($_.Exception.Message)"
        }
        $app = $null
    }
    
    # Force garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "Resource cleanup completed"
}

Write-Host "=== CONVERSION TASK ENDED ==="
Write-Host "End time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# Output result
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
              if (stderr.includes('ERROR') || stderr.includes('Exception')) {
                console.error(`[Office 批量转换] PowerShell错误: ${stderr}`);
              }
            }

            if (error) {
              console.error(`[Office 批量转换] 执行失败: ${error.message}`);

              if (fs.existsSync(tempDir)) {
                fs.rmSync(tempDir, { recursive: true, force: true });
              }

              if (retryCount < MAX_RETRIES - 1) {
                retryCount++;
                setTimeout(executeConversion, RETRY_INTERVAL);
              } else {
                resolve({ success: false, error: error.message });
              }
              return;
            }

            const isSuccess =
              stdout.includes('BATCH_CONVERSION_SUCCESS') ||
              stdout.includes('BATCH_CONVERSION_PARTIAL_SUCCESS');

            if (isSuccess) {
              const pdfResults = [];

              // 先读取PDF文件，再清理临时目录
              for (const pair of inputOutputPairs) {
                const pdfPath = pair.output;

                if (fs.existsSync(pdfPath)) {
                  try {
                    const pdfBuffer = fs.readFileSync(pdfPath);
                    pdfResults.push({
                      name: path.basename(pdfPath),
                      data: pdfBuffer,
                    });
                  } catch (readError) {
                    console.error(
                      `[Office 批量转换] 读取PDF失败: ${pdfPath} - ${readError.message}`
                    );
                  }
                } else {
                  console.error(`[Office 批量转换] PDF文件不存在: ${pdfPath}`);
                }
              }

              // 读取完成后再清理临时目录
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

              if (retryCount < MAX_RETRIES - 1) {
                retryCount++;
                setTimeout(executeConversion, RETRY_INTERVAL);
              } else {
                resolve({ success: false, error: '转换失败' });
              }
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

        if (retryCount < MAX_RETRIES - 1) {
          retryCount++;
          setTimeout(executeConversion, RETRY_INTERVAL);
        } else {
          resolve({ success: false, error: error.message });
        }
      }
    };

    executeConversion();
  });
}

module.exports = {
  checkWordInstallation,
  convertBatchWordToPdf,
};
