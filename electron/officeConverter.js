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

/**
 * 批量转换 Word 为 PDF - 优化版本
 * 只启动一次 Word 应用，批量处理所有文件，大幅提升速度
 */
async function convertBatchWordToPdf(inputOutputPairs) {
  return new Promise((resolve, reject) => {
    let tempScriptPath = null;

    try {
      if (!inputOutputPairs || inputOutputPairs.length === 0) {
        resolve({ success: true, results: [] });
        return;
      }

      // 创建临时目录
      const tempDir = path.join(
        os.tmpdir(),
        'office_batch_convert_' + Date.now()
      );
      if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir, { recursive: true });
      }
      // 确保所有输出目录存在
      inputOutputPairs.forEach(pair => {
        const outputDir = path.dirname(pair.output);
        if (!fs.existsSync(outputDir)) {
          fs.mkdirSync(outputDir, { recursive: true });
        }
      });
      // 构建 PowerShell 脚本 - 使用单个 Word 实例批量处理
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

    $total = $inputOutputPairs.Count
    $current = 0

    foreach ($pair in $inputOutputPairs) {
        $current++
        Write-Host "Processing [$current/$total]: $($pair.Input)"

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

            if (Test-Path -LiteralPath $pair.Output) {
                $results += @{"Success"=$true;"Input"=$pair.Input;"Output"=$pair.Output}
            } else {
                $results += @{"Success"=$false;"Input"=$pair.Input;"Error"="Output file not created"}
            }

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        } catch {
            $results += @{"Success"=$false;"Input"=$pair.Input;"Error"=$_.Exception.Message}
            if ($doc) { $doc.Close(0) }
        }
    }

    $app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null

    Write-Host "BATCH_CONVERSION_SUCCESS"
    $results | ConvertTo-Json -Compress
} catch {
    Write-Host "ERROR_OCCURRED"
    Write-Host $_.Exception.Message
    if ($doc) { $doc.Close(0) }
    if ($app) { $app.Quit() }
    exit 1
}
`;

      // 写入脚本文件
      tempScriptPath = path.join(tempDir, 'batch_convert.ps1');
      fs.writeFileSync(tempScriptPath, psContent, { encoding: 'utf8' });

      // 执行脚本
      const cmd = `powershell -NoProfile -ExecutionPolicy Bypass -File "${tempScriptPath}"`;

      exec(
        cmd,
        {
          encoding: 'utf8',
          timeout: 300000,
          windowsHide: true,
        },
        (error, stdout) => {
          // 清理临时目录
          if (fs.existsSync(tempDir)) {
            fs.rmSync(tempDir, { recursive: true, force: true });
          }

          if (error) {
            console.error(`[Office 批量转换] 失败: ${error.message}`);
            resolve({ success: false, error: error.message });
            return;
          }

          // 解析成功的结果 stdout是命令执行后输出的内容 BATCH_CONVERSION_SUCCESS是我们自定义的成功标记
          if (stdout.includes('BATCH_CONVERSION_SUCCESS')) {
            resolve({ success: true });
          } else {
            console.error(`[Office 批量转换] 输出:\n${stdout}`);

            resolve({ success: false, error: '转换失败' });
          }
        }
      );
    } catch (error) {
      console.error(`[Office 批量转换] 失败: ${error.message}`);
      // 清理残留

      fs.rmSync(path.dirname(tempScriptPath), {
        recursive: true,
        force: true,
      });

      resolve({ success: false, error: error.message });
    }
  });
}

module.exports = {
  checkWordInstallation,
  convertBatchWordToPdf,
};
