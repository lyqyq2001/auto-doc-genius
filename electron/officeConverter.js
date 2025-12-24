const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

/**
 * 转换 Word 为 PDF (兼容 Microsoft Word 和 WPS)
 * 逻辑：优先尝试调用 MS Word，如果失败则尝试调用 WPS
 */
async function convertWordToPdfWithOffice(inputPath, outputPath) {
  let tempScriptPath = null;

  return new Promise((resolve, reject) => {
    try {
      //  路径标准化
      const absInputPath = path.resolve(inputPath);
      const absOutputPath = path.resolve(outputPath);

      // 确保输出目录存在
      const outputDir = path.dirname(absOutputPath);
      if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
      }

      //  创建临时目录
      const tempDir = path.join(os.tmpdir(), 'office_convert_' + Date.now());
      if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir, { recursive: true });
      }

      //  构建 PowerShell 脚本
      const psContent = `
$ErrorActionPreference = "Stop"
$inputPath = "${absInputPath}"
$outputPath = "${absOutputPath}"

Write-Host "Checking input file..."
if (-not (Test-Path -LiteralPath $inputPath)) {
    Write-Error "Input file not found: $inputPath"
    exit 1
}

$app = $null
$appName = ""

try {
    
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
        $appName = "MS Word (Active)"
    } catch {
        try {
            $app = New-Object -ComObject Word.Application
            $appName = "MS Word (New)"
        } catch {
            
            Write-Host "Microsoft Word not found, trying WPS..."
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

    Write-Host "Opening document..."
    
    $doc = $app.Documents.Open($inputPath, $false, $true, $false)
    
    if (-not $doc) { throw "Document object is null." }

    Write-Host "Exporting to PDF..."
    
    $doc.ExportAsFixedFormat($outputPath, 17)
    
    Write-Host "Closing document..."
    $doc.Close(0)
    
    
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
    
    Write-Host "CONVERSION_SUCCESS"
} catch {
    Write-Host "ERROR_OCCURRED"
    Write-Host $_.Exception.Message
    
    if ($doc) { $doc.Close(0) }
    if ($app) { $app.Quit() }
    exit 1
}
`;

      //  写入脚本文件
      tempScriptPath = path.join(tempDir, 'convert.ps1');
      fs.writeFileSync(tempScriptPath, psContent, { encoding: 'utf8' });

      // 执行脚本
      const cmd = `powershell -NoProfile -ExecutionPolicy Bypass -File "${tempScriptPath}"`;

      exec(
        cmd,
        {
          encoding: 'utf8',
          timeout: 120000,
          windowsHide: true,
        },
        (error, stdout, stderr) => {
          //  清理临时目录
          try {
            if (fs.existsSync(tempDir)) {
              fs.rmSync(tempDir, { recursive: true, force: true });
            }
          } catch (e) {}

          if (error) {
            console.error(`[Office] Failed: ${error.message}`);
            if (stdout) console.error(stdout);
            if (stderr) console.error(stderr);
            resolve(false);
            return;
          }

          //  验证结果
          if (
            stdout.includes('CONVERSION_SUCCESS') &&
            fs.existsSync(absOutputPath)
          ) {
            resolve(true);
          } else {
            console.error(`[Office] Output:\n${stdout}`);
            if (stderr) console.error(stderr);
            resolve(false);
          }
        }
      );
    } catch (error) {
      console.error(`[Office] Failed: ${error.message}`);
      // 清理残留
      if (tempScriptPath && fs.existsSync(path.dirname(tempScriptPath))) {
        try {
          fs.rmSync(path.dirname(tempScriptPath), {
            recursive: true,
            force: true,
          });
        } catch (e) {}
      }
      resolve(false);
    }
  });
}

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

      // 准备输入输出路径对
      const pairs = inputOutputPairs.map(pair => ({
        input: path.resolve(pair.input),
        output: path.resolve(pair.output),
      }));

      // 确保所有输出目录存在
      pairs.forEach(pair => {
        const outputDir = path.dirname(pair.output);
        if (!fs.existsSync(outputDir)) {
          fs.mkdirSync(outputDir, { recursive: true });
        }
      });

      // 构建 PowerShell 脚本 - 使用单个 Word 实例批量处理
      const psContent = `
$ErrorActionPreference = "Stop"
$pairs = @(
${pairs.map(p => `    @{Input="${p.input}"; Output="${p.output}"}`).join('\n')}
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

    $total = $pairs.Count
    $current = 0

    foreach ($pair in $pairs) {
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
        (error, stdout, stderr) => {
          // 清理临时目录
          try {
            if (fs.existsSync(tempDir)) {
              fs.rmSync(tempDir, { recursive: true, force: true });
            }
          } catch (e) {}

          if (error) {
            console.error(`[Office Batch] Failed: ${error.message}`);
            if (stdout) console.error(stdout);
            if (stderr) console.error(stderr);
            resolve({ success: false, error: error.message });
            return;
          }

          // 解析结果
          if (stdout.includes('BATCH_CONVERSION_SUCCESS')) {
            try {
              // 查找JSON数组开始和结束的位置
              const successIndex = stdout.indexOf('BATCH_CONVERSION_SUCCESS');
              const jsonStart = stdout.indexOf('[', successIndex);
              const jsonEnd = stdout.lastIndexOf(']');

              if (jsonStart >= 0 && jsonEnd > jsonStart) {
                const jsonStr = stdout.substring(jsonStart, jsonEnd + 1);
                const results = JSON.parse(jsonStr);
                const failed = results.filter(r => !r.Success);
                if (failed.length > 0) {
                  console.warn(`[Office Batch] ${failed.length} files failed`);
                  failed.forEach(f =>
                    console.warn(`  - ${f.Input}: ${f.Error}`)
                  );
                }
                resolve({ success: true, results });
              } else {
                console.warn('[Office Batch] No JSON array found in output');
                resolve({ success: true, results: [] });
              }
            } catch (e) {
              console.error(
                `[Office Batch] Failed to parse results: ${e.message}`
              );
              console.error(`[Office Batch] Output:\n${stdout}`);
              resolve({ success: false, error: 'Failed to parse results' });
            }
          } else {
            console.error(`[Office Batch] Output:\n${stdout}`);
            if (stderr) console.error(stderr);
            resolve({ success: false, error: 'Conversion failed' });
          }
        }
      );
    } catch (error) {
      console.error(`[Office Batch] Failed: ${error.message}`);
      // 清理残留
      if (tempScriptPath && fs.existsSync(path.dirname(tempScriptPath))) {
        try {
          fs.rmSync(path.dirname(tempScriptPath), {
            recursive: true,
            force: true,
          });
        } catch (e) {}
      }
      resolve({ success: false, error: error.message });
    }
  });
}

module.exports = {
  convertWordToPdfWithOffice,
  checkWordInstallation,
  convertBatchWordToPdf,
};
