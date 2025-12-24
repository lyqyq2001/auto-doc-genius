const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

/**
 * 转换 Word 为 PDF (兼容 Microsoft Word 和 WPS)
 * 逻辑：优先尝试调用 MS Word，如果失败则尝试调用 WPS
 */
function convertWordToPdfWithOffice(inputPath, outputPath) {
  let tempScriptPath = null;

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

    const result = execSync(cmd, {
      encoding: 'utf8',
      timeout: 120000,
      windowsHide: true,
    });

    //  清理临时目录
    try {
      if (fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
      }
    } catch (e) {}

    //  验证结果
    if (result.includes('CONVERSION_SUCCESS') && fs.existsSync(absOutputPath)) {
      return true;
    } else {
      console.error(`[Office] Output:\n${result}`);
      return false;
    }
  } catch (error) {
    console.error(`[Office] Failed: ${error.message}`);
    if (error.stdout) console.error(error.stdout.toString());
    // 清理残留
    if (tempScriptPath && fs.existsSync(path.dirname(tempScriptPath))) {
      try {
        fs.rmSync(path.dirname(tempScriptPath), {
          recursive: true,
          force: true,
        });
      } catch (e) {}
    }
    return false;
  }
}

function checkWordInstallation() {
  // 简单检查：只要能调起 Word 或 WPS 就算已安装
  try {
    const checkCmd = `
        try { 
            $a = New-Object -ComObject Word.Application; $a.Quit() 
        } catch { 
            $b = New-Object -ComObject Kwps.Application; $b.Quit() 
        }
        `;
    execSync(`powershell -Command "${checkCmd}"`, {
      stdio: 'ignore',
      timeout: 10000,
    });
    return true;
  } catch (e) {
    return false;
  }
}

module.exports = {
  convertWordToPdfWithOffice,
  checkWordInstallation,
};
