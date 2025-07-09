const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

let stopRequested = false;
let mainWindow = null;
let loginWindow = null;
let isAuthenticated = false;

// 硬编码密码（方案A）
const CORRECT_PASSWORD = 'xa202377..';

function createLoginWindow() {
  loginWindow = new BrowserWindow({
    width: 500,
    height: 400,
    resizable: false,
    center: true,
    frame: true,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });
  
  loginWindow.setMenuBarVisibility(false);
  loginWindow.loadFile(path.join(__dirname, 'renderer', 'login.html'));
  
  loginWindow.on('closed', () => {
    loginWindow = null;
    if (!isAuthenticated) {
      app.quit();
    }
  });
}

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 700,
    show: false, // 初始不显示，等待登录验证
    webPreferences: {
      preload: path.join(__dirname, 'renderer', 'renderer.js'),
      nodeIntegration: true,
      contextIsolation: false
    }
  });
  
  mainWindow.setMenuBarVisibility(false);
  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));

  // 监听渲染进程的文件夹选择请求
  ipcMain.handle('select-folder', async (event) => {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openDirectory']
    });
    if (result.canceled || result.filePaths.length === 0) {
      return '';
    }
    return result.filePaths[0];
  });
  
  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// 密码验证
ipcMain.handle('verify-password', async (event, password) => {
  return {
    success: password === CORRECT_PASSWORD
  };
});

// 登录成功
ipcMain.on('login-success', () => {
  isAuthenticated = true;
  if (loginWindow) {
    loginWindow.close();
  }
  if (mainWindow) {
    mainWindow.show();
    mainWindow.focus();
  }
});

// 登录失败
ipcMain.on('login-failed', () => {
  app.quit();
});

app.whenReady().then(() => {
  createMainWindow();
  createLoginWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createMainWindow();
    createLoginWindow();
  }
});

ipcMain.handle('start-process', async (event, { inputDir, outputDir }) => {
  // 检查是否已经通过身份验证
  if (!isAuthenticated) {
    return;
  }
  
  stopRequested = false;
  const win = BrowserWindow.getFocusedWindow();
  
  function sendLog(msg) {
    win.webContents.send('log', msg);
  }
  
  sendLog(`开始扫描文件夹: ${inputDir}`);
  
  let files;
  try {
    files = fs.readdirSync(inputDir).filter(f => f.endsWith('.xls') || f.endsWith('.xlsx'));
  } catch (e) {
    sendLog('读取文件夹失败: ' + e.message);
    win.webContents.send('process-finished');
    return;
  }
  
  sendLog(`共发现${files.length}个Excel文件。`);
  
  for (let i = 0; i < files.length; i++) {
    if (stopRequested) {
      sendLog('用户中断，处理已停止。');
      break;
    }
    
    const file = files[i];
    sendLog(`\n开始处理文件(${i+1}/${files.length}): ${file}`);
    
    try {
      const filePath = path.join(inputDir, file);
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      
      // 处理Excel文件（修改内容但保持格式）
      for (const worksheet of workbook.worksheets) {
        sendLog(`  处理sheet: ${worksheet.name}`);
        
        let tableCount = 0;
        // 统计表格数量
        for (let rowIdx = 1; rowIdx <= worksheet.rowCount; rowIdx++) {
          let row = worksheet.getRow(rowIdx);
          let rowStr = row.values.map(v => v ? v.toString() : '').join('');
          if (rowStr.replace(/\s/g, '').includes('界址点成果表')) {
            tableCount++;
          }
        }
        sendLog(`    发现${tableCount}个表格`);
        
        // 修改Excel内容（保持格式）
        let modifiedCount = 0;
        for (let rowIdx = 1; rowIdx <= worksheet.rowCount; rowIdx++) {
          let row = worksheet.getRow(rowIdx);
          
          for (let colIdx = 1; colIdx <= row.cellCount; colIdx++) {
            let cell = row.getCell(colIdx);
            if (cell.value) {
              let cellValue = cell.value.toString();
              
              // 处理需要隐藏的行
              if (cellValue.includes('权利人') || cellValue.includes('建筑占地') || 
                  (cellValue.includes('制表') && (cellValue.includes('校审') || cellValue.includes('日期')))) {
                row.hidden = true;
                for (let c = 1; c <= row.cellCount; c++) {
                  row.getCell(c).value = '';
                }
                row.height = 0.1;
                modifiedCount++;
                break;
              }
              
              // 修改内容
              if (cellValue.includes('宗地号')) {
                let colonIndex = cellValue.indexOf(':');
                if (colonIndex !== -1) {
                  let afterColon = cellValue.substring(colonIndex + 1).trim();
                  let dkbh = afterColon.length >= 3 ? afterColon.slice(-3) : afterColon;
                  let newValue = cellValue.replace('宗地号', '地块编号').replace(afterColon, dkbh);
                  cell.value = newValue;
                  modifiedCount++;
                }
              }
              
              if (cellValue.includes('宗地面积')) {
                let newValue = cellValue.replace('宗地面积', '地块面积');
                cell.value = newValue;
                modifiedCount++;
              }
            }
          }
        }
        
        sendLog(`    完成${modifiedCount}处修改`);
      }
      
      // 保存修改后的Excel文件
      const baseName = path.parse(file).name;
      const newExcelPath = path.join(outputDir, `${baseName}_成果表.xlsx`);
      await workbook.xlsx.writeFile(newExcelPath);
      sendLog(`  Excel文件已保存: ${baseName}_成果表.xlsx`);
      
    } catch (error) {
      sendLog(`处理文件失败: ${error.message}`);
    }
  }
  
  sendLog('\n所有文件处理完成！');
  win.webContents.send('process-finished');
});

ipcMain.handle('stop-process', async (event) => {
  stopRequested = true;
  const win = BrowserWindow.getFocusedWindow();
  win.webContents.send('log', '正在停止处理...');
}); 