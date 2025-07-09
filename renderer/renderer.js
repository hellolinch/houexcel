const { ipcRenderer } = require('electron');

let processing = false;

document.getElementById('select-excel-dir').onclick = async () => {
  const folder = await ipcRenderer.invoke('select-folder');
  if (folder) {
    document.getElementById('excel-dir-path').textContent = folder;
    log('已选择输入文件夹: ' + folder);
  }
};
document.getElementById('select-output-dir').onclick = async () => {
  const folder = await ipcRenderer.invoke('select-folder');
  if (folder) {
    document.getElementById('output-dir-path').textContent = folder;
    log('已选择输出文件夹: ' + folder);
  }
};
document.getElementById('start-btn').onclick = () => {
  if (processing) return;
  const inputDir = document.getElementById('excel-dir-path').textContent.trim();
  const outputDir = document.getElementById('output-dir-path').textContent.trim();
  if (!inputDir || !outputDir) {
    log('请先选择输入和输出文件夹！');
    return;
  }
  processing = true;
  document.getElementById('start-btn').disabled = true;
  document.getElementById('stop-btn').disabled = false;
  log('开始处理...');
  ipcRenderer.invoke('start-process', { inputDir, outputDir });
};
document.getElementById('stop-btn').onclick = () => {
  if (!processing) return;
  log('已请求停止处理');
  ipcRenderer.invoke('stop-process');
};
document.getElementById('clear-log').onclick = () => {
  document.getElementById('log-window').textContent = '';
};

ipcRenderer.on('log', (event, msg) => {
  log(msg);
});
ipcRenderer.on('process-finished', () => {
  processing = false;
  document.getElementById('start-btn').disabled = false;
  document.getElementById('stop-btn').disabled = true;
  log('全部处理完成！');
});

function log(msg) {
  const logWin = document.getElementById('log-window');
  logWin.textContent += `[${new Date().toLocaleTimeString()}] ${msg}\n`;
  logWin.scrollTop = logWin.scrollHeight;
} 