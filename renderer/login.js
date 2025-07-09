const { ipcRenderer } = require('electron');

let attemptsLeft = 3;

document.getElementById('login-btn').onclick = () => {
  const password = document.getElementById('password').value;
  const errorMessage = document.getElementById('error-message');
  const attemptsCount = document.getElementById('attempts-count');
  const loginBox = document.querySelector('.login-box');
  
  if (!password) {
    showError('请输入密码！');
    return;
  }
  
  // 发送密码验证请求到主进程
  ipcRenderer.invoke('verify-password', password).then(result => {
    if (result.success) {
      // 密码正确，显示成功信息并关闭登录窗口
      errorMessage.textContent = '登录成功！';
      errorMessage.style.color = '#27ae60';
      
      setTimeout(() => {
        ipcRenderer.send('login-success');
      }, 500);
    } else {
      // 密码错误
      attemptsLeft--;
      attemptsCount.textContent = attemptsLeft;
      
      if (attemptsLeft > 0) {
        showError(`密码错误！还有 ${attemptsLeft} 次尝试机会`);
        // 添加震动效果
        loginBox.classList.add('shake');
        setTimeout(() => {
          loginBox.classList.remove('shake');
        }, 500);
      } else {
        showError('密码错误次数过多，程序将退出！');
        document.getElementById('login-btn').disabled = true;
        document.getElementById('password').disabled = true;
        
        setTimeout(() => {
          ipcRenderer.send('login-failed');
        }, 2000);
      }
    }
  });
  
  // 清空密码输入框
  document.getElementById('password').value = '';
};

// 支持回车键登录
document.getElementById('password').addEventListener('keypress', (e) => {
  if (e.key === 'Enter') {
    document.getElementById('login-btn').click();
  }
});

function showError(message) {
  const errorMessage = document.getElementById('error-message');
  errorMessage.textContent = message;
  errorMessage.style.color = '#e74c3c';
}

// 页面加载完成后聚焦到密码输入框
window.addEventListener('load', () => {
  document.getElementById('password').focus();
}); 