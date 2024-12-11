document.addEventListener('DOMContentLoaded', () => {
    const toggleSwitch = document.getElementById('toggle-switch');

    // 1️⃣ 读取存储的开关状态
    chrome.storage.sync.get('enabled', (data) => {
        toggleSwitch.checked = data.enabled ?? false; // 使用 nullish 合并运算符，确保false不被覆盖
    });

    // 2️⃣ 监听开关的切换事件
    toggleSwitch.addEventListener('change', () => {
        const isEnabled = toggleSwitch.checked; // 获取开关的当前状态
        chrome.storage.sync.set({ enabled: isEnabled }, () => {
            console.log('开关状态已保存:', isEnabled);
        });

        // 3️⃣ 通知内容脚本，显示或隐藏“导出表格”按钮和表格点击
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
            if (tabs.length > 0) {
                chrome.tabs.sendMessage(tabs[0].id, { action: 'toggle', enabled: isEnabled });
                // if (!isEnabled) {
                //     // 关闭插件时，刷新页面
                //     chrome.runtime.sendMessage({ action: 'refreshPage' });
                // }
            } else {
                console.warn('未找到活动的标签页');
            }
        });
    });
});


