document.getElementById('exportTableBtn').addEventListener('click', () => {
    // 向当前活动的页面发送消息，通知它执行表格导出
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        chrome.scripting.executeScript({
            target: { tabId: tabs[0].id },
            files: ['export.js']
        });
    });
});
