// 1️⃣ 插件安装时，默认设置 enabled 为 false
chrome.runtime.onInstalled.addListener(() => {
    console.log('插件已安装');
    chrome.storage.sync.set({ enabled: false });
});

// 2️⃣ 监听消息事件，控制开关状态和刷新当前页面
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    console.log('收到的消息:', message);

    switch (message.action) {
        case 'toggle':
            handleToggleAction(message.enabled);
            break;
        case 'refreshPage':
            handleRefreshPage();
            break;
        default:
            console.warn('未识别的消息操作:', message.action);
    }

    sendResponse({ status: 'success' });
});

// 🟢 处理“toggle”消息，更新开关状态
function handleToggleAction(isEnabled) {
    console.log('开关切换为:', isEnabled);
    chrome.storage.sync.set({ enabled: isEnabled }, () => {
        console.log('开关状态已同步到存储中:', isEnabled);
    });
}

// 🔄 处理“refreshPage”消息，刷新当前活动的标签页
function handleRefreshPage() {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        if (tabs.length > 0) {
            const tabId = tabs[0].id;

            // 1️⃣ 在当前网页上弹出 confirm 对话框
            chrome.scripting.executeScript({
                target: { tabId: tabId },
                func: () => {
                    return confirm('不刷新页面表格无法点击，您确定要刷新页面吗？'); // 在网页中执行 confirm
                }
            }, (results) => {
                // 2️⃣ 结果是一个数组，包含每个选项卡的返回值（true 或 false）
                const userConfirmed = results[0]?.result;
                if (userConfirmed) {
                    // 3️⃣ 如果用户确认，刷新页面
                    chrome.tabs.reload(tabId, () => {
                        console.log('已刷新当前标签页:', tabId);
                    });
                } else {
                    console.log('用户取消了刷新操作');
                }
            });
        } else {
            console.warn('未找到活动的标签页，无法刷新页面');
        }
    });
}
