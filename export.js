(function () {
    let exportButton = null;
    let isEnabled = false;
    let selectedWorkSheets = []; // 存储选中的工作表
    let selectedTables = []; // 存储选中的表格
    let handleClickOnTable; // 将事件处理器设置为全局变量，便于后续删除监听器

    // 🔥 开启监听器：监听点击表格的事件
    function enableTableClick() {
        if (!handleClickOnTable) {
            toggleTableClick(isEnabled); // 开启按钮
            handleClickOnTable = (event) => {
                const table = event.target.closest('table'); // 找到离点击点最近的表格

                if (!table) return; // 如果点击的不是表格，则不处理
                const worksheet = XLSX.utils.table_to_sheet(table);
                const worksheetString = JSON.stringify(worksheet);

                if (selectedWorkSheets.includes(worksheetString)
                    && selectedTables.includes(table)) {
                    // 如果表格已选中，则取消选中
                    table.style.border = '';
                    selectedWorkSheets = selectedWorkSheets.filter(t => t !== worksheetString);
                    selectedTables = selectedTables.filter(t => t !== table);
                } else {
                    // 如果表格未选中，则选中表格
                    table.style.border = '3px solid black'; // 给表格加红色边框
                    selectedWorkSheets.push(worksheetString);
                    selectedTables.push(table);
                }

                // console.log('选中的表格：', selectedTables);
            };
        }
        document.addEventListener('click', handleClickOnTable);
        console.log('✅ 表格点击监听器已启用');
    }

    // 🔥 关闭监听器：移除点击事件监听器
    function disableTableClick() {
        if (handleClickOnTable) {
            document.removeEventListener('click', handleClickOnTable);
            console.log('❌ 表格点击监听器已移除');
        }
    }

    // 显示或隐藏导出按钮
    function toggleExportButton(visible) {
        if (!visible && exportButton) {
            exportButton.remove();
            exportButton = null;
            return;
        }
        if (visible && !exportButton) {
            exportButton = document.createElement('button');
            exportButton.textContent = '导出表格';
            exportButton.style.position = 'fixed';
            exportButton.style.bottom = '10px';
            exportButton.style.right = '10px';
            exportButton.style.zIndex = 9999;
            exportButton.style.padding = '10px 20px';
            exportButton.style.backgroundColor = '#4CAF50';
            exportButton.style.color = '#fff';
            exportButton.style.border = 'none';
            exportButton.style.borderRadius = '5px';
            exportButton.style.cursor = 'pointer';
            document.body.appendChild(exportButton);

            exportButton.addEventListener('click', () => {
                exportTable(selectedWorkSheets);
            });
            //开启监听
            enableTableClick();

            if (typeof XLSX === 'undefined') {
                console.log('XLSX 加载失败，请检查路径或 CSP 设置');
                return;
            }
            console.log('XLSX库加载成功');
        }
    }

    // 禁用或启用表格点击事件
    function toggleTableClick(enabled) {
        document.querySelectorAll('table').forEach(table => {
            if (enabled) {
                table.style.pointerEvents = 'auto';
            } else {
                table.style.pointerEvents = 'none';
            }
        });
    }

    // 日期转换函数
    function formatExcelDate(serial) {
        var utc_days = Math.floor(serial - 25569);
        var utc_value = utc_days * 86400;
        var date_info = new Date(utc_value * 1000);

        var fractional_day = serial - Math.floor(serial) + 0.0000001;

        var total_seconds = Math.floor(86400 * fractional_day);

        var seconds = total_seconds % 60;

        total_seconds -= seconds;

        var hours = Math.floor(total_seconds / (60 * 60));
        var minutes = Math.floor(total_seconds / 60) % 60;

        return formatDate(new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds));
    }

    // 日期转换函数
    function formatDate(date) {
        const year = date.getFullYear();
        const month = (date.getMonth() + 1).toString().padStart(2, '0'); // 月份从0开始，需要加1，并确保两位数
        const day = date.getDate().toString().padStart(2, '0'); // 日期确保两位数
        const hours = date.getHours().toString().padStart(2, '0'); // 小时确保两位数
        const minutes = date.getMinutes().toString().padStart(2, '0'); // 分钟确保两位数
        const seconds = date.getSeconds().toString().padStart(2, '0'); // 秒确保两位数

        return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
    }

    // 表格导出逻辑
    function exportTable(worksheets, filename = '导出表格.xlsx') {
        if (worksheets.length === 0) {
            alert('请先点击一个或多个表格！');
            return;
        }

        const workbook = XLSX.utils.book_new();
        worksheets.forEach((worksheet, index) => {
            worksheet = JSON.parse(worksheet);
            const totalColumns = XLSX.utils.decode_range(worksheet['!ref']).e.c + 1; // 获取总列数
            worksheet['!cols'] = Array(totalColumns).fill({ wch: 25 }); // 每一列的宽度为25

            // 🔥 遍历每一个单元格，处理长整型和时间戳
            Object.keys(worksheet).forEach(cell => {
                if (cell[0] === '!') return; // 忽略 !ref 等元数据

                const cellValue = worksheet[cell].v;

                // 1️⃣ 处理长整型：如果长度大于 15 位，将其转换为字符串
                if (typeof cellValue === 'number' && cellValue > 999999999999999) {
                    worksheet[cell].t = 's'; // 强制设置为字符串类型
                    worksheet[cell].v = String(cellValue); // 转换为字符串，避免丢失精度
                    worksheet[cell].s = worksheet[cell].s || {};  // 初始化单元格样式
                    worksheet[cell].s.numFmt = '@'; // 设置单元格格式为文本
                }
                const datestr = worksheet[cell].z ? worksheet[cell].z : null;
                // 2️⃣ 处理时间戳：如果是 10 位或 13 位的时间戳，将其转换为自定义格式的文本
                if (datestr && datestr === 'm/d/yy') {
                    // 检查是否为合理的日期范围
                    const dateString = formatExcelDate(cellValue);
                    if (dateString) {
                        worksheet[cell].t = 's'; // 强制设置为字符串类型
                        worksheet[cell].v = dateString; // 转换为日期格式的字符串
                        worksheet[cell].s = worksheet[cell].s || {};  // 初始化单元格样式
                        worksheet[cell].s.numFmt = '@'; // 设置单元格格式为文本
                    }
                }

            });

            // 3️⃣ 美化表头样式
            const headerStyle = {
                font: { bold: true, sz: 17, color: { rgb: 'FFFFFF' } }, // 加粗、字体大小14、字体白色
                fill: { fgColor: { rgb: '4F81BD' } }, // 浅蓝色背景
                alignment: { horizontal: 'center', vertical: 'center' } // 水平垂直居中
            };

            const range = XLSX.utils.decode_range(worksheet['!ref']);
            for (let C = range.s.c; C <= range.e.c; C++) {
                const headerCell = XLSX.utils.encode_cell({ r: 0, c: C }); // 第一行的每一列
                if (worksheet[headerCell]) {
                    worksheet[headerCell].s = headerStyle; // 设置表头样式
                }
            }

            // 4️⃣ 隔行变色 (奇数白色，偶数灰色)
            const oddRowColor = 'FFFFFF';
            const evenRowColor = 'F2F2F2';

            for (let R = 1; R <= range.e.r; R++) { // 从第1行（去掉表头）开始遍历
                const backgroundColor = (R % 2 === 0) ? evenRowColor : oddRowColor;
                for (let C = range.s.c; C <= range.e.c; C++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                    if (worksheet[cellAddress]) {
                        worksheet[cellAddress].s = worksheet[cellAddress].s || {};
                        worksheet[cellAddress].s.fill = { fgColor: { rgb: backgroundColor } };
                    }
                }
            }

            XLSX.utils.book_append_sheet(workbook, worksheet, `Sheet${index + 1}`);
        });

        const userFilename = prompt('请输入要保存的文件名', filename);
        if (!userFilename) return; // 用户未输入文件名
        const finalFilename = userFilename.endsWith('.xlsx') ? userFilename : `${userFilename}.xlsx`;

        XLSX.writeFile(workbook, finalFilename, { cellStyles: true });
    }

    //还原页面状态
    function restorePage(isEnabled) {
        if (!isEnabled) {
            selectedTables.forEach((table) => {
                table.style.border = '';
            });
            selectedWorkSheets = [];
            //移除监听器
            disableTableClick();
        }
        return;
    }

    // 5. 接收消息，控制显示导出按钮和表格点击
    chrome.runtime.onMessage.addListener((message) => {
        if (message.action === 'toggle') {
            isEnabled = message.enabled;
            toggleExportButton(isEnabled);
            restorePage(isEnabled);
        }
    });

    // 6. 页面加载时，获取开关状态
    chrome.storage.sync.get('enabled', (data) => {
        isEnabled = data.enabled || false;
        toggleExportButton(isEnabled);
    });

})();
