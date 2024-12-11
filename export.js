(function () {
    let selectedTables = []; // 存储选中的表格

    // 1. 动态加载本地的 XLSX.js 库
    function loadXLSXLibrary(callback) {
        const script = document.createElement('script');
        script.src = chrome.runtime.getURL('xlsx.full.min.js'); // 确保路径正确
        script.type = 'text/javascript';
        script.onload = callback; // 脚本加载完成后，调用回调
        script.onerror = () => console.error('XLSX 加载失败，请检查路径和 CSP');
        document.head.appendChild(script);
    }

    // 2. 当 XLSX 加载完成后，执行表格导出逻辑
    loadXLSXLibrary(() => {
        if (typeof XLSX === 'undefined') {
            console.error('XLSX 加载失败，请检查路径或 CSP 设置');
            return;
        }
        console.log('XLSX库加载成功');

        // 3. 表格导出为 Excel 的逻辑，支持多个表格导出为多个工作表
        function exportTablesToExcel(tables, filename = '导出表格.xlsx') {
            if (tables.length === 0) {
                alert('请先点击一个或多个表格！');
                return;
            }

            const workbook = XLSX.utils.book_new();
            tables.forEach((table, index) => {
                const worksheet = XLSX.utils.table_to_sheet(table);
                // 遍历每一个单元格，处理长整型和时间
                Object.keys(worksheet).forEach(cell => {
                    if (cell[0] === '!') return; // 忽略 !ref 等元数据

                    const cellValue = worksheet[cell].v;

                    // 1️⃣ 处理长整型：如果长度大于 15 位，将其转换为字符串
                    if (typeof cellValue === 'number' && cellValue > 999999999999999) {
                        worksheet[cell].t = 's'; // 将类型强制设置为字符串
                        worksheet[cell].v = String(cellValue);
                    }

                    // 2️⃣ 处理时间：如果是时间格式，则将其转换为Excel的日期格式
                    const datePattern = /^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2})?$/; // 例如 2024-12-11 或 2024-12-11 15:30:00
                    if (typeof cellValue === 'string' && datePattern.test(cellValue)) {
                        const date = new Date(cellValue);
                        const excelDate = (date - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000); // 转换为Excel日期
                        worksheet[cell].t = 'n'; // 类型为数值
                        worksheet[cell].v = excelDate; // 设置Excel的日期值
                        worksheet[cell].z = 'yyyy-mm-dd hh:mm:ss'; // 指定单元格的显示格式
                    }
                });
                // 4️⃣ 美化表头样式（背景色、字体加粗、字体颜色）
                const range = XLSX.utils.decode_range(worksheet['!ref']);
                for (let C = range.s.c; C <= range.e.c; C++) {
                    const headerCell = XLSX.utils.encode_cell({ r: 0, c: C }); // 第一行的每一列
                    if (worksheet[headerCell]) {
                        worksheet[headerCell].s = {
                            font: { bold: true, color: { rgb: "FFFFFF" } },
                            fill: { fgColor: { rgb: "4F81BD" } }, // 浅蓝色
                            alignment: { horizontal: 'center', vertical: 'center' }
                        };
                    }
                }

                // 5️⃣ 隔行变色 (奇数白色，偶数灰色)
                for (let R = 1; R <= range.e.r; R++) { // 从第1行（去掉表头）开始遍历
                    const backgroundColor = (R % 2 === 0) ? "F2F2F2" : "FFFFFF"; // 偶数行灰色，奇数行白色
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

            XLSX.writeFile(workbook, finalFilename);
        }

        // 4. 监听点击事件，选中表格
        document.addEventListener('click', (event) => {
            const table = event.target.closest('table'); // 找到离点击点最近的表格
            if (!table) return; // 如果点击的不是表格，则不处理

            if (selectedTables.includes(table)) {
                // 如果表格已选中，则取消选中
                table.style.border = '';
                selectedTables = selectedTables.filter(t => t !== table);
            } else {
                // 如果表格未选中，则选中表格
                table.style.border = '3px solid red'; // 给表格加红色边框
                selectedTables.push(table);
            }

            console.log('选中的表格：', selectedTables);
        });

        // 5. 监听用户点击“导出”按钮
        const exportButton = document.createElement('button');
        exportButton.textContent = '导出表格';
        exportButton.style.position = 'fixed';
        exportButton.style.top = '10px';
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
            exportTablesToExcel(selectedTables);
        });
    });
})();
