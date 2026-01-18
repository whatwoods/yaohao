/**
 * Excel 文件读取与导出模块
 */

const ExcelModule = {
    /**
     * 读取Excel文件并解析数据
     * @param {File} file - 上传的Excel文件
     * @returns {Promise<Array>} - 解析后的数据数组
     */
    async readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // 获取第一个工作表
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    // 转换为JSON，跳过表头
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    // 跳过第一行（表头），解析数据
                    const result = [];
                    const employeeIds = new Set();
                    const duplicates = [];
                    const invalidGenders = [];

                    for (let i = 1; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        if (!row || row.length < 3) continue;

                        // 获取各列值，忽略空值
                        const employeeId = row[0] != null ? String(row[0]).trim() : '';
                        let gender = row[1] != null ? String(row[1]).trim() : '';
                        const roomNumber = row[2] != null ? String(row[2]).trim() : '';

                        // 跳过任何字段为空的行
                        if (!employeeId || !gender || !roomNumber) continue;

                        // 性别格式校验
                        if (gender !== '男' && gender !== '女') {
                            invalidGenders.push(`${employeeId}: "${gender}"`);
                            // 尝试自动修正常见格式
                            if (['m', 'M', 'male', 'Male', 'MALE', '男性'].includes(gender)) {
                                gender = '男';
                            } else if (['f', 'F', 'female', 'Female', 'FEMALE', '女性'].includes(gender)) {
                                gender = '女';
                            }
                        }

                        // 检查重复工号
                        if (employeeIds.has(employeeId)) {
                            duplicates.push(employeeId);
                            continue;
                        }
                        employeeIds.add(employeeId);

                        result.push({
                            employeeId,
                            gender,
                            roomNumber
                        });
                    }

                    // 警告信息
                    const warnings = [];
                    if (duplicates.length > 0) {
                        warnings.push(`重复工号（已忽略）:\n${duplicates.join(', ')}`);
                    }
                    if (invalidGenders.length > 0) {
                        warnings.push(`非标准性别格式（已尝试自动修正）:\n${invalidGenders.join('\n')}`);
                    }
                    if (warnings.length > 0) {
                        alert(warnings.join('\n\n'));
                    }

                    resolve(result);
                } catch (error) {
                    reject(new Error('Excel文件解析失败: ' + error.message));
                }
            };

            reader.onerror = () => {
                reject(new Error('文件读取失败'));
            };

            reader.readAsArrayBuffer(file);
        });
    },

    /**
     * 导出数据为Excel文件
     * @param {Array} data - 要导出的数据
     * @param {string} filename - 文件名
     */
    exportToExcel(data, filename = '摇号结果.xlsx') {
        // 按房间号排序
        const sortedData = [...data].sort((a, b) => a.roomNumber.localeCompare(b.roomNumber, 'zh-CN'));

        // 准备数据
        const exportData = [
            ['序号', '工号', '性别', '房间号']
        ];

        sortedData.forEach((item, index) => {
            exportData.push([
                index + 1,
                item.employeeId,
                item.gender,
                item.roomNumber
            ]);
        });

        // 创建工作簿
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_array ?
            XLSX.utils.aoa_to_sheet(exportData) :
            XLSX.utils.aoa_to_sheet(exportData);

        // 设置列宽
        ws['!cols'] = [
            { wch: 8 },   // 序号
            { wch: 15 },  // 工号
            { wch: 8 },   // 性别
            { wch: 25 }   // 房间号
        ];

        XLSX.utils.book_append_sheet(wb, ws, '摇号结果');

        // 下载文件
        XLSX.writeFile(wb, filename);
    },

    /**
     * 导出数据为PDF文件（支持分页）
     * @param {Array} data - 要导出的数据
     * @param {string} filename - 文件名
     */
    async exportToPDF(data, filename = '摇号结果.pdf') {
        // 按房间号排序
        const sortedData = [...data].sort((a, b) => a.roomNumber.localeCompare(b.roomNumber, 'zh-CN'));

        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');

        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = pdf.internal.pageSize.getHeight();

        // 每页显示的行数
        const ROWS_PER_PAGE = 25;
        const totalPages = Math.ceil(sortedData.length / ROWS_PER_PAGE);

        for (let pageIndex = 0; pageIndex < totalPages; pageIndex++) {
            // 获取当前页的数据
            const startIdx = pageIndex * ROWS_PER_PAGE;
            const endIdx = Math.min(startIdx + ROWS_PER_PAGE, sortedData.length);
            const pageData = sortedData.slice(startIdx, endIdx);

            // 创建临时容器用于渲染
            const container = document.createElement('div');
            container.style.cssText = `
                position: fixed;
                left: -9999px;
                top: 0;
                width: 800px;
                padding: 40px;
                background: white;
                font-family: system-ui, -apple-system, 'Microsoft YaHei', sans-serif;
            `;

            // 构建HTML内容
            container.innerHTML = `
                <div style="display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 8px;">
                    <h1 style="font-size: 24px; color: #333; margin: 0;">公寓摇号结果</h1>
                    <span style="font-size: 12px; color: #999;">第 ${pageIndex + 1} / ${totalPages} 页</span>
                </div>
                <p style="font-size: 14px; color: #666; margin-bottom: 20px;">生成时间: ${new Date().toLocaleString('zh-CN')}</p>
                <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                    <thead>
                        <tr style="background: linear-gradient(135deg, #667eea, #764ba2); color: white;">
                            <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">序号</th>
                            <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">工号</th>
                            <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">性别</th>
                            <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">房间号</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${pageData.map((item, index) => `
                            <tr style="background: ${index % 2 === 0 ? '#f8f9fa' : 'white'};">
                                <td style="padding: 10px; text-align: center; border: 1px solid #ddd;">${startIdx + index + 1}</td>
                                <td style="padding: 10px; text-align: center; border: 1px solid #ddd;">${item.employeeId}</td>
                                <td style="padding: 10px; text-align: center; border: 1px solid #ddd; color: ${item.gender === '男' ? '#3b82f6' : '#ec4899'};">${item.gender}</td>
                                <td style="padding: 10px; text-align: left; border: 1px solid #ddd;">${item.roomNumber}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            `;

            document.body.appendChild(container);

            try {
                // 使用 html2canvas 渲染
                const canvas = await html2canvas(container, {
                    scale: 2,
                    useCORS: true,
                    logging: false
                });

                // 计算缩放比例
                const imgData = canvas.toDataURL('image/png');
                const imgWidth = canvas.width;
                const imgHeight = canvas.height;

                const ratio = Math.min(pdfWidth / (imgWidth * 0.264583), pdfHeight / (imgHeight * 0.264583));
                const scaledWidth = imgWidth * 0.264583 * ratio * 0.9;
                const scaledHeight = imgHeight * 0.264583 * ratio * 0.9;

                // 居中放置
                const x = (pdfWidth - scaledWidth) / 2;
                const y = 10;

                // 添加新页（除了第一页）
                if (pageIndex > 0) {
                    pdf.addPage();
                }

                pdf.addImage(imgData, 'PNG', x, y, scaledWidth, scaledHeight);
            } finally {
                document.body.removeChild(container);
            }
        }

        pdf.save(filename);
    }
};

// 导出模块
window.ExcelModule = ExcelModule;
