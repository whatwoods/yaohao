/**
 * Excel 文件读取与导出模块
 */

/**
 * HTML 转义，防止 XSS
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
    if (str == null) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

const ExcelModule = {
    /**
     * 读取Excel文件并解析数据
     * @param {File} file - 上传的Excel文件
     * @returns {Promise<{data: Array, idLabel: string}>} - data: 数据数组；idLabel: 第一列表头（"工号" 或 "姓名" 等）
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

                    // 读取表头第一列，识别是"工号"还是"姓名"
                    // 兼容多种常见写法；默认为"工号"
                    const rawHeader = jsonData[0] || [];
                    const firstColHeader = rawHeader[0] != null ? String(rawHeader[0]).trim() : '';
                    let idLabel = '工号';
                    const nameKeywords = ['姓名', '名字', '名称', '员工姓名', '人员姓名'];
                    const idKeywords = ['工号', '员工编号', '编号', '员工号', '人员编号', '工牌号', 'EmpID', 'empid', 'emp_id'];
                    if (nameKeywords.some(k => firstColHeader.includes(k))) {
                        idLabel = '姓名';
                    } else if (idKeywords.some(k => firstColHeader.toLowerCase().includes(k.toLowerCase()))) {
                        idLabel = '工号';
                    } else if (firstColHeader) {
                        // 无法识别的表头，沿用原表头文字
                        idLabel = firstColHeader;
                    }

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
                            // 尝试自动修正常见格式
                            let normalizedGender = null;
                            if (['m', 'M', 'male', 'Male', 'MALE', '男性'].includes(gender)) {
                                normalizedGender = '男';
                            } else if (['f', 'F', 'female', 'Female', 'FEMALE', '女性'].includes(gender)) {
                                normalizedGender = '女';
                            }

                            if (normalizedGender) {
                                invalidGenders.push(`${employeeId}: "${gender}" -> "${normalizedGender}"`);
                                gender = normalizedGender;
                            } else {
                                // 无法识别的性别，跳过该行
                                invalidGenders.push(`${employeeId}: "${gender}"（无法识别，已跳过该行）`);
                                continue;
                            }
                        }

                        // 检查重复工号/姓名
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
                        warnings.push(`重复${idLabel}（已忽略）:\n${duplicates.join(', ')}`);
                    }
                    if (invalidGenders.length > 0) {
                        warnings.push(`非标准性别格式（已尝试自动修正）:\n${invalidGenders.join('\n')}`);
                    }
                    if (warnings.length > 0) {
                        alert(warnings.join('\n\n'));
                    }

                    resolve({ data: result, idLabel });
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
     * @param {string} idLabel - 第一列表头（"工号" 或 "姓名" 等）
     */
    exportToExcel(data, filename = '摇号结果.xlsx', idLabel = '工号') {
        // 按房间号排序
        const sortedData = [...data].sort((a, b) => a.roomNumber.localeCompare(b.roomNumber, 'zh-CN'));

        // 准备数据（表头使用传入的 idLabel）
        const exportData = [
            ['序号', idLabel, '性别', '房间号']
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
        const ws = XLSX.utils.aoa_to_sheet(exportData);

        // 设置列宽
        ws['!cols'] = [
            { wch: 8 },   // 序号
            { wch: 15 },  // 工号/姓名
            { wch: 8 },   // 性别
            { wch: 25 }   // 房间号
        ];

        XLSX.utils.book_append_sheet(wb, ws, '摇号结果');

        // 下载文件
        XLSX.writeFile(wb, filename);
    },

    /**
     * 导出数据为PDF文件（高清 JPEG 嵌入版）
     *
     * 参数取舍说明：
     *   - 用户期望体积约 1MB（之前 312KB 版本清晰度不足）
     *   - scale=2 + JPEG q=0.92 + 容器 800px：综合体积约 0.8~1.2MB（50 人场景）
     *   - 相比 312KB 版本（scale=1.5, q=0.85, 700px），清晰度显著提升
     *
     * @param {Array} data - 要导出的数据
     * @param {string} filename - 文件名
     * @param {string} idLabel - 第一列表头（"工号" 或 "姓名" 等）
     */
    async exportToPDF(data, filename = '摇号结果.pdf', idLabel = '工号') {
        // 按房间号排序
        const sortedData = [...data].sort((a, b) => a.roomNumber.localeCompare(b.roomNumber, 'zh-CN'));

        const total = sortedData.length;
        const male = sortedData.filter(item => item.gender === '男').length;
        const female = sortedData.filter(item => item.gender === '女').length;

        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');

        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = pdf.internal.pageSize.getHeight();

        // 每页显示的行数
        const ROWS_PER_PAGE = 25;
        const totalPages = Math.ceil(sortedData.length / ROWS_PER_PAGE);

        // 设计 token（与页面 CSS 保持一致）
        const bg = '#FDFCF8';
        const surface = '#FFFFFF';
        const border = '#E8E4D9';
        const textPrimary = '#1C1C1A';
        const textSecondary = '#6B6A63';
        const textMuted = '#9C9B91';
        const accent = '#D97706';
        const maleColor = '#2563EB';
        const femaleColor = '#DB2777';

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
                background: ${surface};
                font-family: 'Noto Sans SC', 'Geist', 'PingFang SC', 'Microsoft YaHei', sans-serif;
                color: ${textPrimary};
            `;

            const headerHtml = pageIndex === 0 ? `
                <div style="margin-bottom: 16px;">
                    <div style="display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 8px;">
                        <h1 style="font-size: 24px; font-weight: 600; color: ${textPrimary}; margin: 0;">公寓摇号结果</h1>
                        <span style="font-size: 12px; color: ${textMuted};">生成时间: ${new Date().toLocaleString('zh-CN')}</span>
                    </div>
                    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 16px;">
                        <div style="background: ${bg}; border: 2px solid ${border}; border-radius: 6px; padding: 16px;">
                            <div style="font-size: 12px; color: ${textSecondary}; margin-bottom: 4px;">总人数</div>
                            <div style="font-size: 24px; font-weight: 600; color: ${textPrimary};">${total}</div>
                        </div>
                        <div style="background: ${bg}; border: 2px solid ${border}; border-radius: 6px; padding: 16px;">
                            <div style="font-size: 12px; color: ${textSecondary}; margin-bottom: 4px;">男生</div>
                            <div style="font-size: 24px; font-weight: 600; color: ${textPrimary};">${male}</div>
                        </div>
                        <div style="background: ${bg}; border: 2px solid ${border}; border-radius: 6px; padding: 16px;">
                            <div style="font-size: 12px; color: ${textSecondary}; margin-bottom: 4px;">女生</div>
                            <div style="font-size: 24px; font-weight: 600; color: ${textPrimary};">${female}</div>
                        </div>
                    </div>
                </div>
            ` : `
                <div style="display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 20px;">
                    <h1 style="font-size: 24px; font-weight: 600; color: ${textPrimary}; margin: 0;">公寓摇号结果</h1>
                    <span style="font-size: 12px; color: ${textMuted};">第 ${pageIndex + 1} / ${totalPages} 页</span>
                </div>
            `;

            // 构建HTML内容（表头使用传入的 idLabel）
            container.innerHTML = `
                ${headerHtml}
                <table style="width: 100%; border-collapse: collapse; font-size: 14px; border: 2px solid ${border}; border-radius: 6px; overflow: hidden;">
                    <thead>
                        <tr style="background: ${bg}; color: ${textPrimary}; border-bottom: 2px solid ${border};">
                            <th style="padding: 12px; text-align: center; border-right: 1px solid ${border}; font-weight: 600;">序号</th>
                            <th style="padding: 12px; text-align: center; border-right: 1px solid ${border}; font-weight: 600;">${escapeHtml(idLabel)}</th>
                            <th style="padding: 12px; text-align: center; border-right: 1px solid ${border}; font-weight: 600;">性别</th>
                            <th style="padding: 12px; text-align: left; font-weight: 600;">房间号</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${pageData.map((item, index) => `
                            <tr style="background: ${index % 2 === 0 ? surface : bg}; border-bottom: 1px solid ${border};">
                                <td style="padding: 10px; text-align: center; border-right: 1px solid ${border};">${startIdx + index + 1}</td>
                                <td style="padding: 10px; text-align: center; border-right: 1px solid ${border};">${escapeHtml(item.employeeId)}</td>
                                <td style="padding: 10px; text-align: center; border-right: 1px solid ${border}; color: ${item.gender === '男' ? maleColor : femaleColor}; font-weight: 500;">${escapeHtml(item.gender)}</td>
                                <td style="padding: 10px; text-align: left;">${escapeHtml(item.roomNumber)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
                ${pageIndex === 0 && totalPages > 1 ? `<div style="margin-top: 8px; text-align: right; font-size: 12px; color: ${textMuted};">第 1 / ${totalPages} 页</div>` : ''}
            `;

            document.body.appendChild(container);

            try {
                // 使用 html2canvas 渲染
                // scale=2 保证高清，配合 JPEG q=0.92 平衡体积（约 1MB）
                const canvas = await html2canvas(container, {
                    scale: 2,
                    useCORS: true,
                    logging: false,
                    backgroundColor: surface
                });

                // JPEG + quality 0.92（高清但有压缩，体积约 1MB）
                const imgData = canvas.toDataURL('image/jpeg', 0.92);
                const imgWidth = canvas.width;
                const imgHeight = canvas.height;

                // px → mm 换算
                const pxToMm = 0.264583;
                const ratio = Math.min(pdfWidth / (imgWidth * pxToMm), pdfHeight / (imgHeight * pxToMm));
                const scaledWidth = imgWidth * pxToMm * ratio * 0.92;
                const scaledHeight = imgHeight * pxToMm * ratio * 0.92;

                // 居中放置
                const x = (pdfWidth - scaledWidth) / 2;
                const y = 10;

                // 添加新页（除了第一页）
                if (pageIndex > 0) {
                    pdf.addPage();
                }

                pdf.addImage(imgData, 'JPEG', x, y, scaledWidth, scaledHeight, undefined, 'FAST');
            } finally {
                document.body.removeChild(container);
            }
        }

        pdf.save(filename);
    }
};

// 导出模块
window.ExcelModule = ExcelModule;
