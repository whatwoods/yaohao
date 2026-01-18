/**
 * 公寓摇号系统 - 主应用逻辑
 */

document.addEventListener('DOMContentLoaded', () => {
    // DOM 元素
    const uploadSection = document.getElementById('upload-section');
    const lotterySection = document.getElementById('lottery-section');
    const fileInput = document.getElementById('file-input');
    const browseBtn = document.getElementById('browse-btn');
    const dropZone = document.getElementById('drop-zone');
    const tableBody = document.getElementById('table-body');
    const lotteryBtn = document.getElementById('lottery-btn');
    const exportExcelBtn = document.getElementById('export-excel-btn');
    const exportPdfBtn = document.getElementById('export-pdf-btn');
    const homeBtn = document.getElementById('home-btn');

    let isLotteryStarted = false;
    let isLotteryComplete = false;

    // ========================================
    // 文件上传处理
    // ========================================

    // 点击浏览按钮
    browseBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        fileInput.click();
    });

    // 点击拖拽区域
    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    // 文件选择处理
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            handleFile(file);
        }
    });

    // 拖拽事件
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');

        const file = e.dataTransfer.files[0];
        if (file) {
            handleFile(file);
        }
    });

    /**
     * 处理上传的文件
     * @param {File} file
     */
    async function handleFile(file) {
        // 验证文件类型
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];

        if (!validTypes.includes(file.type) &&
            !file.name.endsWith('.xlsx') &&
            !file.name.endsWith('.xls')) {
            alert('请上传 Excel 文件 (.xlsx 或 .xls)');
            return;
        }

        try {
            // 读取文件
            const data = await ExcelModule.readFile(file);

            if (data.length === 0) {
                alert('文件中没有有效数据');
                return;
            }

            // 初始化摇号模块
            LotteryModule.init(data);

            // 显示数据表格
            showLotterySection(data);

        } catch (error) {
            console.error('文件处理错误:', error);
            alert('文件处理失败: ' + error.message);
        }
    }

    // ========================================
    // 数据展示
    // ========================================

    /**
     * 显示摇号区域
     * @param {Array} data
     */
    function showLotterySection(data) {
        uploadSection.classList.add('hidden');
        lotterySection.classList.remove('hidden');

        // 重置状态
        isLotteryStarted = false;
        isLotteryComplete = false;
        lotteryBtn.disabled = false;
        lotteryBtn.textContent = '🎲 开始摇号';
        lotteryBtn.classList.remove('btn-danger');
        lotteryBtn.classList.add('btn-success');
        exportExcelBtn.classList.add('hidden');
        exportPdfBtn.classList.add('hidden');

        // 渲染表格
        renderTable(data, false);
    }

    /**
     * 渲染表格
     * @param {Array} data
     * @param {boolean} isRolling - 是否正在滚动
     */
    function renderTable(data, isRolling = false) {
        tableBody.innerHTML = '';

        data.forEach((item, index) => {
            const tr = document.createElement('tr');

            const genderClass = item.gender === '男' ? 'gender-male' : 'gender-female';
            const rollingClass = isRolling ? 'rolling' : '';

            tr.innerHTML = `
                <td>${index + 1}</td>
                <td class="${rollingClass}">${item.employeeId}</td>
                <td class="${genderClass} ${rollingClass}">${item.gender}</td>
                <td class="${rollingClass}">${item.roomNumber}</td>
            `;

            tableBody.appendChild(tr);
        });
    }

    // ========================================
    // 摇号逻辑
    // ========================================

    lotteryBtn.addEventListener('click', () => {
        if (isLotteryComplete) {
            // 重新摇号：重置状态并立即开始
            isLotteryStarted = false;
            isLotteryComplete = false;
            exportExcelBtn.classList.add('hidden');
            exportPdfBtn.classList.add('hidden');
            homeBtn.classList.add('hidden');
            startLottery();
            return;
        }

        if (!isLotteryStarted) {
            // 开始摇号
            startLottery();
        } else {
            // 停止摇号
            stopLottery();
        }
    });

    /**
     * 开始摇号动画
     */
    function startLottery() {
        isLotteryStarted = true;
        isLotteryComplete = false;
        lotteryBtn.textContent = '🛑 停止摇号';
        lotteryBtn.classList.remove('btn-success');
        lotteryBtn.classList.add('btn-danger');

        LotteryModule.startRolling((data, isRolling) => {
            renderTable(data, isRolling);
        });
    }

    /**
     * 停止摇号并显示结果
     */
    function stopLottery() {
        isLotteryComplete = true;
        lotteryBtn.disabled = false;
        lotteryBtn.textContent = '🔄 重新摇号';
        lotteryBtn.classList.remove('btn-danger');
        lotteryBtn.classList.add('btn-success');

        const result = LotteryModule.stopRolling((data, isRolling) => {
            renderTable(data, isRolling);
        });

        // 显示导出按钮和返回首页按钮
        exportExcelBtn.classList.remove('hidden');
        exportPdfBtn.classList.remove('hidden');
        homeBtn.classList.remove('hidden');

        // 触发撒花动画
        triggerConfetti();
    }

    // ========================================
    // 撒花动画
    // ========================================

    /**
     * 触发撒花庆祝动画
     */
    function triggerConfetti() {
        const duration = 500; // 0.5秒
        const end = Date.now() + duration;

        const colors = ['#667eea', '#764ba2', '#11998e', '#38ef7d', '#ff416c', '#ff4b2b'];

        (function frame() {
            confetti({
                particleCount: 7,
                angle: 60,
                spread: 55,
                origin: { x: 0 },
                colors: colors
            });

            confetti({
                particleCount: 7,
                angle: 120,
                spread: 55,
                origin: { x: 1 },
                colors: colors
            });

            if (Date.now() < end) {
                requestAnimationFrame(frame);
            }
        })();
    }

    // ========================================
    // 导出功能
    // ========================================

    exportExcelBtn.addEventListener('click', () => {
        const data = LotteryModule.getCurrentData();
        const timestamp = new Date().toISOString().slice(0, 10);
        ExcelModule.exportToExcel(data, `摇号结果_${timestamp}.xlsx`);
    });

    exportPdfBtn.addEventListener('click', () => {
        const data = LotteryModule.getCurrentData();
        const timestamp = new Date().toISOString().slice(0, 10);
        ExcelModule.exportToPDF(data, `摇号结果_${timestamp}.pdf`);
    });

    // ========================================
    // 返回首页
    // ========================================

    homeBtn.addEventListener('click', () => {
        // 重置所有状态
        isLotteryStarted = false;
        isLotteryComplete = false;

        // 隐藏摇号区，显示上传区
        lotterySection.classList.add('hidden');
        uploadSection.classList.remove('hidden');

        // 重置按钮显示
        exportExcelBtn.classList.add('hidden');
        exportPdfBtn.classList.add('hidden');
        homeBtn.classList.add('hidden');

        // 重置文件输入（允许重复选择同一文件）
        fileInput.value = '';

        // 重置表格
        tableBody.innerHTML = '';

        // 重置摇号按钮
        lotteryBtn.textContent = '🎲 开始摇号';
        lotteryBtn.classList.remove('btn-danger', 'btn-secondary');
        lotteryBtn.classList.add('btn-success');
    });
});
