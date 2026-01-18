/**
 * 摇号核心算法模块
 * 处理同房间同性别分配逻辑
 */

const LotteryModule = {
    originalData: [],      // 原始数据备份
    currentData: [],       // 当前显示数据
    isRolling: false,      // 是否正在摇号
    rollInterval: null,    // 滚动定时器

    /**
     * 初始化数据
     * @param {Array} data - 原始数据
     */
    init(data) {
        this.originalData = JSON.parse(JSON.stringify(data));
        this.currentData = JSON.parse(JSON.stringify(data));
        this.isRolling = false;
    },

    /**
     * 解析房间号，提取基础房间ID（用于判断同一房间）
     * 例如："南-16-1-1001-A" -> "南-16-1-1001"
     * @param {string} roomNumber - 完整房间号
     * @returns {object} - { baseRoom: 基础房间号, bedroom: 卧室号或null }
     */
    parseRoomNumber(roomNumber) {
        const parts = roomNumber.split('-');

        // 检查最后一部分是否为卧室号（通常是单个字母如A、B）
        const lastPart = parts[parts.length - 1];
        const isBedroomSuffix = /^[A-Za-z]$/.test(lastPart);

        if (isBedroomSuffix && parts.length > 4) {
            return {
                baseRoom: parts.slice(0, -1).join('-'),
                bedroom: lastPart
            };
        }

        return {
            baseRoom: roomNumber,
            bedroom: null
        };
    },

    /**
     * 核心摇号算法
     * 确保同一房间内的不同卧室分配给同性别的人
     * @returns {Array} - 分配结果
     */
    performLottery() {
        const data = JSON.parse(JSON.stringify(this.originalData));

        // 分离人员和房间
        const employees = data.map(d => ({
            employeeId: d.employeeId,
            gender: d.gender
        }));
        const rooms = data.map(d => d.roomNumber);

        // 按性别分组人员
        const maleEmployees = employees.filter(e => e.gender === '男');
        const femaleEmployees = employees.filter(e => e.gender === '女');

        // 分析房间结构，找出共享房间
        const roomGroups = {};
        rooms.forEach((room, index) => {
            const parsed = this.parseRoomNumber(room);
            if (!roomGroups[parsed.baseRoom]) {
                roomGroups[parsed.baseRoom] = [];
            }
            roomGroups[parsed.baseRoom].push({
                fullRoom: room,
                bedroom: parsed.bedroom,
                originalIndex: index
            });
        });

        // 分离共享房间和单独房间
        const sharedRooms = [];  // 需要分配2人的房间组
        const singleRooms = [];  // 只需分配1人的房间

        Object.entries(roomGroups).forEach(([baseRoom, roomList]) => {
            if (roomList.length >= 2) {
                sharedRooms.push(roomList);
            } else {
                singleRooms.push(roomList[0]);
            }
        });

        // 随机打乱人员顺序
        this.shuffleArray(maleEmployees);
        this.shuffleArray(femaleEmployees);

        // 结果数组
        const result = new Array(data.length);

        // 打乱共享房间顺序
        this.shuffleArray(sharedRooms);

        // 分配共享房间 - 严格同性别
        const unassignedSharedRooms = [];
        for (const roomGroup of sharedRooms) {
            const roomCount = roomGroup.length;

            // 严格选择有足够人数的性别池
            let selectedPool = null;
            if (maleEmployees.length >= roomCount && femaleEmployees.length >= roomCount) {
                // 两边都够，随机选择
                selectedPool = Math.random() < 0.5 ? maleEmployees : femaleEmployees;
            } else if (maleEmployees.length >= roomCount) {
                selectedPool = maleEmployees;
            } else if (femaleEmployees.length >= roomCount) {
                selectedPool = femaleEmployees;
            }

            // 如果没有足够的同性别人员，将这些房间转为单独房间处理
            if (!selectedPool) {
                for (const roomInfo of roomGroup) {
                    unassignedSharedRooms.push(roomInfo);
                }
                continue;
            }

            // 从选定的池中取出人员分配到这组房间
            for (const roomInfo of roomGroup) {
                const employee = selectedPool.shift();
                result[roomInfo.originalIndex] = {
                    employeeId: employee.employeeId,
                    gender: employee.gender,
                    roomNumber: roomInfo.fullRoom
                };
            }
        }

        // 合并剩余人员
        const remainingEmployees = [...maleEmployees, ...femaleEmployees];
        this.shuffleArray(remainingEmployees);

        // 合并未分配的共享房间到单独房间
        const allSingleRooms = [...singleRooms, ...unassignedSharedRooms];
        this.shuffleArray(allSingleRooms);

        // 分配单独房间（允许房间多于人数，多余房间留空）
        for (const roomInfo of allSingleRooms) {
            if (remainingEmployees.length > 0) {
                const employee = remainingEmployees.shift();
                result[roomInfo.originalIndex] = {
                    employeeId: employee.employeeId,
                    gender: employee.gender,
                    roomNumber: roomInfo.fullRoom
                };
            }
            // 如果没有剩余人员，房间保持未分配（result中为undefined）
        }

        // 过滤掉未分配的房间
        const finalResult = result.filter(r => r !== undefined && r !== null);

        // 随机打乱最终显示顺序
        this.shuffleArray(finalResult);

        return finalResult;
    },

    /**
     * Fisher-Yates 洗牌算法
     * @param {Array} array - 要打乱的数组
     */
    shuffleArray(array) {
        for (let i = array.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [array[i], array[j]] = [array[j], array[i]];
        }
    },

    /**
     * 开始滚动动画
     * @param {Function} updateCallback - 更新UI的回调函数
     */
    startRolling(updateCallback) {
        if (this.isRolling) return;

        this.isRolling = true;
        const data = this.originalData;

        // 每50ms更新一次显示
        this.rollInterval = setInterval(() => {
            // 生成随机显示数据
            const randomDisplay = data.map((item, index) => {
                const randomEmpIndex = Math.floor(Math.random() * data.length);
                const randomRoomIndex = Math.floor(Math.random() * data.length);

                return {
                    employeeId: data[randomEmpIndex].employeeId,
                    gender: data[randomEmpIndex].gender,
                    roomNumber: data[randomRoomIndex].roomNumber
                };
            });

            updateCallback(randomDisplay, true);
        }, 50);
    },

    /**
     * 停止滚动并执行真正的摇号
     * @param {Function} updateCallback - 更新UI的回调函数
     * @returns {Array} - 最终结果
     */
    stopRolling(updateCallback) {
        if (!this.isRolling) return this.currentData;

        this.isRolling = false;

        if (this.rollInterval) {
            clearInterval(this.rollInterval);
            this.rollInterval = null;
        }

        // 执行真正的摇号算法
        const result = this.performLottery();
        this.currentData = result;

        updateCallback(result, false);

        return result;
    },

    /**
     * 获取当前数据
     * @returns {Array}
     */
    getCurrentData() {
        return this.currentData;
    }
};

// 导出模块
window.LotteryModule = LotteryModule;
