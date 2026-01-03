// 礼簿应用主逻辑 - Flask API版本
class GiftBookApp {
    constructor() {
        this.records = [];
        this.currentPage = 1;
        this.recordsPerPage = 12;
        this.editingIndex = -1;
        this.viewMode = 'traditional'; // traditional 或 edit
        this.showAmountInName = false; // 是否在姓名框显示金额

        this.init();
    }

    async init() {
        // 从服务器加载数据
        await this.loadData();
        this.loadSettings();

        // 绑定事件
        this.bindEvents();

        // 渲染页面
        this.render();
    }

    bindEvents() {
        // 表单提交
        document.getElementById('guestForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            await this.addRecord();
        });

        // 打开Excel文件
        document.getElementById('openFileBtn').addEventListener('click', async () => {
            await this.loadExcelFile();
        });

        // 分页控制
        document.getElementById('prevPage').addEventListener('click', () => {
            if (this.currentPage > 1) {
                this.currentPage--;
                this.render();
            }
        });

        document.getElementById('nextPage').addEventListener('click', () => {
            const totalPages = this.getTotalPages();
            if (this.currentPage < totalPages) {
                this.currentPage++;
                this.render();
            }
        });

        // 视图模式切换
        document.getElementById('editModeBtn').addEventListener('click', () => {
            this.toggleViewMode();
        });

        // 封面显示/隐藏
        document.getElementById('toggleCoverBtn').addEventListener('click', () => {
            this.toggleCover();
        });

        // 统计信息展开/收起
        document.getElementById('toggleSummaryBtn').addEventListener('click', () => {
            this.toggleSummary();
        });

        // 每页汇总展开/收起
        document.getElementById('togglePageTotalsBtn').addEventListener('click', () => {
            this.togglePageTotals();
        });

        // 切换姓名框中金额显示
        document.getElementById('toggleAmountInNameBtn').addEventListener('click', () => {
            this.toggleAmountInName();
        });

        // 导出Excel
        document.getElementById('exportBtn').addEventListener('click', async () => {
            await this.exportToExcel();
        });

        // 导出打印页
        document.getElementById('exportPrintBtn').addEventListener('click', () => {
            this.exportPrintPage();
        });

        // 导入数据
        document.getElementById('importBtn').addEventListener('click', () => {
            this.toggleImportPanel();
        });

        document.getElementById('confirmImport').addEventListener('click', async () => {
            await this.importData();
        });

        document.getElementById('closeImport').addEventListener('click', () => {
            this.toggleImportPanel();
        });

        // 设置面板
        document.getElementById('settingsBtn').addEventListener('click', () => {
            this.toggleSettings();
        });

        document.getElementById('saveSettings').addEventListener('click', () => {
            this.saveSettings();
        });

        document.getElementById('closeSettings').addEventListener('click', () => {
            this.toggleSettings();
        });
    }

    // 数据管理 - 使用API
    async loadData() {
        try {
            const response = await fetch('/api/records');
            const data = await response.json();
            this.records = data;
            // 跳转到最后一页
            this.currentPage = Math.ceil(this.records.length / this.recordsPerPage) || 1;
        } catch (error) {
            console.error('加载数据失败：', error);
            alert('加载数据失败：' + error.message);
        }
    }

    async saveData() {
        try {
            const response = await fetch('/api/records', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(this.records)
            });
            const result = await response.json();
            if (!result.success) {
                console.error('保存失败：', result.message);
            }
        } catch (error) {
            console.error('保存数据失败：', error);
        }
    }

    async loadExcelFile() {
        try {
            const response = await fetch('/api/excel/load');
            const result = await response.json();

            if (result.success) {
                this.records = result.records;
                this.currentPage = Math.ceil(this.records.length / this.recordsPerPage) || 1;
                this.render();
                alert(`成功从Excel加载 ${this.records.length} 条记录！`);
            } else {
                alert('加载失败：' + result.message);
            }
        } catch (error) {
            console.error('加载Excel失败：', error);
            alert('加载Excel失败：' + error.message);
        }
    }

    loadSettings() {
        const saved = localStorage.getItem('giftBookSettings');
        if (saved) {
            const settings = JSON.parse(saved);
            this.recordsPerPage = settings.recordsPerPage || 12;
            document.getElementById('recordsPerPage').value = this.recordsPerPage;
        }
    }

    saveSettings() {
        const recordsPerPage = parseInt(document.getElementById('recordsPerPage').value);

        if (recordsPerPage > 0 && recordsPerPage <= 50) {
            this.recordsPerPage = recordsPerPage;
            localStorage.setItem('giftBookSettings', JSON.stringify({
                recordsPerPage: this.recordsPerPage
            }));
            this.currentPage = 1;
            this.render();
            this.toggleSettings();
            alert('设置已保存！');
        } else {
            alert('每页记录数必须在1-50之间！');
        }
    }

    toggleSettings() {
        const panel = document.getElementById('settingsPanel');
        panel.classList.toggle('hidden');
    }

    toggleImportPanel() {
        const panel = document.getElementById('importPanel');
        panel.classList.toggle('hidden');
    }

    toggleViewMode() {
        if (this.viewMode === 'traditional') {
            this.viewMode = 'edit';
            document.getElementById('traditionalView').classList.add('hidden');
            document.getElementById('editView').classList.remove('hidden');
            document.getElementById('editModeBtn').textContent = '传统视图';
            document.getElementById('toggleCoverBtn').classList.add('hidden');
            document.getElementById('togglePageTotalsBtn').classList.add('hidden');
            document.getElementById('toggleAmountInNameBtn').classList.add('hidden');
        } else {
            this.viewMode = 'traditional';
            document.getElementById('traditionalView').classList.remove('hidden');
            document.getElementById('editView').classList.add('hidden');
            document.getElementById('editModeBtn').textContent = '编辑模式';
            document.getElementById('toggleCoverBtn').classList.remove('hidden');
            document.getElementById('togglePageTotalsBtn').classList.remove('hidden');
            document.getElementById('toggleAmountInNameBtn').classList.remove('hidden');
        }
        this.render();
    }

    toggleCover() {
        const cover = document.getElementById('bookCover');
        const title = document.getElementById('bookTitle');
        const btn = document.getElementById('toggleCoverBtn');

        if (cover.classList.contains('hidden')) {
            cover.classList.remove('hidden');
            title.classList.add('hidden');
            btn.textContent = '隐藏封面';
        } else {
            cover.classList.add('hidden');
            title.classList.remove('hidden');
            btn.textContent = '显示封面';
        }
    }

    toggleSummary() {
        const summary = document.getElementById('summarySection');
        const btn = document.getElementById('toggleSummaryBtn');

        if (summary.classList.contains('hidden')) {
            summary.classList.remove('hidden');
            btn.textContent = '收起汇总信息';
        } else {
            summary.classList.add('hidden');
            btn.textContent = '展开汇总信息';
        }
    }

    togglePageTotals() {
        const pageTotals = document.getElementById('pageTotals');
        const btn = document.getElementById('togglePageTotalsBtn');

        if (pageTotals.classList.contains('hidden')) {
            pageTotals.classList.remove('hidden');
            btn.textContent = '收起本页汇总';
        } else {
            pageTotals.classList.add('hidden');
            btn.textContent = '展开本页汇总';
        }
    }

    toggleAmountInName() {
        const btn = document.getElementById('toggleAmountInNameBtn');
        this.showAmountInName = !this.showAmountInName;

        if (this.showAmountInName) {
            btn.textContent = '隐藏金额';
        } else {
            btn.textContent = '显示金额';
        }

        this.render();
    }

    async addRecord() {
        const name = document.getElementById('guestName').value.trim();
        const amount = parseFloat(document.getElementById('giftAmount').value);
        const paymentMethod = document.getElementById('paymentMethod').value;

        if (!name || !amount || !paymentMethod) {
            alert('请填写完整信息！');
            return;
        }

        // 检查姓名是否重复（编辑模式除外）
        if (this.editingIndex < 0) {
            const duplicateName = this.records.find(record => record.name === name);
            if (duplicateName) {
                if (!confirm(`姓名"${name}"已存在（礼金：${duplicateName.amount}元），是否继续添加？`)) {
                    return;
                }
            }
        }

        const record = {
            id: Date.now(),
            name: name,
            amount: amount,
            amountChinese: this.numberToChinese(amount),
            paymentMethod: paymentMethod,
            timestamp: new Date().toISOString()
        };

        if (this.editingIndex >= 0) {
            // 编辑模式
            this.records[this.editingIndex] = record;
            this.editingIndex = -1;
            document.querySelector('#guestForm button[type="submit"]').textContent = '添加记录';
        } else {
            // 新增模式
            this.records.push(record);
            // 跳转到最后一页显示新记录
            this.currentPage = Math.ceil(this.records.length / this.recordsPerPage);
        }

        await this.saveData(); // 自动保存到服务器
        this.render();
        this.clearForm();
    }

    editRecord(index) {
        const record = this.records[index];
        document.getElementById('guestName').value = record.name;
        document.getElementById('giftAmount').value = record.amount;
        document.getElementById('paymentMethod').value = record.paymentMethod;

        this.editingIndex = index;
        document.querySelector('#guestForm button[type="submit"]').textContent = '保存修改';

        // 滚动到表单
        document.querySelector('.input-section').scrollIntoView({ behavior: 'smooth' });
    }

    async deleteRecord(index) {
        if (confirm('确定要删除这条记录吗？')) {
            this.records.splice(index, 1);
            await this.saveData(); // 自动保存到服务器

            // 如果当前页没有记录了，返回上一页
            const totalPages = this.getTotalPages();
            if (this.currentPage > totalPages && this.currentPage > 1) {
                this.currentPage = totalPages;
            }

            this.render();
        }
    }

    clearForm() {
        document.getElementById('guestForm').reset();
        this.editingIndex = -1;
        document.querySelector('#guestForm button[type="submit"]').textContent = '添加记录';
    }

    // 导入数据
    async importData() {
        const fileInput = document.getElementById('importFile');
        const file = fileInput.files[0];

        if (!file) {
            alert('请先选择要导入的文件！');
            return;
        }

        try {
            const formData = new FormData();
            formData.append('file', file);

            const response = await fetch('/api/import', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                this.records = result.records;
                this.currentPage = Math.ceil(this.records.length / this.recordsPerPage) || 1;
                this.render();
                this.toggleImportPanel();
                alert(`成功导入 ${this.records.length} 条记录！`);
                fileInput.value = '';
            } else {
                alert('导入失败：' + result.message);
            }
        } catch (error) {
            console.error('导入失败：', error);
            alert('导入失败：' + error.message);
        }
    }

    // 分页计算
    getTotalPages() {
        return Math.ceil(this.records.length / this.recordsPerPage) || 1;
    }

    getCurrentPageRecords() {
        const start = (this.currentPage - 1) * this.recordsPerPage;
        const end = start + this.recordsPerPage;
        return this.records.slice(start, end);
    }

    // 数字转中文大写
    numberToChinese(num) {
        if (num === 0) return '零元整';

        const digits = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
        const units = ['', '拾', '佰', '仟'];
        const bigUnits = ['', '万', '亿'];

        // 分离整数和小数部分
        const parts = num.toFixed(2).split('.');
        let integerPart = parseInt(parts[0]);
        const decimalPart = parts[1];

        if (integerPart === 0) {
            const jiao = parseInt(decimalPart[0]);
            const fen = parseInt(decimalPart[1]);
            let result = '';
            if (jiao > 0) result += digits[jiao] + '角';
            if (fen > 0) result += digits[fen] + '分';
            return result || '零元整';
        }

        // 处理整数部分
        let result = '';
        let unitIndex = 0;
        let zeroCount = 0;

        while (integerPart > 0) {
            const section = integerPart % 10000;
            if (section !== 0) {
                let sectionStr = '';
                let sectionNum = section;
                let pos = 0;

                while (sectionNum > 0) {
                    const digit = sectionNum % 10;
                    if (digit === 0) {
                        zeroCount++;
                    } else {
                        if (zeroCount > 0) {
                            sectionStr = '零' + sectionStr;
                            zeroCount = 0;
                        }
                        sectionStr = digits[digit] + units[pos] + sectionStr;
                    }
                    sectionNum = Math.floor(sectionNum / 10);
                    pos++;
                }

                result = sectionStr + bigUnits[unitIndex] + result;
            } else if (result !== '') {
                zeroCount++;
            }

            integerPart = Math.floor(integerPart / 10000);
            unitIndex++;
        }

        if (result.endsWith('零')) {
            result = result.slice(0, -1);
        }

        result += '元';

        // 处理小数部分
        const jiao = parseInt(decimalPart[0]);
        const fen = parseInt(decimalPart[1]);

        if (jiao === 0 && fen === 0) {
            result += '整';
        } else {
            if (jiao > 0) {
                result += digits[jiao] + '角';
            }
            if (fen > 0) {
                if (jiao === 0) {
                    result += '零';
                }
                result += digits[fen] + '分';
            }
        }

        return result;
    }

    // 计算总金额
    getTotalAmount() {
        return this.records.reduce((sum, record) => sum + record.amount, 0);
    }

    // 计算页面统计
    getPageStats(records) {
        let mobileTotal = 0;
        let cashTotal = 0;
        let pageTotal = 0;

        records.forEach(record => {
            pageTotal += record.amount;
            if (record.paymentMethod === '现金') {
                cashTotal += record.amount;
            } else if (['微信', '支付宝'].includes(record.paymentMethod)) {
                mobileTotal += record.amount;
            }
        });

        return { mobileTotal, cashTotal, pageTotal };
    }

    // 渲染页面（保持原有的render逻辑不变）
    render() {
        if (this.viewMode === 'traditional') {
            this.renderTraditionalView();
        } else {
            this.renderEditView();
        }
        this.renderSummary();
        this.renderPagination();
    }

    // 渲染传统视图（保持原有逻辑）
    renderTraditionalView() {
        const currentRecords = this.getCurrentPageRecords();
        const namesRow = document.getElementById('namesRow');
        const amountsRow = document.getElementById('amountsRow');

        namesRow.innerHTML = '';
        amountsRow.innerHTML = '';

        for (let i = 0; i < this.recordsPerPage; i++) {
            const record = currentRecords[i];

            // 姓名竖栏
            const nameDiv = document.createElement('div');
            nameDiv.className = 'record-item-vertical' + (record ? '' : ' empty');

            if (record) {
                const match = record.name.match(/^([^()（）]+)[()（]([^()（）]+)[)）]$/);
                if (match) {
                    const mainSpan = document.createElement('span');
                    mainSpan.className = 'name-main';
                    mainSpan.textContent = match[1];

                    const noteSpan = document.createElement('span');
                    noteSpan.className = 'name-note';
                    noteSpan.textContent = `(${match[2]})`;

                    nameDiv.appendChild(mainSpan);
                    nameDiv.appendChild(noteSpan);
                } else {
                    const mainSpan = document.createElement('span');
                    mainSpan.className = 'name-main';
                    mainSpan.textContent = record.name;
                    nameDiv.appendChild(mainSpan);
                }
            }

            namesRow.appendChild(nameDiv);

            // 金额竖栏
            const amountDiv = document.createElement('div');
            amountDiv.className = 'record-item-vertical' + (record ? '' : ' empty');

            if (record) {
                const chineseSpan = document.createElement('span');
                chineseSpan.className = 'amount-chinese';
                chineseSpan.textContent = record.amountChinese;
                amountDiv.appendChild(chineseSpan);

                if (this.showAmountInName) {
                    const numberSpan = document.createElement('span');
                    numberSpan.className = 'amount-number';
                    numberSpan.textContent = `¥${record.amount.toFixed(2)}`;
                    amountDiv.appendChild(numberSpan);
                }
            }

            // 添加支付方式图标
            if (record && (record.paymentMethod === '微信' || record.paymentMethod === '支付宝')) {
                const icon = document.createElement('img');
                icon.className = 'payment-icon';
                if (record.paymentMethod === '微信') {
                    icon.src = '/static/weixin.png';
                    icon.alt = '微信';
                } else if (record.paymentMethod === '支付宝') {
                    icon.src = '/static/favicon.ico';
                    icon.alt = '支付宝';
                }
                amountDiv.appendChild(icon);
            }

            amountsRow.appendChild(amountDiv);
        }

        document.getElementById('currentPageNum').textContent = this.currentPage;

        const stats = this.getPageStats(currentRecords);
        document.getElementById('mobileTotalValue').textContent = stats.mobileTotal.toFixed(2) + ' 元';
        document.getElementById('mobileTotalChinese').textContent = this.numberToChinese(stats.mobileTotal);
        document.getElementById('cashTotalValue').textContent = stats.cashTotal.toFixed(2) + ' 元';
        document.getElementById('cashTotalChinese').textContent = this.numberToChinese(stats.cashTotal);
        document.getElementById('pageTotalValue').textContent = stats.pageTotal.toFixed(2) + ' 元';
        document.getElementById('pageTotalChinese').textContent = this.numberToChinese(stats.pageTotal);
    }

    // 渲染编辑视图（保持原有逻辑）
    renderEditView() {
        const tbody = document.getElementById('recordsBody');
        const currentRecords = this.getCurrentPageRecords();

        tbody.innerHTML = '';

        const startIndex = (this.currentPage - 1) * this.recordsPerPage;

        currentRecords.forEach((record, idx) => {
            const globalIndex = startIndex + idx;
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${globalIndex + 1}</td>
                <td>${record.name}</td>
                <td>${record.amount.toFixed(2)}</td>
                <td>${record.amountChinese}</td>
                <td>${record.paymentMethod}</td>
                <td>
                    <button class="btn btn-edit" onclick="app.editRecord(${globalIndex})">编辑</button>
                    <button class="btn btn-danger" onclick="app.deleteRecord(${globalIndex})">删除</button>
                </td>
            `;
            tbody.appendChild(row);
        });

        const pageTotal = currentRecords.reduce((sum, record) => sum + record.amount, 0);
        document.getElementById('pageTotal').textContent = pageTotal.toFixed(2);
        document.getElementById('pageTotalChinese2').textContent = this.numberToChinese(pageTotal);

        document.getElementById('pageTitle').textContent = `(第 ${this.currentPage} 页)`;
    }

    renderSummary() {
        const totalRecords = this.records.length;
        const totalAmount = this.getTotalAmount();

        let weixinTotal = 0;
        let cashTotal = 0;
        let alipayTotal = 0;

        this.records.forEach(record => {
            if (record.paymentMethod === '微信') {
                weixinTotal += record.amount;
            } else if (record.paymentMethod === '现金') {
                cashTotal += record.amount;
            } else if (record.paymentMethod === '支付宝') {
                alipayTotal += record.amount;
            }
        });

        document.getElementById('totalRecords').textContent = totalRecords;
        document.getElementById('totalWeixinAmount').textContent = weixinTotal.toFixed(2);
        document.getElementById('totalCashAmount').textContent = cashTotal.toFixed(2);
        document.getElementById('totalAmount').textContent = totalAmount.toFixed(2);
        document.getElementById('totalAmountChinese').textContent = this.numberToChinese(totalAmount);
    }

    renderPagination() {
        const totalPages = this.getTotalPages();
        document.getElementById('pageInfo').textContent = `第 ${this.currentPage} 页，共 ${totalPages} 页`;

        document.getElementById('prevPage').disabled = this.currentPage === 1;
        document.getElementById('nextPage').disabled = this.currentPage === totalPages;
    }

    // 导出Excel
    async exportToExcel() {
        if (this.records.length === 0) {
            alert('没有数据可导出！');
            return;
        }

        try {
            const response = await fetch('/api/excel/export', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(this.records)
            });

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = '礼簿_导出.xlsx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            alert('导出成功！');
        } catch (error) {
            console.error('导出失败：', error);
            alert('导出失败：' + error.message);
        }
    }

    // 导出打印页面
    exportPrintPage() {
        if (this.records.length === 0) {
            alert('没有数据可导出！');
            return;
        }

        const totalPages = Math.ceil(this.records.length / this.recordsPerPage);
        const printWindow = window.open('', '_blank');

        let html = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>礼簿打印 - ${new Date().toLocaleDateString('zh-CN')}</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'STKaiti', 'KaiTi', 'Microsoft YaHei', serif; background: white; padding: 20px; }
        .page { page-break-after: always; margin-bottom: 40px; background: #FFF9E6; border: 4px solid #D32F2F; border-radius: 10px; padding: 30px; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .page:last-child { page-break-after: auto; }
        .book-title { text-align: center; font-size: 28px; font-weight: bold; color: #B71C1C; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 3px double #D32F2F; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .book-content { display: flex; flex-direction: column; gap: 0; }
        .names-row, .amounts-row { display: flex; flex-wrap: nowrap; gap: 10px; padding: 20px; background: white; }
        .names-row { border-bottom: none; }
        .amounts-row { border-top: none; }
        .record-item-vertical { flex: 1; min-width: 50px; padding: 25px 12px; background: #FFF9E6; border: 2px solid #FFD700; border-radius: 8px; min-height: 250px; display: flex; align-items: center; justify-content: center; writing-mode: vertical-rl; text-orientation: upright; font-size: 28px; color: #333; font-weight: 600; letter-spacing: 12px; position: relative; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .record-item-vertical.empty { background: #FAFAFA; border: 1px dashed #CCC; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .name-note { position: absolute; bottom: 8px; right: 8px; font-size: 14px; color: #666; font-weight: normal; writing-mode: horizontal-tb; text-orientation: mixed; letter-spacing: 0; background: rgba(255, 255, 255, 0.9); padding: 2px 4px; border-radius: 3px; white-space: nowrap; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .amount-chinese { display: inline; }
        .amount-number { position: absolute; bottom: 8px; left: 8px; font-size: 12px; color: #D32F2F; font-weight: bold; writing-mode: horizontal-tb; text-orientation: mixed; letter-spacing: 0; background: rgba(255, 249, 230, 0.95); padding: 2px 6px; border-radius: 3px; box-shadow: 0 1px 3px rgba(211, 47, 47, 0.2); border: 1px solid #FFD700; white-space: nowrap; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .payment-icon { position: absolute; top: 8px; right: 8px; width: 28px; height: 28px; object-fit: contain; }
        .divider-row { display: flex; justify-content: center; align-items: center; position: relative; height: 80px; background: white; padding: 10px; }
        .horizontal-divider { height: 4px; width: 100%; background: linear-gradient(90deg, #D32F2F 0%, #E91E63 50%, #D32F2F 100%); border-radius: 2px; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .divider-text-horizontal { position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 15px 30px; font-size: 32px; font-weight: bold; color: #E91E63; border: 3px solid #D32F2F; border-radius: 50px; letter-spacing: 20px; padding-left: 40px; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .page-totals { margin-top: 30px; padding: 20px; background: linear-gradient(135deg, #FFF3E0 0%, #FFE0B2 100%); border: 3px solid #D32F2F; border-radius: 10px; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .totals-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }
        .total-item { text-align: center; padding: 15px; background: white; border: 2px solid #FFD700; border-radius: 8px; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .total-label { font-size: 16px; color: #8B4513; font-weight: bold; margin-bottom: 10px; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .total-value { font-size: 22px; color: #B71C1C; font-weight: bold; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        .total-chinese { font-size: 14px; color: #666; margin-top: 5px; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        @media print {
            body { padding: 0; margin: 0; }
            .page { margin-bottom: 0; border: 2px solid #000; padding: 15px; border-radius: 0; }
            .book-title { font-size: 20px; margin-bottom: 10px; padding-bottom: 8px; }
            .names-row, .amounts-row { gap: 6px; padding: 10px; }
            .record-item-vertical { min-height: 180px; padding: 15px 8px; font-size: 22px; letter-spacing: 8px; }
            .divider-row { height: 50px; padding: 5px; }
            .divider-text-horizontal { font-size: 24px; padding: 10px 20px; letter-spacing: 15px; padding-left: 30px; }
            .page-totals { margin-top: 15px; padding: 12px; }
            .totals-grid { gap: 12px; }
            .total-item { padding: 10px; }
            .total-label { font-size: 13px; margin-bottom: 6px; }
            .total-value { font-size: 18px; }
            .total-chinese { font-size: 12px; }
            .name-note { font-size: 11px; bottom: 6px; right: 6px; }
            .amount-number { font-size: 10px; bottom: 6px; left: 6px; padding: 1px 4px; }
            .payment-icon { width: 22px; height: 22px; top: 6px; right: 6px; }
        }
    </style>
</head>
<body>
`;

        for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
            const startIndex = (pageNum - 1) * this.recordsPerPage;
            const endIndex = Math.min(startIndex + this.recordsPerPage, this.records.length);
            const pageRecords = this.records.slice(startIndex, endIndex);

            html += `
    <div class="page">
        <div class="book-title">人情簿 - 第 ${pageNum} 页</div>
        <div class="book-content">
            <div class="names-row">
`;

            for (let i = 0; i < this.recordsPerPage; i++) {
                const record = pageRecords[i];
                if (record) {
                    const match = record.name.match(/^([^()（）]+)[()（]([^()（）]+)[)）]$/);
                    if (match) {
                        html += `
                <div class="record-item-vertical">
                    <span class="name-main">${match[1]}</span>
                    <span class="name-note">(${match[2]})</span>
                </div>
`;
                    } else {
                        html += `
                <div class="record-item-vertical">${record.name}</div>
`;
                    }
                } else {
                    html += `
                <div class="record-item-vertical empty"></div>
`;
                }
            }

            html += `
            </div>
            <div class="divider-row">
                <div class="horizontal-divider"></div>
                <div class="divider-text-horizontal">礼金</div>
            </div>
            <div class="amounts-row">
`;

            for (let i = 0; i < this.recordsPerPage; i++) {
                const record = pageRecords[i];
                if (record) {
                    let paymentIcon = '';
                    if (record.paymentMethod === '微信') {
                        paymentIcon = '<img src="static/weixin.png" alt="微信" class="payment-icon" />';
                    } else if (record.paymentMethod === '支付宝') {
                        paymentIcon = '<img src="static/favicon.ico" alt="支付宝" class="payment-icon" />';
                    }

                    html += `
                <div class="record-item-vertical">
                    <span class="amount-chinese">${record.amountChinese}</span>
                    <span class="amount-number">¥${record.amount.toFixed(2)}</span>
                    ${paymentIcon}
                </div>
`;
                } else {
                    html += `
                <div class="record-item-vertical empty"></div>
`;
                }
            }

            html += `
            </div>
        </div>
        <div class="page-totals">
            <div class="totals-grid">
`;

            const stats = this.getPageStats(pageRecords);
            html += `
                <div class="total-item">
                    <div class="total-label">移动支付总计</div>
                    <div class="total-value">${stats.mobileTotal.toFixed(2)} 元</div>
                    <div class="total-chinese">${this.numberToChinese(stats.mobileTotal)}</div>
                </div>
                <div class="total-item">
                    <div class="total-label">现金支付总计</div>
                    <div class="total-value">${stats.cashTotal.toFixed(2)} 元</div>
                    <div class="total-chinese">${this.numberToChinese(stats.cashTotal)}</div>
                </div>
                <div class="total-item">
                    <div class="total-label">本页总计</div>
                    <div class="total-value">${stats.pageTotal.toFixed(2)} 元</div>
                    <div class="total-chinese">${this.numberToChinese(stats.pageTotal)}</div>
                </div>
`;

            html += `
            </div>
        </div>
    </div>
`;
        }

        const totalAmount = this.getTotalAmount();
        let weixinTotal = 0, cashTotal = 0, alipayTotal = 0;
        this.records.forEach(record => {
            if (record.paymentMethod === '微信') weixinTotal += record.amount;
            else if (record.paymentMethod === '现金') cashTotal += record.amount;
            else if (record.paymentMethod === '支付宝') alipayTotal += record.amount;
        });

        html += `
    <div class="page">
        <div class="book-title">总汇总</div>
        <div class="page-totals">
            <div class="totals-grid">
                <div class="total-item">
                    <div class="total-label">总记录数</div>
                    <div class="total-value">${this.records.length} 条</div>
                </div>
                <div class="total-item">
                    <div class="total-label">微信总金额</div>
                    <div class="total-value">${weixinTotal.toFixed(2)} 元</div>
                    <div class="total-chinese">${this.numberToChinese(weixinTotal)}</div>
                </div>
                <div class="total-item">
                    <div class="total-label">现金总金额</div>
                    <div class="total-value">${cashTotal.toFixed(2)} 元</div>
                    <div class="total-chinese">${this.numberToChinese(cashTotal)}</div>
                </div>
                <div class="total-item">
                    <div class="total-label">支付宝总金额</div>
                    <div class="total-value">${alipayTotal.toFixed(2)} 元</div>
                    <div class="total-chinese">${this.numberToChinese(alipayTotal)}</div>
                </div>
                <div class="total-item">
                    <div class="total-label">总金额</div>
                    <div class="total-value">${totalAmount.toFixed(2)} 元</div>
                    <div class="total-chinese">${this.numberToChinese(totalAmount)}</div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
`;

        printWindow.document.write(html);
        printWindow.document.close();
        printWindow.onload = function() {
            printWindow.print();
        };
    }
}

// 初始化应用
let app;
document.addEventListener('DOMContentLoaded', () => {
    app = new GiftBookApp();
});
