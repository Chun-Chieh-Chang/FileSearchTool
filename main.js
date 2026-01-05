/**
 * FileSearchTool v1.6.0 - Web Version
 * Core Logic for Searching Excel and PDF files
 */

// Initialize PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

class WebFileSearchTool {
    constructor() {
        this.selectedFiles = [];
        this.stopSearchFlag = false;
        this.results = [];
        this.initEventListeners();
    }

    initEventListeners() {
        const folderInput = document.getElementById('folderInput');
        const fileInput = document.getElementById('fileInput');
        const startBtn = document.getElementById('startSearch');
        const stopBtn = document.getElementById('stopSearch');

        folderInput.addEventListener('change', (e) => this.handleFileSelection(e));
        fileInput.addEventListener('change', (e) => this.handleFileSelection(e));
        startBtn.addEventListener('click', () => this.startSearch());
        stopBtn.addEventListener('click', () => this.stopSearch());
    }

    handleFileSelection(event) {
        const files = Array.from(event.target.files);
        // Filter by extensions
        this.selectedFiles = files.filter(file => {
            const ext = file.name.split('.').pop().toLowerCase();
            return ['xlsx', 'xls', 'pdf'].includes(ext);
        });

        document.getElementById('fileStats').innerText = `已選擇 ${this.selectedFiles.length} 個符合的檔案`;
    }

    async startSearch() {
        const kw1 = document.getElementById('keyword1').value.trim();
        const kw2 = document.getElementById('keyword2').value.trim();
        const logic = document.querySelector('input[name="keywordLogic"]:checked').value;
        const wholeWord = document.getElementById('wholeWord').checked;
        const caseSensitive = document.getElementById('caseSensitive').checked;
        const typeFilter = document.getElementById('typeFilter').value;

        if (!kw1) {
            alert('請至少輸入關鍵字 1');
            return;
        }

        if (this.selectedFiles.length === 0) {
            alert('請先選擇資料夾或檔案');
            return;
        }

        // Reset UI
        this.stopSearchFlag = false;
        this.results = [];
        const tableBody = document.querySelector('#resultsTable tbody');
        tableBody.innerHTML = '';

        const progressSection = document.getElementById('progressSection');
        const progressBar = document.getElementById('progressBar');
        const progressLabel = document.getElementById('progressLabel');
        const statusLabel = document.getElementById('statusLabel');
        const startBtn = document.getElementById('startSearch');
        const stopBtn = document.getElementById('stopSearch');

        progressSection.style.display = 'block';
        startBtn.disabled = true;
        stopBtn.disabled = false;

        let processedCount = 0;
        const totalFiles = this.selectedFiles.length;

        for (const file of this.selectedFiles) {
            if (this.stopSearchFlag) break;

            const ext = file.name.split('.').pop().toLowerCase();

            // Filter by type
            if (typeFilter === 'Excel' && !['xlsx', 'xls'].includes(ext)) {
                processedCount++;
                continue;
            }
            if (typeFilter === 'PDF' && ext !== 'pdf') {
                processedCount++;
                continue;
            }

            statusLabel.innerText = `正在處理: ${file.name}`;

            try {
                const matchData = await this.searchInFile(file, kw1, kw2, logic, wholeWord, caseSensitive);
                if (matchData.isMatch) {
                    this.addResultToTable(file, matchData.totalCount, matchData.type, matchData.location);
                }
            } catch (err) {
                console.error(`Error processing ${file.name}:`, err);
            }

            processedCount++;
            const percent = Math.round((processedCount / totalFiles) * 100);
            progressBar.style.width = `${percent}%`;
            progressLabel.innerText = `進度: ${percent}% (${processedCount}/${totalFiles})`;
        }

        statusLabel.innerText = this.stopSearchFlag ? '搜尋已停止' : '搜尋完成';
        startBtn.disabled = false;
        stopBtn.disabled = true;
    }

    stopSearch() {
        this.stopSearchFlag = true;
    }

    async searchInFile(file, kw1, kw2, logic, wholeWord, caseSensitive) {
        const ext = file.name.split('.').pop().toLowerCase();
        let textContent = '';
        let location = '內容中';
        let matchData = { isMatch: false, totalCount: 0, type: '', location: '' };

        if (['xlsx', 'xls'].includes(ext)) {
            matchData.type = 'Excel';
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });

            let kw1Found = false;
            let kw2Found = false;
            let kw1Count = 0;
            let kw2Count = 0;
            let firstLoc = '';

            for (const sheetName of workbook.SheetNames) {
                const sheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

                rows.forEach((row, rIdx) => {
                    row.forEach((cell, cIdx) => {
                        const cellStr = String(cell);
                        const m1 = this.countMatches(cellStr, kw1, wholeWord, caseSensitive);
                        if (m1 > 0) {
                            kw1Found = true;
                            kw1Count += m1;
                            if (!firstLoc) firstLoc = `Sheet: ${sheetName}, Cell: ${this.getColLetter(cIdx)}${rIdx + 1}`;
                        }

                        if (kw2) {
                            const m2 = this.countMatches(cellStr, kw2, wholeWord, caseSensitive);
                            if (m2 > 0) {
                                kw2Found = true;
                                kw2Count += m2;
                                if (!firstLoc) firstLoc = `Sheet: ${sheetName}, Cell: ${this.getColLetter(cIdx)}${rIdx + 1}`;
                            }
                        }
                    });
                });
            }

            matchData.isMatch = this.checkLogic(kw1Found, kw2Found, kw2, logic);
            matchData.totalCount = kw1Count + kw2Count;
            matchData.location = firstLoc || '未知位置';

        } else if (ext === 'pdf') {
            matchData.type = 'PDF';
            const data = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data }).promise;

            let kw1Found = false;
            let kw2Found = false;
            let kw1Count = 0;
            let kw2Count = 0;
            let firstLoc = '';

            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const textDetail = await page.getTextContent();
                const pageText = textDetail.items.map(item => item.str).join(' ');

                const m1 = this.countMatches(pageText, kw1, wholeWord, caseSensitive);
                if (m1 > 0) {
                    kw1Found = true;
                    kw1Count += m1;
                    if (!firstLoc) firstLoc = `Page ${i}`;
                }

                if (kw2) {
                    const m2 = this.countMatches(pageText, kw2, wholeWord, caseSensitive);
                    if (m2 > 0) {
                        kw2Found = true;
                        kw2Count += m2;
                        if (!firstLoc) firstLoc = `Page ${i}`;
                    }
                }
            }

            matchData.isMatch = this.checkLogic(kw1Found, kw2Found, kw2, logic);
            matchData.totalCount = kw1Count + kw2Count;
            matchData.location = firstLoc || '未知位置';
        }

        return matchData;
    }

    countMatches(text, keyword, wholeWord, caseSensitive) {
        if (!keyword) return 0;
        let flags = 'g';
        if (!caseSensitive) flags += 'i';

        let pattern = this.escapeRegExp(keyword);
        if (wholeWord) {
            pattern = `\\b${pattern}\\b`;
        }

        const regex = new RegExp(pattern, flags);
        const matches = text.match(regex);
        return matches ? matches.length : 0;
    }

    checkLogic(found1, found2, kw2Active, logic) {
        if (!kw2Active) return found1;
        if (logic === 'AND') return found1 && found2;
        return found1 || found2;
    }

    escapeRegExp(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    getColLetter(n) {
        let letter = '';
        while (n >= 0) {
            letter = String.fromCharCode((n % 26) + 65) + letter;
            n = Math.floor(n / 26) - 1;
        }
        return letter;
    }

    addResultToTable(file, count, type, location) {
        const tbody = document.querySelector('#resultsTable tbody');
        const row = document.createElement('tr');
        row.innerHTML = `
            <td class="highlight-row">${file.name}</td>
            <td>${count}</td>
            <td><span class="version-tag">${type}</span></td>
            <td style="font-size: 0.8rem; color: #94a3b8;">${location}</td>
        `;

        // Add click event to open file
        row.addEventListener('click', () => {
            const url = URL.createObjectURL(file);
            const a = document.createElement('a');
            a.href = url;
            a.download = file.name; // For Excel, this triggers download
            a.target = '_blank';    // For PDF, this might open in new tab

            // For PDFs, we prefer opening to downloading if supported
            if (file.type === 'application/pdf') {
                window.open(url, '_blank');
            } else {
                // For Excel and others, trigger download link
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            }

            // Revoke URL after a delay to free memory
            setTimeout(() => URL.revokeObjectURL(url), 100);
        });

        tbody.appendChild(row);
    }
}

// Instantiate tool
document.addEventListener('DOMContentLoaded', () => {
    window.fileSearchTool = new WebFileSearchTool();
});
