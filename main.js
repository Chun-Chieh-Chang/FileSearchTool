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
        this.resultFiles = [];
        this.blobUrls = [];
        this.initEventListeners();
    }

    initEventListeners() {
        const folderInput = document.getElementById('folderInput');
        const fileInput = document.getElementById('fileInput');
        const startBtn = document.getElementById('startSearch');
        const stopBtn = document.getElementById('stopSearch');
        const saveBtn = document.getElementById('saveFiles');
        const clearMemBtn = document.getElementById('clearMemory');

        folderInput.addEventListener('change', (e) => this.handleFileSelection(e));
        fileInput.addEventListener('change', (e) => this.handleFileSelection(e));
        startBtn.addEventListener('click', () => this.startSearch());
        stopBtn.addEventListener('click', () => this.stopSearch());
        saveBtn.addEventListener('click', () => this.saveFilesToFolder());
        clearMemBtn.addEventListener('click', () => this.clearMemory());
    }

    handleFileSelection(event) {
        const files = Array.from(event.target.files);
        const unsupportedFiles = [];
        // Filter by extensions
        this.selectedFiles = files.filter(file => {
            const ext = file.name.split('.').pop().toLowerCase();
            const isSupported = ['xlsx', 'xls', 'pdf'].includes(ext);
            if (!isSupported) {
                unsupportedFiles.push(file.name);
            }
            return isSupported;
        });

        if (unsupportedFiles.length > 0) {
            const maxFilesToShow = 3;
            const fileList = unsupportedFiles.slice(0, maxFilesToShow).join(', ') + 
                (unsupportedFiles.length > maxFilesToShow ? ` 等 ${unsupportedFiles.length} 個檔案` : '');
            alert(`以下檔案類型不支援，已自動忽略：\n${fileList}\n\n本工具僅支援：Excel (.xlsx, .xls) 與 PDF (.pdf)`);
        }

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

        this.stopSearchFlag = false;
        this.results = [];
        this.resultFiles = [];
        this.errorFiles = [];
        const tableBody = document.querySelector('#resultsTable tbody');
        tableBody.innerHTML = '';

        const progressSection = document.getElementById('progressSection');
        const progressBar = document.getElementById('progressBar');
        const progressLabel = document.getElementById('progressLabel');
        const statusLabel = document.getElementById('statusLabel');
        const startBtn = document.getElementById('startSearch');
        const stopBtn = document.getElementById('stopSearch');
        const saveBtn = document.getElementById('saveFiles');

        progressSection.style.display = 'block';
        startBtn.disabled = true;
        stopBtn.disabled = false;
        saveBtn.disabled = true;

        let processedCount = 0;
        const totalFiles = this.selectedFiles.length;
        let successCount = 0;
        let errorCount = 0;

        const CONCURRENCY = 4;
        const filesToProcess = this.selectedFiles.filter(file => {
            const ext = file.name.split('.').pop().toLowerCase();
            if (typeFilter === 'Excel') return ['xlsx', 'xls'].includes(ext);
            if (typeFilter === 'PDF') return ext === 'pdf';
            return ['xlsx', 'xls', 'pdf'].includes(ext);
        });

        const validTotalFiles = filesToProcess.length;

        for (let i = 0; i < validTotalFiles; i += CONCURRENCY) {
            if (this.stopSearchFlag) break;

            const batch = filesToProcess.slice(i, i + CONCURRENCY);
            const batchPromises = batch.map(file => 
                Promise.race([
                    this.searchInFile(file, kw1, kw2, logic, wholeWord, caseSensitive),
                    new Promise((_, reject) => 
                        setTimeout(() => reject(new Error('處理逾時')), 30000)
                    )
                ]).then(matchData => ({ success: true, data: matchData, file }))
                 .catch(err => ({ success: false, error: err, file }))
            );

            const results = await Promise.allSettled(batchPromises);

            for (const result of results) {
                if (result.status === 'fulfilled') {
                    const { success, data, error, file } = result.value;
                    
                    if (success) {
                        statusLabel.innerText = `正在處理: ${file.name}`;
                        if (data.isMatch) {
                            this.addResultToTable(file, data.totalCount, data.type, data.location);
                        }
                        successCount++;
                    } else {
                        errorCount++;
                        this.errorFiles.push({ name: file.name, error: error.message });
                        console.error(`Error processing ${file.name}:`, error);
                        statusLabel.innerText = `錯誤: ${file.name}`;
                    }
                }
                processedCount++;
                const percent = Math.round((processedCount / validTotalFiles) * 100);
                progressBar.style.width = `${percent}%`;
                progressLabel.innerText = `進度: ${percent}% (${processedCount}/${validTotalFiles})`;
            }
        }

        processedCount = this.selectedFiles.length - validTotalFiles + processedCount;
        const finalPercent = Math.round((processedCount / totalFiles) * 100);
        progressBar.style.width = `${finalPercent}%`;
        progressLabel.innerText = `進度: ${finalPercent}% (${processedCount}/${totalFiles})`;

        let finalMessage = this.stopSearchFlag ? '搜尋已停止' : '搜尋完成';
        if (errorCount > 0) {
            finalMessage += ` (成功: ${successCount}, 錯誤: ${errorCount})`;
            if (errorCount <= 5) {
                const errorList = this.errorFiles.map(e => `${e.name}: ${e.error}`).join('\n');
                setTimeout(() => alert(`以下檔案處理失敗：\n\n${errorList}`), 500);
            } else {
                setTimeout(() => alert(`${errorCount} 個檔案處理失敗，請查看主控台了解詳情`), 500);
            }
        }
        statusLabel.innerText = finalMessage;
        startBtn.disabled = false;
        stopBtn.disabled = true;
        saveBtn.disabled = this.resultFiles.length === 0;
    }

    stopSearch() {
        this.stopSearchFlag = true;
    }

    async saveFilesToFolder() {
        if (this.resultFiles.length === 0) {
            alert('沒有可另存的檔案');
            return;
        }

        if (!window.showDirectoryPicker) {
            this.fallbackDownload();
            return;
        }

        try {
            const dirHandle = await window.showDirectoryPicker({
                mode: 'readwrite',
                startIn: 'downloads'
            });

            let savedCount = 0;
            for (const file of this.resultFiles) {
                try {
                    const fileHandle = await dirHandle.getFileHandle(file.name, { create: true });
                    const writable = await fileHandle.createWritable();
                    await writable.write(file);
                    await writable.close();
                    savedCount++;
                } catch (err) {
                    console.error(`無法儲存 ${file.name}:`, err);
                }
            }

            alert(`成功將 ${savedCount} / ${this.resultFiles.length} 個檔案另存到選擇的資料夾`);
        } catch (err) {
            if (err.name === 'AbortError') {
                return;
            }
            console.error('儲存檔案時發生錯誤:', err);
            this.fallbackDownload();
        }
    }

    fallbackDownload() {
        if (confirm('您的瀏覽器不支援資料夾選擇功能，是否改用逐個下載方式？')) {
            let delay = 0;
            this.resultFiles.forEach((file, index) => {
                setTimeout(() => {
                    const url = URL.createObjectURL(file);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = file.name;
                    a.style.display = 'none';
                    document.body.appendChild(a);
                    a.click();
                    setTimeout(() => {
                        document.body.removeChild(a);
                        URL.revokeObjectURL(url);
                    }, 500);
                }, delay);
                delay += 300;
            });
            alert(`將開始下載 ${this.resultFiles.length} 個檔案，請允許瀏覽器的多檔案下載`);
        }
    }

    async searchInFile(file, kw1, kw2, logic, wholeWord, caseSensitive) {
        const ext = file.name.split('.').pop().toLowerCase();
        let matchData = { isMatch: false, totalCount: 0, type: '', location: '' };

        try {
            if (['xlsx', 'xls'].includes(ext)) {
                matchData.type = 'Excel';

                if (file.size > 50 * 1024 * 1024) {
                    throw new Error('檔案過大 (超過 50MB)，無法處理');
                }

                let data;
                try {
                    data = await file.arrayBuffer();
                } catch (e) {
                    throw new Error('無法讀取檔案內容');
                }

                let workbook;
                try {
                    workbook = XLSX.read(data, { type: 'array' });
                } catch (e) {
                    throw new Error('檔案格式錯誤或已損壞');
                }

                let kw1Found = false;
                let kw2Found = false;
                let kw1Count = 0;
                let kw2Count = 0;
                let firstLoc = '';

                try {
                    for (const sheetName of workbook.SheetNames) {
                        const sheet = workbook.Sheets[sheetName];
                        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

                        let shouldExitSheet = false;
                        for (let rIdx = 0; rIdx < rows.length; rIdx++) {
                            if (shouldExitSheet) break;

                            const row = rows[rIdx];
                            for (let cIdx = 0; cIdx < row.length; cIdx++) {
                                const cell = row[cIdx];
                                if (cell === null || cell === undefined) continue;

                                try {
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

                                    if (this.checkLogic(kw1Found, kw2Found, kw2, logic) && kw1Found && (!kw2 || kw2Found)) {
                                        shouldExitSheet = true;
                                        break;
                                    }
                                } catch (cellErr) {
                                    console.warn(`Error in cell ${rIdx},${cIdx}:`, cellErr);
                                }
                            }
                        }
                    }
                } catch (sheetErr) {
                    throw new Error(`工作表處理錯誤: ${sheetErr.message}`);
                }

                matchData.isMatch = this.checkLogic(kw1Found, kw2Found, kw2, logic);
                matchData.totalCount = kw1Count + kw2Count;
                matchData.location = firstLoc || '未知位置';

            } else if (ext === 'pdf') {
                matchData.type = 'PDF';

                if (file.size > 100 * 1024 * 1024) {
                    throw new Error('檔案過大 (超過 100MB)，無法處理');
                }

                let data;
                try {
                    data = await file.arrayBuffer();
                } catch (e) {
                    throw new Error('無法讀取檔案內容');
                }

                let pdf;
                try {
                    const loadingTask = pdfjsLib.getDocument({ data });
                    pdf = await loadingTask.promise;
                } catch (e) {
                    throw new Error('PDF 檔案格式錯誤或已損壞');
                }

                let kw1Found = false;
                let kw2Found = false;
                let kw1Count = 0;
                let kw2Count = 0;
                let firstLoc = '';

                try {
                    const maxPages = Math.min(pdf.numPages, 500);
                    for (let i = 1; i <= maxPages; i++) {
                        if (this.checkLogic(kw1Found, kw2Found, kw2, logic) && kw1Found && (!kw2 || kw2Found)) {
                            break;
                        }

                        let page;
                        try {
                            page = await pdf.getPage(i);
                        } catch (pageErr) {
                            console.warn(`Error loading page ${i}:`, pageErr);
                            continue;
                        }

                        try {
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
                        } catch (textErr) {
                            console.warn(`Error reading text from page ${i}:`, textErr);
                        }

                        page.cleanup();
                    }

                    if (pdf.numPages > 500) {
                        console.warn(`PDF has ${pdf.numPages} pages, only first 500 were searched`);
                    }
                } catch (pageErr) {
                    throw new Error(`頁面處理錯誤: ${pageErr.message}`);
                } finally {
                    try {
                        pdf.destroy();
                    } catch (e) {
                        console.warn('Error destroying PDF:', e);
                    }
                }

                matchData.isMatch = this.checkLogic(kw1Found, kw2Found, kw2, logic);
                matchData.totalCount = kw1Count + kw2Count;
                matchData.location = firstLoc || '未知位置';
            }
        } catch (err) {
            throw new Error(`${file.name}: ${err.message}`);
        }

        return matchData;
    }

    countMatches(text, keyword, wholeWord, caseSensitive) {
        if (!keyword) return 0;
        const cacheKey = `${keyword}-${wholeWord}-${caseSensitive}`;

        if (!this.regexCache) {
            this.regexCache = new Map();
        }

        let regex = this.regexCache.get(cacheKey);
        if (!regex) {
            let flags = 'g';
            if (!caseSensitive) flags += 'i';

            let pattern = this.escapeRegExp(keyword);
            if (wholeWord) {
                pattern = `\\b${pattern}\\b`;
            }

            regex = new RegExp(pattern, flags);
            this.regexCache.set(cacheKey, regex);
        }

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

        // Add click event to open/download file
        row.addEventListener('click', () => {
            try {
                const url = URL.createObjectURL(file);
                const a = document.createElement('a');
                a.href = url;
                a.download = file.name;
                a.style.display = 'none';

                document.body.appendChild(a);
                a.click();

                // Keep the URL alive for 1 second instead of 100ms
                setTimeout(() => {
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                }, 1000);
            } catch (err) {
                console.error("無法開啟檔案:", err);
                alert("無法開啟該檔案，請檢查瀏覽器設定。");
            }
        });

        tbody.appendChild(row);
        this.resultFiles.push(file);
    }

    clearMemory() {
        const beforeCount = {
            selectedFiles: this.selectedFiles.length,
            resultFiles: this.resultFiles.length,
            results: this.results.length,
            blobUrls: this.blobUrls.length,
            regexCache: this.regexCache ? this.regexCache.size : 0
        };

        this.selectedFiles = [];
        this.resultFiles = [];
        this.results = [];
        this.errorFiles = [];

        if (this.regexCache) {
            this.regexCache.clear();
            this.regexCache = null;
        }

        this.blobUrls.forEach(url => {
            try {
                URL.revokeObjectURL(url);
            } catch (e) {
                console.warn('Error revoking URL:', e);
            }
        });
        this.blobUrls = [];

        document.getElementById('folderInput').value = '';
        document.getElementById('fileInput').value = '';
        document.getElementById('keyword1').value = '';
        document.getElementById('keyword2').value = '';
        document.getElementById('fileStats').innerText = '未選擇任何檔案';

        const tableBody = document.querySelector('#resultsTable tbody');
        tableBody.innerHTML = '';

        document.getElementById('progressBar').style.width = '0%';
        document.getElementById('progressLabel').innerText = '進度: 0%';
        document.getElementById('statusLabel').innerText = '記憶體已清空';
        document.getElementById('saveFiles').disabled = true;

        const afterCount = {
            selectedFiles: this.selectedFiles.length,
            resultFiles: this.resultFiles.length,
            results: this.results.length,
            blobUrls: this.blobUrls.length,
            regexCache: 0
        };

        console.log('Memory cleared:', { before: beforeCount, after: afterCount });

        setTimeout(() => {
            if (typeof gc === 'function') {
                gc();
                console.log('Garbage collection triggered');
            }
            if (window.gc) {
                window.gc();
            }
        }, 100);

        alert('已清空記憶體：\n\n' +
            `已釋放 ${beforeCount.selectedFiles} 個選擇的檔案\n` +
            `已釋放 ${beforeCount.resultFiles} 個結果檔案\n` +
            `已釋放 ${beforeCount.blobUrls} 個 Blob URLs\n` +
            `已清空搜尋結果`);
    }
}

// Instantiate tool
document.addEventListener('DOMContentLoaded', () => {
    window.fileSearchTool = new WebFileSearchTool();
});
