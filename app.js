// DOM Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const uploadSection = document.getElementById('upload-section');
const loadingSection = document.getElementById('loading-section');
const resultsSection = document.getElementById('results-section');
const resetBtn = document.getElementById('reset-btn');
const dateToleranceInput = document.getElementById('date-tolerance');
const dateToleranceVal = document.getElementById('date-tolerance-val');
const statTotal = document.getElementById('stat-total');
const statGroups = document.getElementById('stat-groups');
const statRecords = document.getElementById('stat-records');
const duplicatesContainer = document.getElementById('duplicates-container');
const emptyState = document.getElementById('empty-state');
const filterBtns = document.querySelectorAll('.filter-btn');
const searchInput = document.getElementById('search-input');

// State
let allDuplicates = [];
let currentFilter = 'all'; // 'all', 'exact-date', 'near-date'

// --- Event Listeners ---

// Update Range Value display
dateToleranceInput.addEventListener('input', (e) => {
    const val = e.target.value;
    dateToleranceVal.textContent = `±${val}日`;
});

// Reset Application
resetBtn.addEventListener('click', () => {
    uploadSection.style.display = 'flex';
    resultsSection.style.display = 'none';
    resetBtn.classList.add('hidden');
    fileInput.value = '';
    allDuplicates = [];
    duplicatesContainer.innerHTML = '';
});

// Filter Buttons
filterBtns.forEach(btn => {
    btn.addEventListener('click', (e) => {
        filterBtns.forEach(b => {
            b.classList.remove('active', 'bg-indigo-100', 'text-indigo-700');
            b.classList.add('text-slate-600', 'hover:bg-slate-100');
        });
        
        const target = e.target;
        target.classList.remove('text-slate-600', 'hover:bg-slate-100');
        target.classList.add('active', 'bg-indigo-100', 'text-indigo-700');
        
        currentFilter = target.dataset.filter;
        renderResults();
    });
});

// Search Input
if(searchInput) {
    searchInput.addEventListener('input', () => {
        renderResults();
    });
}

// Drag & Drop Handling
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, highlight, false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, unhighlight, false);
});

function highlight(e) {
    dropZone.querySelector('.relative').classList.add('border-indigo-500', 'bg-indigo-50/50');
}

function unhighlight(e) {
    dropZone.querySelector('.relative').classList.remove('border-indigo-500', 'bg-indigo-50/50');
}

dropZone.addEventListener('drop', handleDrop, false);
fileInput.addEventListener('change', handleFileSelect, false);

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    if (e.target.files.length > 0) {
        processFile(e.target.files[0]);
    }
}

// --- Core Logic ---

function processFile(file) {
    if (!file.name.toLowerCase().endsWith('.csv')) {
        alert('CSVファイルを選択してください。');
        return;
    }

    // Show loading UI
    uploadSection.style.display = 'none';
    loadingSection.style.display = 'flex';
    resetBtn.classList.add('hidden');

    Papa.parse(file, {
        header: false, // Parse as array of arrays first to find headers
        skipEmptyLines: true,
        encoding: "Shift-JIS", // MJS usually exports in Shift-JIS
        complete: function(results) {
            analyzeData(results.data);
        },
        error: function(error) {
            console.error(error);
            alert('ファイルの読み込みに失敗しました。エンコーディング等を変更して再試行してください。');
            showUpload();
        }
    });
}

function showUpload() {
    uploadSection.style.display = 'flex';
    loadingSection.style.display = 'none';
    resultsSection.style.display = 'none';
}

function analyzeData(rows) {
    if (rows.length === 0) {
        alert('データが空です。');
        showUpload();
        return;
    }

    setTimeout(() => {
        // 1. Find Header Row (MJS CSV might have header at top or slightly down)
        let headerRowIdx = -1;
        const potentialDateCols = ['日付', '伝票日付', '処理日', '発生日', '取引日'];
        const potentialAmountCols = ['金額', '借方金額', '貸方金額', '取引金額', '合計'];
        const potentialDescCols = ['摘要', '内容', '備考', '取引先'];
        const potentialDebitCols = ['借方科目', '科目（借）', '科目']; // Left account
        const potentialCreditCols = ['貸方科目', '科目（貸）', '相手科目']; // Right account

        let colMap = { date: -1, amount: -1, amountL: -1, amountR: -1, desc: -1, debit: -1, credit: -1 };

        // Scan first 10 rows to build column map
        for (let i = 0; i < Math.min(10, rows.length); i++) {
            const row = rows[i];
            let dateFound = -1;
            let amountFound = -1;
            let amountLFound = -1;
            let amountRFound = -1;
            let descFound = -1;
            let debitFound = -1;
            let creditFound = -1;

            row.forEach((cell, idx) => {
                if (!cell) return;
                const text = String(cell).trim();
                
                if (potentialDateCols.includes(text) || text.includes('日付')) dateFound = idx;
                else if (text === '借方金額') amountLFound = idx;
                else if (text === '貸方金額') amountRFound = idx;
                else if (potentialAmountCols.includes(text) || text.includes('金額')) amountFound = idx;
                else if (potentialDescCols.includes(text) || text.includes('摘要')) descFound = idx;
                else if (potentialDebitCols.includes(text) || text.includes('借方科目')) debitFound = idx;
                else if (potentialCreditCols.includes(text) || text.includes('貸方科目')) creditFound = idx;
            });

            // If we found at least Date and one type of Amount, we consider this the header row
            if (dateFound !== -1 && (amountFound !== -1 || amountLFound !== -1 || amountRFound !== -1)) {
                headerRowIdx = i;
                colMap = {
                    date: dateFound,
                    amount: amountFound,
                    amountL: amountLFound,
                    amountR: amountRFound,
                    desc: descFound,
                    debit: debitFound,
                    credit: creditFound
                };
                break;
            }
        }

        // Fallback: If no clear header found, try to guess by data types
        if (headerRowIdx === -1) {
            // Assume Row 0 is data if it looks like a date, else Row 0 is unknown header
            headerRowIdx = 0; 
            // Basic guessing based on standard MJS/Accounting layout: [0] Date, [1] Code, [2] Name, [3] Amount L, etc.
            // But let's ask user or rely on generic format
            console.warn("Could not reliably detect standard headers. Attempting data type guessing.");
            showUpload();
            alert("列名（日付、金額、摘要など）を自動検出できませんでした。一般的な形式のCSVファイルをアップロードしてください。");
            return;
        }

        // Extract Data
        const parsedRecords = [];
        let totalProcessed = 0;

        for (let i = headerRowIdx + 1; i < rows.length; i++) {
            const row = rows[i];
            if (row.length === 0) continue;

            totalProcessed++;

            // Extract values safely
            let rawDate = colMap.date !== -1 ? row[colMap.date] : '';
            let rawDesc = colMap.desc !== -1 ? row[colMap.desc] : '';
            let rawDebit = colMap.debit !== -1 ? row[colMap.debit] : '';
            let rawCredit = colMap.credit !== -1 ? row[colMap.credit] : '';
            
            // Format account representation (e.g., 交際費 / 現金)
            let accountStr = '';
            if (rawDebit || rawCredit) {
                const d = rawDebit || '不明';
                const c = rawCredit || '不明';
                accountStr = `${d} / ${c}`;
            }

            // Amount handling: MJS has Left/Right amounts. We take whichever is non-zero
            let parsedAmount = 0;
            if (colMap.amount !== -1) {
                parsedAmount = parseNumber(row[colMap.amount]);
            } else if (colMap.amountL !== -1 || colMap.amountR !== -1) {
                const amountL = colMap.amountL !== -1 ? parseNumber(row[colMap.amountL]) : 0;
                const amountR = colMap.amountR !== -1 ? parseNumber(row[colMap.amountR]) : 0;
                parsedAmount = amountL > 0 ? amountL : amountR;
            }

            // Skip zero amounts or invalid
            if (!parsedAmount || parsedAmount === 0 || isNaN(parsedAmount)) continue;

            const dateObj = parseDate(rawDate);
            if (!dateObj) continue; // Skip invalid dates

            parsedRecords.push({
                originalIndex: i, // Keep track of row for reference
                dateStr: rawDate,
                dateObj: dateObj, // Date object for math
                amount: parsedAmount,
                desc: rawDesc || '(空欄)',
                account: accountStr || '-' // Add account info
            });
        }

        statTotal.innerHTML = `${totalProcessed}<span class="text-sm font-normal text-slate-500 ml-1">件</span>`;

        findDuplicates(parsedRecords);

    }, 500); // Small UI delay
}

function parseNumber(str) {
    if (!str) return 0;
    // Remove commas, spaces, currency symbols
    const cleanStr = String(str).replace(/[,\s\\¥\\\\]/g, '');
    const num = Number(cleanStr);
    return isNaN(num) ? 0 : num;
}

function parseDate(dateStr) {
    if (!dateStr) return null;
    let d = String(dateStr).trim();
    // Replace slashes with hyphens, handle MJS dates like 2023/04/01 or R5/4/1
    
    // Simplistic handling of YYYY/MM/DD or YYYY-MM-DD
    const match = d.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (match) {
        return new Date(match[1], parseInt(match[2])-1, match[3]);
    }
    
    // Fallback Date parsing
    const dateObj = new Date(d);
    if (!isNaN(dateObj.getTime())) {
        return dateObj;
    }
    return null;
}

function findDuplicates(records) {
    // 1. Group strictly by Amount
    const amountGroups = {};
    records.forEach(r => {
        if (!amountGroups[r.amount]) {
            amountGroups[r.amount] = [];
        }
        amountGroups[r.amount].push(r);
    });

    const toleranceDays = parseInt(dateToleranceInput.value, 10);
    const msPerDay = 1000 * 60 * 60 * 24;

    allDuplicates = [];

    // 2. Identify duplicates within amount groups
    for (const amount in amountGroups) {
        const items = amountGroups[amount];
        if (items.length < 2) continue; // Single item, no duplicate possible

        // Sort by date inside group
        items.sort((a, b) => a.dateObj - b.dateObj);

        let visited = new Set();

        for (let i = 0; i < items.length; i++) {
            if (visited.has(i)) continue;
            
            const currentItem = items[i];
            const currentGroup = [currentItem];
            visited.add(i);

            // Look ahead for near dates
            for (let j = i + 1; j < items.length; j++) {
                if (visited.has(j)) continue;

                const compareItem = items[j];
                const dayDiff = Math.abs((compareItem.dateObj - currentItem.dateObj) / msPerDay);

                if (dayDiff <= toleranceDays) {
                    currentGroup.push(compareItem);
                    visited.add(j);
                } else {
                    // Since sorted by date, if it exceeds tolerance, break early
                    // BUT wait, a chain A -> B -> C could extend the tolerance. 
                    // To keep it simple, we compare strictly against the base item of the sub-group.
                }
            }

            if (currentGroup.length > 1) {
                // Determine if exact match or near match
                let isExactDate = true;
                const firstDate = currentGroup[0].dateObj.getTime();
                
                for (let k = 1; k < currentGroup.length; k++) {
                    if (currentGroup[k].dateObj.getTime() !== firstDate) {
                        isExactDate = false;
                        break;
                    }
                }

                allDuplicates.push({
                    amount: currentItem.amount,
                    items: currentGroup,
                    isExactDate: isExactDate,
                    id: `group-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`
                });
            }
        }
    }

    // Sort duplicates by Amount descending
    allDuplicates.sort((a, b) => b.amount - a.amount);

    // Update stats
    statGroups.innerHTML = `${allDuplicates.length}<span class="text-sm font-normal text-slate-500 ml-1">件</span>`;
    
    let totalDuplicatedRecords = 0;
    allDuplicates.forEach(g => totalDuplicatedRecords += g.items.length);
    statRecords.innerHTML = `${totalDuplicatedRecords}<span class="text-sm font-normal text-slate-500 ml-1">件</span>`;

    // Initialize UI
    document.querySelector('.filter-btn[data-filter="all"]').click(); // Triggers renderResults
    
    // Switch Views
    loadingSection.style.display = 'none';
    resultsSection.style.display = 'flex';
    resetBtn.classList.remove('hidden');
}

function renderResults() {
    duplicatesContainer.innerHTML = '';
    
    const searchTerm = searchInput ? searchInput.value.toLowerCase() : '';

    const filtered = allDuplicates.filter(group => {
        // 1. Check Category Filter
        let categoryMatch = false;
        if (currentFilter === 'all') categoryMatch = true;
        else if (currentFilter === 'exact-date' && group.isExactDate) categoryMatch = true;
        else if (currentFilter === 'near-date' && !group.isExactDate) categoryMatch = true;

        if (!categoryMatch) return false;

        // 2. Check Search Search term
        if (!searchTerm) return true;

        const amountStr = String(group.amount);
        if (amountStr.includes(searchTerm)) return true;

        // Check inside items
        return group.items.some(item => 
            item.desc.toLowerCase().includes(searchTerm) || 
            item.dateStr.includes(searchTerm)
        );
    });

    if (filtered.length === 0) {
        if(allDuplicates.length === 0) {
             emptyState.style.display = 'block';
             document.querySelector('.grid').style.display = 'grid'; // Keep stats visible
             document.querySelector('.filter-btn').parentElement.parentElement.style.display = 'none'; // hide filters
        } else {
             emptyState.style.display = 'none';
             duplicatesContainer.innerHTML = `<p class="text-center text-slate-500 py-8">検索条件に一致する結果がありません。</p>`;
        }
        return;
    }

    emptyState.style.display = 'none';
    document.querySelector('.filter-btn').parentElement.parentElement.style.display = 'flex'; // show filters

    // Add wrapper around cards
    let html = '';

    filtered.forEach((group, index) => {
        // Format Currency
        const formatter = new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' });
        const amountFormatted = formatter.format(group.amount);

        // Tags
        const tagsHtml = group.isExactDate 
            ? `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-800"><i class="fa-solid fa-calendar-day"></i> 日付・金額 完全一致</span>`
            : `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-800"><i class="fa-solid fa-calendar-days"></i> 日付 近接 (${group.items.length}件)</span>`;

        let rowsHtml = '';
        group.items.forEach(item => {
            // Check similarities in desc (Basic naive highlight for demonstration)
            let descHtml = escapeHtml(item.desc);
            let accountHtml = escapeHtml(item.account);

            rowsHtml += `
                <tr class="border-b border-slate-100 last:border-0 hover:bg-slate-50 transition-colors">
                    <td class="px-4 py-3 whitespace-nowrap text-sm text-slate-600 font-medium w-32">
                        ${escapeHtml(item.dateStr)}
                    </td>
                    <td class="px-4 py-3 text-sm text-slate-500 w-24">
                        <span class="bg-slate-100 px-1.5 py-0.5 rounded text-xs font-mono border border-slate-200" title="CSV行番号">行 ${item.originalIndex + 1}</span>
                    </td>
                    <td class="px-4 py-3 text-sm text-slate-600 w-48 truncate" title="${accountHtml}">
                        <div class="flex items-center gap-1.5 opacity-90">
                            <i class="fa-solid fa-tags text-slate-400 text-xs"></i> ${accountHtml}
                        </div>
                    </td>
                    <td class="px-4 py-3 text-sm text-slate-800 w-full truncate" title="${descHtml}">
                        ${descHtml}
                    </td>
                </tr>
            `;
        });

        html += `
            <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden group-card animate-fade-in" style="animation-delay: ${index * 0.05}s">
                <!-- Group Header -->
                <div class="bg-slate-50 px-6 py-4 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
                    <div class="flex items-center gap-4">
                        <div class="text-2xl font-bold text-slate-800 tracking-tight">${amountFormatted}</div>
                        ${tagsHtml}
                    </div>
                </div>
                
                <!-- Items Table -->
                <div class="overflow-x-auto">
                    <table class="min-w-full divide-y divide-slate-200">
                        <tbody class="divide-y divide-slate-100 bg-white">
                            ${rowsHtml}
                        </tbody>
                    </table>
                </div>
            </div>
        `;
    });

    duplicatesContainer.innerHTML = html;
}

// Utility
function escapeHtml(unsafe) {
    if (!unsafe) return '';
    return unsafe
         .replace(/&/g, "&amp;")
         .replace(/</g, "&lt;")
         .replace(/>/g, "&gt;")
         .replace(/"/g, "&quot;")
         .replace(/'/g, "&#039;");
}
