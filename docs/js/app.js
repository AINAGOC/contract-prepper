/**
 * 契約書一括整形・チェックツール - Client-side version
 * All processing happens in the browser - no server required
 */

// Configuration
const CONFIG = {
    fileTypes: {
        contract: { label: '① 契約書', naming: '基本契約書_{company}', type: 'docx' },
        estimate: { label: '② 見積書', naming: '別紙１_{company}', type: 'xlsx' },
        oath: { label: '③ 誓約書', naming: '誓約書_{company}', type: 'docx' },
        checklist: { label: '④ チェックシート', naming: 'チェックシート_{company}', type: 'xlsx' },
        confirmation: { label: '⑤ 確認書', naming: '確認書_{company}', type: 'docx' }
    },
    appendix2: {
        keywords: ['別紙2', '別紙２'],
        latestKeywords: ['愛知・名古屋2026', '2026アジア・アジアパラ競技大会'],
        oldKeywords: ['第20回アジア競技大会', '2026年アジア競技大会']
    },
    sealClause: '本契約の成立を証するため',
    partnerKeywords: ['カテゴリー及びパートナー', 'カテゴリー・パートナー']
};

// DOM Elements
let form, spinner, resultArea, submitBtn, spinnerText;

document.addEventListener('DOMContentLoaded', () => {
    form = document.getElementById('mainForm');
    spinner = document.getElementById('spinner');
    resultArea = document.getElementById('resultArea');
    submitBtn = document.getElementById('submitBtn');
    spinnerText = document.getElementById('spinnerText');

    setupDropZones();
    setupFormSubmit();
});

// Setup drag & drop zones
function setupDropZones() {
    document.querySelectorAll('.drop-zone').forEach(zone => {
        const input = zone.querySelector('.drop-zone__input');
        const key = zone.dataset.input;

        zone.addEventListener('click', () => input.click());

        zone.addEventListener('dragover', e => {
            e.preventDefault();
            zone.classList.add('drop-zone--over');
        });

        zone.addEventListener('dragleave', () => {
            zone.classList.remove('drop-zone--over');
        });

        zone.addEventListener('drop', e => {
            e.preventDefault();
            zone.classList.remove('drop-zone--over');
            if (e.dataTransfer.files.length) {
                input.files = e.dataTransfer.files;
                updateFileName(key, input.files[0].name);
                zone.classList.add('drop-zone--has-file');
            }
        });

        input.addEventListener('change', () => {
            if (input.files.length) {
                updateFileName(key, input.files[0].name);
                zone.classList.add('drop-zone--has-file');
            }
        });
    });
}

function updateFileName(key, name) {
    const el = document.getElementById('fname_' + key);
    if (el) el.textContent = name;
}

// Setup form submission
function setupFormSubmit() {
    form.addEventListener('submit', async e => {
        e.preventDefault();

        const companyName = document.getElementById('company_name').value.trim();
        if (!companyName) {
            showResult('error', ['会社名を入力してください。']);
            return;
        }

        const approvalType = document.querySelector('input[name="approval_type"]:checked').value;

        const files = {};
        let hasFile = false;
        for (const key of Object.keys(CONFIG.fileTypes)) {
            const input = document.querySelector(`input[name="${key}"]`);
            if (input && input.files.length > 0) {
                files[key] = input.files[0];
                hasFile = true;
            }
        }

        if (!hasFile) {
            showResult('error', ['少なくとも1つのファイルをアップロードしてください。']);
            return;
        }

        spinner.style.display = 'block';
        resultArea.style.display = 'none';
        submitBtn.disabled = true;

        try {
            const result = await processFiles(files, companyName, approvalType);
            displayResults(result);
        } catch (err) {
            console.error('Processing error:', err);
            showResult('error', ['処理中にエラーが発生しました: ' + err.message]);
        } finally {
            spinner.style.display = 'none';
            submitBtn.disabled = false;
        }
    });
}

// Main processing function
async function processFiles(files, companyName, approvalType) {
    const results = {
        processed: [],
        errors: [],
        warnings: [],
        outputFiles: []
    };

    const zip = new JSZip();
    const outputFolder = zip.folder('成果物');
    const backupFolder = zip.folder('バックアップ');

    for (const [key, file] of Object.entries(files)) {
        setSpinnerText(`${CONFIG.fileTypes[key].label} を処理中...`);

        try {
            const backupData = await file.arrayBuffer();
            backupFolder.file(file.name, backupData);

            const ext = file.name.split('.').pop().toLowerCase();

            if (ext === 'pdf') {
                // PDF - copy with new name
                const newName = CONFIG.fileTypes[key].naming.replace('{company}', companyName) + '.pdf';
                outputFolder.file(newName, backupData);
                results.processed.push(newName);
            } else if (ext === 'docx') {
                // Word document processing
                const processResult = await processDocx(file, key, companyName, approvalType);
                results.errors.push(...processResult.errors);
                results.warnings.push(...processResult.warnings);

                // Output original DOCX (layout preserved)
                const docxName = CONFIG.fileTypes[key].naming.replace('{company}', companyName) + '.docx';
                outputFolder.file(docxName, backupData);
                results.processed.push(docxName);
            } else if (ext === 'xlsx') {
                // Excel document processing
                const processResult = await processXlsx(file, key, companyName);
                results.errors.push(...processResult.errors);
                results.warnings.push(...processResult.warnings);

                // Output original XLSX (layout preserved)
                const xlsxName = CONFIG.fileTypes[key].naming.replace('{company}', companyName) + '.xlsx';
                outputFolder.file(xlsxName, backupData);
                results.processed.push(xlsxName);
            }
        } catch (err) {
            console.error(`Error processing ${key}:`, err);
            results.errors.push(`${CONFIG.fileTypes[key].label}: 処理エラー - ${err.message}`);
        }
    }

    if (results.processed.length > 0) {
        setSpinnerText('ZIPファイルを作成中...');
        const timestamp = new Date().toISOString().replace(/[:\-T]/g, '').slice(0, 14);
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        saveAs(zipBlob, `契約書_${companyName}_${timestamp}.zip`);
    }

    return results;
}

// Process Word document - validate only
async function processDocx(file, fileType, companyName, approvalType) {
    const result = { errors: [], warnings: [] };

    try {
        const arrayBuffer = await file.arrayBuffer();

        // Extract text for validation
        setSpinnerText(`${CONFIG.fileTypes[fileType].label} をチェック中...`);
        const fullText = await mammoth.extractRawText({ arrayBuffer }).then(r => r.value);

        // Validation
        if (fileType === 'contract') {
            validateContract(fullText, approvalType, result);
        }
        if (fileType === 'oath') {
            validateOath(fullText, result);
        }

    } catch (err) {
        console.error('DOCX processing error:', err);
        result.errors.push(`Word処理エラー: ${err.message}`);
    }

    return result;
}

// Process Excel file - validate only
async function processXlsx(file, fileType, companyName) {
    const result = { errors: [], warnings: [] };

    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Extract text for validation
        setSpinnerText(`${CONFIG.fileTypes[fileType].label} をチェック中...`);
        let fullText = '';
        for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];
            const text = XLSX.utils.sheet_to_txt(sheet);
            fullText += text + '\n';
        }

        // Validation
        if (fileType === 'estimate') {
            if (!fullText.includes(companyName) && companyName.length > 2) {
                result.warnings.push('【見積書確認】会社名が見積書内で確認できませんでした。');
            }
        }

        if (fileType === 'checklist') {
            if (!fullText.includes('チェック')) {
                result.warnings.push('【チェックシート確認】チェック項目の形式を確認してください。');
            }
        }

    } catch (err) {
        console.error('XLSX processing error:', err);
        result.errors.push(`Excel処理エラー: ${err.message}`);
    }

    return result;
}

// Validate contract document
function validateContract(fullText, approvalType, result) {
    if (approvalType === 'paper') {
        if (!fullText.includes(CONFIG.sealClause)) {
            result.errors.push('【契約書エラー】署名捺印条項が見つかりません。');
        }
    }

    // Check appendix2 version
    checkAppendix2Version(fullText, result);

    // Check for partner pages
    for (const kw of CONFIG.partnerKeywords) {
        if (fullText.includes(kw)) {
            result.warnings.push('【契約書注意】「カテゴリー及びパートナー」セクションが含まれています。別紙2・3以外の箇所は手動で削除してください。');
            break;
        }
    }
}

// Validate oath document
function validateOath(fullText, result) {
    if (fullText.includes('第20回アジア競技大会')) {
        result.errors.push('【誓約書エラー】旧件名「第20回アジア競技大会」が含まれています。');
    }
}

// Check appendix2 version
function checkAppendix2Version(text, result) {
    let inAppendix2 = false;
    let appendix2Text = '';

    const lines = text.split('\n');
    for (const line of lines) {
        if (CONFIG.appendix2.keywords.some(kw => line.includes(kw))) {
            inAppendix2 = true;
            continue;
        }
        if (line.includes('別紙3') || line.includes('別紙３')) {
            break;
        }
        if (inAppendix2) {
            appendix2Text += line + '\n';
        }
    }

    if (!appendix2Text) return;

    for (const oldKw of CONFIG.appendix2.oldKeywords) {
        if (appendix2Text.includes(oldKw)) {
            result.errors.push(`【別紙2エラー】旧様式「${oldKw}」が検出されました。最新版に差し替えてください。`);
            return;
        }
    }

    const hasLatest = CONFIG.appendix2.latestKeywords.some(kw => appendix2Text.includes(kw));
    if (!hasLatest) {
        result.warnings.push('【別紙2確認】最新様式のキーワードが見つかりません。別紙2が最新版か確認してください。');
    }
}

// Display results
function displayResults(results) {
    resultArea.style.display = 'block';
    const alert = document.getElementById('resultAlert');

    let html = '';

    if (results.processed.length > 0) {
        html += '<div class="validation-success"><strong>処理完了:</strong><ul class="mb-0">';
        for (const f of results.processed) {
            html += `<li>${f}</li>`;
        }
        html += '</ul></div>';
    }

    if (results.errors.length > 0) {
        html += '<div class="validation-error mt-2"><strong>エラー:</strong><ul class="mb-0">';
        for (const e of results.errors) {
            html += `<li>${e}</li>`;
        }
        html += '</ul></div>';
    }

    if (results.warnings.length > 0) {
        html += '<div class="validation-warning mt-2"><strong>警告・確認事項:</strong><ul class="mb-0">';
        for (const w of results.warnings) {
            html += `<li>${w}</li>`;
        }
        html += '</ul></div>';
    }

    // PDF conversion instructions
    html += `<div class="alert alert-info mt-3">
        <strong>PDF変換方法:</strong><br>
        <small>出力されたファイルをWord/Excelで開き、「ファイル」→「エクスポート」→「PDF/XPSの作成」でPDFに変換してください。<br>
        ※レイアウトを維持するため、元のファイル形式で出力しています。</small>
    </div>`;

    alert.className = '';
    alert.innerHTML = html;
}

function showResult(type, messages) {
    resultArea.style.display = 'block';
    const alert = document.getElementById('resultAlert');
    alert.className = 'alert ' + (type === 'success' ? 'alert-success' : 'alert-danger');
    alert.innerHTML = messages.map(m => `<p class="mb-1">${m}</p>`).join('');
}

function setSpinnerText(text) {
    if (spinnerText) spinnerText.textContent = text;
}
