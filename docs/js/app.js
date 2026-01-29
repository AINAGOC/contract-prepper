/**
 * 契約書一括整形・チェックツール - Client-side version
 * All processing happens in the browser - no server required
 */

// Configuration
const CONFIG = {
    fileTypes: {
        contract: { label: '① 契約書', naming: '基本契約書_{company}' },
        estimate: { label: '② 見積書', naming: '別紙１_{company}' },
        oath: { label: '③ 誓約書', naming: '誓約書_{company}' },
        checklist: { label: '④ チェックシート', naming: 'チェックシート_{company}' },
        confirmation: { label: '⑤ 確認書', naming: '確認書_{company}' }
    },
    // 別紙2チェック用キーワード
    appendix2: {
        keywords: ['別紙2', '別紙２'],
        latestKeywords: ['愛知・名古屋2026', '2026アジア・アジアパラ競技大会'],
        oldKeywords: ['第20回アジア競技大会', '2026年アジア競技大会']
    },
    // 署名捺印条項
    sealClause: '本契約の成立を証するため',
    // パートナーページキーワード
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

        // Collect uploaded files
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

        // Start processing
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

    // Process each file type
    for (const [key, file] of Object.entries(files)) {
        setSpinnerText(`${CONFIG.fileTypes[key].label} を処理中...`);

        try {
            // Add to backup
            const backupData = await file.arrayBuffer();
            backupFolder.file(file.name, backupData);

            // Process based on file type
            const ext = file.name.split('.').pop().toLowerCase();

            if (ext === 'pdf') {
                // PDF files - just copy with new name
                const newName = CONFIG.fileTypes[key].naming.replace('{company}', companyName) + '.pdf';
                outputFolder.file(newName, backupData);
                results.processed.push(newName);
                results.warnings.push(`${CONFIG.fileTypes[key].label}: PDF形式のため整形処理はスキップされました。`);
            } else if (ext === 'docx') {
                // Word files
                const processResult = await processDocx(file, key, companyName, approvalType);
                results.errors.push(...processResult.errors);
                results.warnings.push(...processResult.warnings);

                if (processResult.outputData) {
                    const newName = CONFIG.fileTypes[key].naming.replace('{company}', companyName) + '.docx';
                    outputFolder.file(newName, processResult.outputData);
                    results.processed.push(newName);
                }
            } else if (ext === 'xlsx') {
                // Excel files
                const processResult = await processXlsx(file, key, companyName);
                results.errors.push(...processResult.errors);
                results.warnings.push(...processResult.warnings);

                if (processResult.outputData) {
                    const newName = CONFIG.fileTypes[key].naming.replace('{company}', companyName) + '.xlsx';
                    outputFolder.file(newName, processResult.outputData);
                    results.processed.push(newName);
                }
            }
        } catch (err) {
            results.errors.push(`${CONFIG.fileTypes[key].label}: 処理エラー - ${err.message}`);
        }
    }

    // Generate ZIP
    if (results.processed.length > 0) {
        setSpinnerText('ZIPファイルを作成中...');
        const timestamp = new Date().toISOString().replace(/[:\-T]/g, '').slice(0, 14);
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        saveAs(zipBlob, `契約書_${companyName}_${timestamp}.zip`);
    }

    return results;
}

// Process Word document
async function processDocx(file, fileType, companyName, approvalType) {
    const result = { errors: [], warnings: [], outputData: null };

    try {
        const arrayBuffer = await file.arrayBuffer();

        // Extract text using mammoth for validation
        const mammothResult = await mammoth.extractRawText({ arrayBuffer });
        const fullText = mammothResult.value;

        // Validation based on file type
        if (fileType === 'contract') {
            // Contract-specific validation
            if (approvalType === 'paper') {
                // Check date fields
                const datePattern = /\d{1,2}\s*月\s*\d{1,2}\s*日/;
                if (datePattern.test(fullText)) {
                    result.errors.push('【契約書エラー】日付欄に具体的な月日が記入されている可能性があります。確認してください。');
                }

                // Check seal clause exists
                if (!fullText.includes(CONFIG.sealClause)) {
                    result.errors.push('【契約書エラー】署名捺印条項「本契約の成立を証するため〜」が見つかりません。');
                }
            }

            // Check appendix2 version
            checkAppendix2Version(fullText, result);

            // Check for partner pages
            for (const kw of CONFIG.partnerKeywords) {
                if (fullText.includes(kw)) {
                    result.warnings.push('【契約書注意】「カテゴリー及びパートナー」のセクションが含まれています。手動で削除してください。');
                    break;
                }
            }
        }

        if (fileType === 'oath') {
            // Oath document validation
            // Check for old project name
            if (fullText.includes('第20回アジア競技大会')) {
                result.errors.push('【誓約書エラー】旧件名「第20回アジア競技大会」が含まれています。最新の件名に修正してください。');
            }
        }

        // Output the original file (browser cannot easily modify docx)
        result.outputData = arrayBuffer;
        result.warnings.push(`${CONFIG.fileTypes[fileType].label}: 元のファイルをそのまま出力しました（ブラウザでの書式変更は非対応）。`);

    } catch (err) {
        result.errors.push(`Word処理エラー: ${err.message}`);
    }

    return result;
}

// Check appendix2 version
function checkAppendix2Version(text, result) {
    // Find appendix2 section
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

    // Check for old keywords
    for (const oldKw of CONFIG.appendix2.oldKeywords) {
        if (appendix2Text.includes(oldKw)) {
            result.errors.push(`【別紙2エラー】旧様式の可能性があります。「${oldKw}」が検出されました。最新の別紙2に差し替えてください。`);
            return;
        }
    }

    // Check for latest keywords
    const hasLatest = CONFIG.appendix2.latestKeywords.some(kw => appendix2Text.includes(kw));
    if (!hasLatest) {
        result.warnings.push('【別紙2確認】最新様式のキーワードが見つかりませんでした。別紙2が最新版であることを確認してください。');
    }
}

// Process Excel file
async function processXlsx(file, fileType, companyName) {
    const result = { errors: [], warnings: [], outputData: null };

    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Extract text from all sheets for validation
        let fullText = '';
        for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];
            const text = XLSX.utils.sheet_to_txt(sheet);
            fullText += text + '\n';
        }

        // Basic validation
        if (fileType === 'estimate') {
            // Check for company name in estimate
            if (!fullText.includes(companyName) && companyName.length > 2) {
                result.warnings.push('【見積書確認】会社名が見積書内で確認できませんでした。');
            }
        }

        if (fileType === 'checklist') {
            // Checklist validation - check for required items
            if (!fullText.includes('チェック')) {
                result.warnings.push('【チェックシート確認】チェック項目が見つかりませんでした。正しいファイルか確認してください。');
            }
        }

        // Output the original file
        result.outputData = arrayBuffer;

    } catch (err) {
        result.errors.push(`Excel処理エラー: ${err.message}`);
    }

    return result;
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

    alert.className = 'alert';
    alert.innerHTML = html;
}

// Show simple result message
function showResult(type, messages) {
    resultArea.style.display = 'block';
    const alert = document.getElementById('resultAlert');
    alert.className = 'alert ' + (type === 'success' ? 'alert-success' : 'alert-danger');
    alert.innerHTML = messages.map(m => `<p class="mb-1">${m}</p>`).join('');
}

// Update spinner text
function setSpinnerText(text) {
    if (spinnerText) spinnerText.textContent = text;
}
