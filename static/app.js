document.addEventListener('DOMContentLoaded', () => {
    // Drag & drop zones
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

    function updateFileName(key, name) {
        const el = document.getElementById('fname_' + key);
        if (el) el.textContent = name;
    }

    // Form submission
    const form = document.getElementById('mainForm');
    const spinner = document.getElementById('spinner');
    const resultArea = document.getElementById('resultArea');
    const submitBtn = document.getElementById('submitBtn');
    const spinnerText = document.getElementById('spinnerText');

    // Wake up service before processing
    async function wakeUpService() {
        try {
            console.log('Attempting wake-up call to /health...');
            const resp = await fetch('/health', {
                method: 'GET',
                cache: 'no-store'
            });
            console.log('Wake-up response:', resp.status);
            return resp.ok;
        } catch (e) {
            console.error('Wake-up failed:', e.name, e.message);
            return false;
        }
    }

    // Retry fetch with exponential backoff
    async function fetchWithRetry(url, options, maxRetries = 3) {
        let lastError;
        for (let i = 0; i < maxRetries; i++) {
            try {
                console.log(`Fetch attempt ${i + 1} to ${url}...`);
                const resp = await fetch(url, options);
                console.log(`Fetch response: ${resp.status}`);
                return resp;
            } catch (e) {
                lastError = e;
                console.error(`Attempt ${i + 1} failed:`, e.name, e.message);
                if (i < maxRetries - 1) {
                    // Wait before retry (2s, 4s, 8s)
                    const waitTime = 2000 * Math.pow(2, i);
                    console.log(`Waiting ${waitTime}ms before retry...`);
                    await new Promise(r => setTimeout(r, waitTime));
                }
            }
        }
        throw lastError;
    }

    form.addEventListener('submit', async e => {
        e.preventDefault();
        const formData = new FormData(form);

        // Basic client-side validation
        if (!formData.get('company_name').trim()) {
            showResult('error', ['会社名を入力してください。']);
            return;
        }

        spinner.style.display = 'block';
        resultArea.style.display = 'none';
        submitBtn.disabled = true;

        try {
            // Step 1: Wake up the service (Render free tier sleeps after 15 min)
            if (spinnerText) spinnerText.textContent = 'サーバーに接続中...';
            console.log('Starting process...');

            let awake = await wakeUpService();
            let attempts = 0;
            while (!awake && attempts < 6) {
                attempts++;
                if (spinnerText) spinnerText.textContent = `サーバー起動を待機中... (${attempts}/6)`;
                console.log(`Wake-up retry ${attempts}/6, waiting 10s...`);
                await new Promise(r => setTimeout(r, 10000)); // Wait 10s
                awake = await wakeUpService();
            }

            if (!awake) {
                throw new Error('サーバーが応答しません。ページを更新してから再試行してください。');
            }

            // Step 2: Process files
            if (spinnerText) spinnerText.textContent = 'ファイルを処理中...（PDF変換に時間がかかる場合があります）';
            console.log('Server awake, sending files for processing...');

            const resp = await fetchWithRetry('/process', { method: 'POST', body: formData });

            if (resp.ok) {
                if (spinnerText) spinnerText.textContent = 'ダウンロード準備中...';
                const blob = await resp.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = resp.headers.get('content-disposition')
                    ?.match(/filename\*?=(?:UTF-8'')?(.+)/)?.[1]
                    || '契約書処理結果.zip';
                document.body.appendChild(a);
                a.click();
                a.remove();
                URL.revokeObjectURL(url);
                showResult('success', ['処理が完了しました。ZIPファイルをダウンロードしています。']);
            } else {
                const data = await resp.json();
                showResult('error', [data.error || '処理に失敗しました。']);
            }
        } catch (err) {
            let message = err.message;
            if (err.name === 'AbortError') {
                message = 'リクエストがタイムアウトしました。サーバーが混雑している可能性があります。';
            }
            showResult('error', ['エラー: ' + message]);
        } finally {
            spinner.style.display = 'none';
            submitBtn.disabled = false;
            if (spinnerText) spinnerText.textContent = '処理中...';
        }
    });

    function showResult(type, messages) {
        resultArea.style.display = 'block';
        const alert = document.getElementById('resultAlert');
        alert.className = 'alert ' + (type === 'success' ? 'alert-success' : 'alert-danger');
        alert.innerHTML = messages.map(m => `<p class="mb-1">${m}</p>`).join('');
    }

    // Pre-warm service when page loads (in background)
    wakeUpService().then(ok => {
        if (ok) console.log('Service is awake');
    });
});
