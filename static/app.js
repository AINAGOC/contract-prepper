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
            const resp = await fetch('/process', { method: 'POST', body: formData });

            if (resp.ok) {
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
            showResult('error', ['通信エラーが発生しました: ' + err.message]);
        } finally {
            spinner.style.display = 'none';
            submitBtn.disabled = false;
        }
    });

    function showResult(type, messages) {
        resultArea.style.display = 'block';
        const alert = document.getElementById('resultAlert');
        alert.className = 'alert ' + (type === 'success' ? 'alert-success' : 'alert-danger');
        alert.innerHTML = messages.map(m => `<p class="mb-1">${m}</p>`).join('');
    }
});
