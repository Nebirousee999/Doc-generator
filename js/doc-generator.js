(function () {
    'use strict';

    const DOCX_OPTIONS = {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: { start: '{{', end: '}}' },
        nullGetter: () => ''
    };

    const state = {
        excelData: [],
        excelColumns: [],
        excelFileName: '',
        templateBuffer: null,
        templateFileName: '',
        templateTags: [],
        selectedFilenameTags: [],
        filenameSeparator: ' '
    };

    const excelDropZone = document.getElementById('excelDropZone');
    const excelInput = document.getElementById('excelInput');
    const excelInfo = document.getElementById('excelInfo');
    const excelError = document.getElementById('excelError');

    const wordDropZone = document.getElementById('wordDropZone');
    const wordInput = document.getElementById('wordInput');
    const wordInfo = document.getElementById('wordInfo');
    const wordError = document.getElementById('wordError');
    const tagsContainer = document.getElementById('tagsContainer');
    const tagsList = document.getElementById('tagsList');
    const columnsList = document.getElementById('columnsList');

    const filenameConfig = document.getElementById('filenameConfig');
    const tagToggles = document.getElementById('tagToggles');
    const filenamePreview = document.getElementById('filenamePreview');
    const separatorButtons = document.querySelectorAll('.sep-btn');

    const generateBtn = document.getElementById('generateBtn');
    const resetBtn = document.getElementById('resetBtn');
    const successMsg = document.getElementById('successMsg');
    const generateError = document.getElementById('generateError');

    // ===== Стоимость и лимиты =====
    function updateGenerateCostBadge() {
        const badge = document.getElementById('generateCostBadge');
        if (!badge) return;
        if (state.excelData.length === 0) {
            badge.textContent = '\ud83d\udc8e \u2014';
            return;
        }
        const cost = PRICING.costs.generateBatch || 1;
        badge.textContent = '\ud83d\udc8e ' + cost;
    }

    function getGenLimit(key, fallback) {
        return (window.PRICING && PRICING.limits && PRICING.limits[key]) || fallback;
    }

    function showError(el, msg) { el.textContent = msg; el.classList.add('visible'); }
    function hideError(el) { el.classList.remove('visible'); }
    function showSuccess(el, msg) { el.textContent = msg; el.classList.add('visible'); }
    function hideSuccess(el) { el.classList.remove('visible'); }

    function formatDocxError(err) {
        if (err && err.properties && Array.isArray(err.properties.errors)) {
            const details = err.properties.errors.map((e, i) => {
                const name = e.properties && e.properties.id ? e.properties.id : 'ошибка';
                const expl = e.properties && e.properties.explanation
                    ? e.properties.explanation
                    : (e.message || 'неизвестно');
                const ctx = e.properties && e.properties.context
                    ? ` (контекст: "${e.properties.context}")` : '';
                return `${i + 1}. ${name}: ${expl}${ctx}`;
            }).join('\n');
            return 'Проблемы в шаблоне:\n' + details;
        }
        return err.message || String(err);
    }

    function extractTagsFromDocx(buffer) {
        const zip = new PizZip(buffer);
        const files = [
            'word/document.xml',
            'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
            'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml'
        ];
        const tags = new Set();

        files.forEach(name => {
            const file = zip.file(name);
            if (!file) return;
            const xml = file.asText();
            const textParts = xml.match(/<w:t[^>]*>[\s\S]*?<\/w:t>/g) || [];
            const fullText = textParts.map(t => t.replace(/<[^>]+>/g, '')).join('');
            const matches = fullText.matchAll(/\{\{([^{}]+)\}\}/g);
            for (const m of matches) tags.add(m[1].trim());
        });

        return [...tags];
    }

    function getMatchingTags() {
        return state.templateTags.filter(t => state.excelColumns.includes(t));
    }

    function updateStepStates() {
        const step1 = document.getElementById('step1');
        const step2 = document.getElementById('step2');
        const step3 = document.getElementById('step3');

        const hasExcel = state.excelData.length > 0;
        const hasWord = !!state.templateBuffer;

        // Step 1 — всегда видим
        step1.style.display = '';

        // Step 2 — показываем только после загрузки Excel
        step2.style.display = hasExcel ? '' : 'none';

        // Step 3 — показываем только когда и Excel, и Word загружены
        step3.style.display = (hasExcel && hasWord) ? '' : 'none';

        // Классы active/done
        if (hasExcel) {
            step1.classList.remove('active');
            step1.classList.add('done');
        } else {
            step1.classList.add('active');
            step1.classList.remove('done');
        }

        if (hasWord) {
            step2.classList.remove('active');
            step2.classList.add('done');
            step3.classList.add('active');
        } else if (hasExcel) {
            step2.classList.add('active');
            step2.classList.remove('done');
            step3.classList.remove('active', 'done');
        } else {
            step2.classList.remove('active', 'done');
            step3.classList.remove('active', 'done');
        }

        generateBtn.disabled = !(hasExcel && hasWord);
        updateGenerateCostBadge();
    }

    // ===== Drag & Drop Excel =====
    excelDropZone.addEventListener('click', () => excelInput.click());
    excelDropZone.addEventListener('dragover', (e) => {
        e.preventDefault(); excelDropZone.classList.add('dragover');
    });
    excelDropZone.addEventListener('dragleave', () => excelDropZone.classList.remove('dragover'));
    excelDropZone.addEventListener('drop', (e) => {
        e.preventDefault(); excelDropZone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) handleExcelFile(file);
    });
    excelInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleExcelFile(file);
    });

    // ===== Drag & Drop Word =====
    wordDropZone.addEventListener('click', () => wordInput.click());
    wordDropZone.addEventListener('dragover', (e) => {
        e.preventDefault(); wordDropZone.classList.add('dragover');
    });
    wordDropZone.addEventListener('dragleave', () => wordDropZone.classList.remove('dragover'));
    wordDropZone.addEventListener('drop', (e) => {
        e.preventDefault(); wordDropZone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) handleWordFile(file);
    });
    wordInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleWordFile(file);
    });

    function handleExcelFile(file) {
        hideError(excelError);
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

                if (rows.length < 2) {
                    showError(excelError, 'Файл должен содержать минимум одну строку с данными (кроме заголовков).');
                    return;
                }

                state.excelColumns = rows[0].map(h => String(h).trim());
                state.excelData = rows.slice(1).map(row => {
                    const obj = {};
                    state.excelColumns.forEach((col, i) => {
                        obj[col] = row[i] !== undefined ? String(row[i]).trim() : '';
                    });
                    return obj;
                }).filter(obj => Object.values(obj).some(v => v !== ''));

                const maxRows = getGenLimit('maxExcelRows', 2000);
                if (state.excelData.length > maxRows) {
                    showError(excelError,
                        'Максимум ' + maxRows +
                        ' строк в Excel. У вас ' +
                        state.excelData.length + '. Разделите таблицу на части.');
                    state.excelData = [];
                    state.excelColumns = [];
                    state.excelFileName = '';
                    excelDropZone.classList.remove('has-file');
                    excelInfo.textContent = '';
                    updateStepStates();
                    return;
                }

                state.excelFileName = file.name;
                excelDropZone.classList.add('has-file');
                excelInfo.textContent = `✅ ${file.name} — ${state.excelData.length} строк, ${state.excelColumns.length} колонок`;
                excelDropZone.querySelector('.drop-zone-text').innerHTML =
                    `📊 <strong>${file.name}</strong> — заменить файл`;

                syncDefaultFilenameTags();
                updateStepStates();
                updateTagMatching();
                renderTagToggles();
                updateFilenamePreview();
            } catch (err) {
                console.error(err);
                showError(excelError, 'Не удалось прочитать Excel-файл: ' + err.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // DOC → DOCX через LibreOffice
    async function libreOfficeToDocx(file) {
        if (!window.__libreOfficeReady) {
            if (window.__libreOfficeError) {
                throw new Error('LibreOffice не запустился: ' + window.__libreOfficeError.message);
            }
            await new Promise((resolve, reject) => {
                const timeout = setTimeout(
                    () => reject(new Error('LibreOffice не ответил за 240 секунд')),
                    240000
                );
                window.addEventListener('libreoffice-ready', () => {
                    clearTimeout(timeout);
                    resolve();
                }, { once: true });
            });
        }
        const bytes = new Uint8Array(await file.arrayBuffer());
        const result = await window.LibreOfficeConverter.convert(
            bytes,
            { outputFormat: 'docx' },
            file.name
        );
        return new Blob([result.data], {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
    }

    async function handleWordFile(file) {
        hideError(wordError);
        const lower = file.name.toLowerCase();
        if (!lower.endsWith('.docx') && !lower.endsWith('.doc')) {
            showError(wordError, 'Поддерживаются только .docx и .doc.');
            return;
        }

        let docxFile = file;
        if (lower.endsWith('.doc')) {
            try {
                wordInfo.textContent = '⏳ Конвертация .doc → .docx через LibreOffice...';
                const blob = await libreOfficeToDocx(file);
                docxFile = new File(
                    [blob],
                    file.name.replace(/\.doc$/i, '.docx'),
                    { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
                );
            } catch (err) {
                console.error(err);
                showError(wordError, 'Не удалось конвертировать .doc: ' + err.message);
                return;
            }
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                state.templateBuffer = e.target.result;
                state.templateFileName = docxFile.name;

                const zip = new PizZip(state.templateBuffer);
                new window.docxtemplater(zip, DOCX_OPTIONS);

                state.templateTags = extractTagsFromDocx(state.templateBuffer);

                wordDropZone.classList.add('has-file');
                wordInfo.textContent = `✅ ${docxFile.name} — найдено тегов: ${state.templateTags.length}`;
                wordDropZone.querySelector('.drop-zone-text').innerHTML =
                    `📝 <strong>${docxFile.name}</strong> — заменить файл`;

                syncDefaultFilenameTags();
                updateStepStates();
                updateTagMatching();
                renderTagToggles();
                updateFilenamePreview();
            } catch (err) {
                console.error(err);
                showError(wordError, 'Не удалось прочитать шаблон:\n' + formatDocxError(err));
                state.templateBuffer = null;
                state.templateTags = [];
                updateStepStates();
                renderTagToggles();
            }
        };
        reader.readAsArrayBuffer(docxFile);
    }

    function syncDefaultFilenameTags() {
        const matching = getMatchingTags();
        if (state.selectedFilenameTags.length === 0) {
            state.selectedFilenameTags = matching.slice(0, 2);
        } else {
            state.selectedFilenameTags = state.selectedFilenameTags.filter(t => matching.includes(t));
        }
    }

    function updateTagMatching() {
        if (state.templateTags.length === 0 && state.excelColumns.length === 0) {
            tagsContainer.classList.remove('visible');
            return;
        }
        tagsContainer.classList.add('visible');

        tagsList.innerHTML = state.templateTags.map(tag => {
            const matched = state.excelColumns.includes(tag);
            return `<span class="tag ${matched ? 'matched' : 'unmatched'}">${escapeHtml(tag)}</span>`;
        }).join('') || '<span style="color:#a0aec0;font-size:13px;">Теги не найдены</span>';

        columnsList.innerHTML = state.excelColumns.map(col => {
            const matched = state.templateTags.includes(col);
            return `<span class="tag ${matched ? 'matched' : ''}">${escapeHtml(col)}</span>`;
        }).join('') || '<span style="color:#a0aec0;font-size:13px;">Колонки не найдены</span>';
    }

    function renderTagToggles() {
        const matching = getMatchingTags();
        if (matching.length === 0) {
            filenameConfig.classList.remove('visible');
            tagToggles.innerHTML = '';
            return;
        }
        filenameConfig.classList.add('visible');

        tagToggles.innerHTML = matching.map(tag => {
            const active = state.selectedFilenameTags.includes(tag);
            return `<button type="button" class="tag-toggle ${active ? 'active' : ''}" data-tag="${escapeHtml(tag)}">${escapeHtml(tag)}</button>`;
        }).join('');

        tagToggles.querySelectorAll('.tag-toggle').forEach(btn => {
            btn.addEventListener('click', () => {
                const tag = btn.dataset.tag;
                const idx = state.selectedFilenameTags.indexOf(tag);
                if (idx === -1) {
                    state.selectedFilenameTags.push(tag);
                } else {
                    state.selectedFilenameTags.splice(idx, 1);
                }
                btn.classList.toggle('active');
                updateFilenamePreview();
            });
        });
    }

    const SEP_MAP = { 'space': ' ', '_': '_', '-': '-' };
    separatorButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const key = btn.dataset.sep;
            state.filenameSeparator = SEP_MAP[key] || ' ';
            separatorButtons.forEach(b => b.classList.toggle('active', b === btn));
            updateFilenamePreview();
        });
    });

    function updateFilenamePreview() {
        if (state.excelData.length === 0) {
            filenamePreview.innerHTML = 'Пример имени: <strong>—</strong>';
            return;
        }
        const sample = buildFilename(state.excelData[0], 0);
        filenamePreview.innerHTML = 'Пример имени: <strong>' + escapeHtml(sample) + '</strong>';
    }

    function buildFilename(rowData, index) {
        const fallback = `document_${index + 1}.docx`;
        if (state.selectedFilenameTags.length === 0) return fallback;

        const parts = state.selectedFilenameTags
            .map(tag => (rowData[tag] || '').trim())
            .filter(v => v !== '');

        if (parts.length === 0) return fallback;

        const sep = state.filenameSeparator || ' ';
        const safeName = parts
            .join(sep)
            .replace(/[<>:"/\\|?*]/g, '_')
            .trim()
            .substring(0, 100);

        return safeName ? `${safeName}.docx` : fallback;
    }

    function escapeHtml(str) {
        return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;')
            .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    generateBtn.addEventListener('click', async () => {
        hideError(generateError);
        hideSuccess(successMsg);
        if (!state.templateBuffer || state.excelData.length === 0) return;

        const plannedCost = PRICING.costs.generateBatch || 1;

        if (window.Credits && plannedCost > 0) {
            const canPay = await Credits.canPerform(plannedCost);
            if (!canPay) {
                Credits.openPurchaseModal();
                return;
            }
        }

        generateBtn.disabled = true;
        generateBtn.innerHTML = '<span class="spinner"></span> Генерация...';

        try {
            const zip = new JSZip();
            let successCount = 0;
            let errorCount = 0;
            let firstError = null;

            for (let i = 0; i < state.excelData.length; i++) {
                const rowData = state.excelData[i];
                try {
                    const templateZip = new PizZip(state.templateBuffer);
                    const doc = new window.docxtemplater(templateZip, DOCX_OPTIONS);
                    doc.setData(rowData);
                    doc.render();

                    const output = doc.getZip().generate({
                        type: 'blob',
                        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    });

                    const fileName = buildFilename(rowData, i);

                    let finalName = fileName;
                    let counter = 1;
                    while (zip.file(finalName)) {
                        finalName = fileName.replace('.docx', `_${counter}.docx`);
                        counter++;
                    }

                    zip.file(finalName, output);
                    successCount++;
                } catch (rowErr) {
                    console.error(`Ошибка в строке ${i + 1}:`, rowErr);
                    if (!firstError) firstError = rowErr;
                    errorCount++;
                }
            }

            if (successCount === 0) {
                throw new Error(
                    'Не удалось сгенерировать ни одного документа.\n' +
                    (firstError ? formatDocxError(firstError) : 'Проверьте шаблон.')
                );
            }

            const content = await zip.generateAsync({ type: 'blob' });
            const url = URL.createObjectURL(content);
            const a = document.createElement('a');
            a.href = url;
            a.download = `documents_${new Date().toISOString().slice(0, 10)}.zip`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            const actualCost = PRICING.costs.generateBatch || 1;
            if (window.Credits && actualCost > 0) {
                await Credits.spend(actualCost, 'generateDoc');
            }

            showSuccess(successMsg,
                `✅ Готово! Создано ${successCount} документов` +
                (errorCount > 0 ? ` (ошибок: ${errorCount})` : '') +
                `. Архив скачан.`
            );
        } catch (err) {
            console.error(err);
            showError(generateError, 'Ошибка генерации:\n' + err.message);
        } finally {
            generateBtn.disabled = false;
            generateBtn.innerHTML = '🚀 Сгенерировать документы <span class="cost-badge" id="generateCostBadge">💎 —</span>';
            updateGenerateCostBadge();
        }
    });

    resetBtn.addEventListener('click', () => {
        state.excelData = [];
        state.excelColumns = [];
        state.excelFileName = '';
        state.templateBuffer = null;
        state.templateFileName = '';
        state.templateTags = [];
        state.selectedFilenameTags = [];
        state.filenameSeparator = ' ';

        excelInput.value = '';
        wordInput.value = '';
        excelDropZone.classList.remove('has-file');
        wordDropZone.classList.remove('has-file');
        excelDropZone.querySelector('.drop-zone-text').innerHTML =
            'Перетащите файл сюда или <strong>выберите на компьютере</strong>';
        wordDropZone.querySelector('.drop-zone-text').innerHTML =
            'Перетащите файл <strong>.docx</strong> или <strong>.doc</strong> сюда или выберите на компьютере';
        excelInfo.textContent = '';
        wordInfo.textContent = '';
        tagsContainer.classList.remove('visible');
        filenameConfig.classList.remove('visible');
        tagToggles.innerHTML = '';
        filenamePreview.innerHTML = 'Пример имени: <strong>—</strong>';
        hideError(excelError);
        hideError(wordError);
        hideError(generateError);
        hideSuccess(successMsg);

        separatorButtons.forEach(b => b.classList.toggle('active', b.dataset.sep === 'space'));

        updateStepStates();
    });

    updateStepStates();

})();