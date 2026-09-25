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
        templateTags: []
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

    const generateBtn = document.getElementById('generateBtn');
    const resetBtn = document.getElementById('resetBtn');
    const successMsg = document.getElementById('successMsg');
    const generateError = document.getElementById('generateError');

    function showError(el, msg) {
        el.textContent = msg;
        el.classList.add('visible');
    }
    function hideError(el) { el.classList.remove('visible'); }
    function showSuccess(el, msg) {
        el.textContent = msg;
        el.classList.add('visible');
    }
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

    // Извлекаем теги напрямую из XML внутри .docx — надёжно работает независимо от версии docxtemplater
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

            // Склеиваем все текстовые узлы <w:t> в один текст.
            // Это важно: Word часто разбивает {{Имя}} на части в разных run-ах.
            const textParts = xml.match(/<w:t[^>]*>[\s\S]*?<\/w:t>/g) || [];
            const fullText = textParts
                .map(t => t.replace(/<[^>]+>/g, ''))
                .join('');

            const matches = fullText.matchAll(/\{\{([^{}]+)\}\}/g);
            for (const m of matches) {
                tags.add(m[1].trim());
            }
        });

        return [...tags];
    }

    function updateStepStates() {
        const step1 = document.getElementById('step1');
        const step2 = document.getElementById('step2');
        const step3 = document.getElementById('step3');

        if (state.excelData.length > 0) {
            step1.classList.remove('active');
            step1.classList.add('done');
            step2.classList.add('active');
        } else {
            step1.classList.add('active');
            step1.classList.remove('done');
            step2.classList.remove('active', 'done');
            step3.classList.remove('active', 'done');
        }

        if (state.templateBuffer) {
            step2.classList.remove('active');
            step2.classList.add('done');
            step3.classList.add('active');
        } else if (state.excelData.length > 0) {
            step2.classList.add('active');
            step2.classList.remove('done');
            step3.classList.remove('active', 'done');
        }

        if (state.excelData.length > 0 && state.templateBuffer) {
            step3.classList.add('active');
        } else {
            step3.classList.remove('active', 'done');
        }

        generateBtn.disabled = !(state.excelData.length > 0 && state.templateBuffer);
    }

    // Drag & Drop Excel
    excelDropZone.addEventListener('click', () => excelInput.click());
    excelDropZone.addEventListener('dragover', (e) => {
        e.preventDefault(); excelDropZone.classList.add('dragover');
    });
    excelDropZone.addEventListener('dragleave', () => {
        excelDropZone.classList.remove('dragover');
    });
    excelDropZone.addEventListener('drop', (e) => {
        e.preventDefault(); excelDropZone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) handleExcelFile(file);
    });
    excelInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleExcelFile(file);
    });

    // Drag & Drop Word
    wordDropZone.addEventListener('click', () => wordInput.click());
    wordDropZone.addEventListener('dragover', (e) => {
        e.preventDefault(); wordDropZone.classList.add('dragover');
    });
    wordDropZone.addEventListener('dragleave', () => {
        wordDropZone.classList.remove('dragover');
    });
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

                state.excelFileName = file.name;
                excelDropZone.classList.add('has-file');
                excelInfo.textContent = `✅ ${file.name} — ${state.excelData.length} строк, ${state.excelColumns.length} колонок`;
                excelDropZone.querySelector('.drop-zone-text').innerHTML =
                    `📊 <strong>${file.name}</strong> — заменить файл`;

                updateStepStates();
                updateTagMatching();
            } catch (err) {
                console.error(err);
                showError(excelError, 'Не удалось прочитать Excel-файл: ' + err.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function handleWordFile(file) {
        hideError(wordError);
        if (!file.name.toLowerCase().endsWith('.docx')) {
            showError(wordError, 'Поддерживается только формат .docx (не .doc, не .rtf).');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                state.templateBuffer = e.target.result;
                state.templateFileName = file.name;

                // Проверяем валидность через docxtemplater (чтобы отловить битые шаблоны)
                const zip = new PizZip(state.templateBuffer);
                new window.docxtemplater(zip, DOCX_OPTIONS);

                // Извлекаем теги напрямую из XML
                state.templateTags = extractTagsFromDocx(state.templateBuffer);

                wordDropZone.classList.add('has-file');
                wordInfo.textContent = `✅ ${file.name} — найдено тегов: ${state.templateTags.length}`;
                wordDropZone.querySelector('.drop-zone-text').innerHTML =
                    `📝 <strong>${file.name}</strong> — заменить файл`;

                updateStepStates();
                updateTagMatching();
            } catch (err) {
                console.error(err);
                showError(wordError, 'Не удалось прочитать шаблон:\n' + formatDocxError(err));
                state.templateBuffer = null;
                state.templateTags = [];
                updateStepStates();
            }
        };
        reader.readAsArrayBuffer(file);
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

    function escapeHtml(str) {
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    generateBtn.addEventListener('click', async () => {
        hideError(generateError);
        hideSuccess(successMsg);
        if (!state.templateBuffer || state.excelData.length === 0) return;

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

                    let fileName = `document_${i + 1}.docx`;
                    if (state.excelColumns.length > 0) {
                        const firstColValue = rowData[state.excelColumns[0]];
                        if (firstColValue && firstColValue.trim()) {
                            const safeName = firstColValue.trim()
                                .replace(/[<>:"/\\|?*]/g, '_').substring(0, 100);
                            fileName = `${safeName}.docx`;
                        }
                    }

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
            generateBtn.innerHTML = '🚀 Сгенерировать документы';
        }
    });

    resetBtn.addEventListener('click', () => {
        state.excelData = [];
        state.excelColumns = [];
        state.excelFileName = '';
        state.templateBuffer = null;
        state.templateFileName = '';
        state.templateTags = [];

        excelInput.value = '';
        wordInput.value = '';
        excelDropZone.classList.remove('has-file');
        wordDropZone.classList.remove('has-file');
        excelDropZone.querySelector('.drop-zone-text').innerHTML =
            'Перетащите файл сюда или <strong>выберите на компьютере</strong>';
        wordDropZone.querySelector('.drop-zone-text').innerHTML =
            'Перетащите файл <strong>.docx</strong> сюда или выберите на компьютере';
        excelInfo.textContent = '';
        wordInfo.textContent = '';
        tagsContainer.classList.remove('visible');
        hideError(excelError);
        hideError(wordError);
        hideError(generateError);
        hideSuccess(successMsg);

        updateStepStates();
    });

    updateStepStates();
})();