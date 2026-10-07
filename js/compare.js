(function () {
    'use strict';

    if (window.pdfjsLib) {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc =
            'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
    }

    const state = {
        fileA: null,
        fileB: null,
        diffParts: null,
        structuralResult: null,
        risksResult: null,
        mode: 'line',
        compareMode: 'compare',
        running: false
    };

    // ===== DOM =====
    const cmpDropA = document.getElementById('cmpDropA');
    const cmpDropB = document.getElementById('cmpDropB');
    const cmpInputA = document.getElementById('cmpInputA');
    const cmpInputB = document.getElementById('cmpInputB');
    const cmpInfoA = document.getElementById('cmpInfoA');
    const cmpInfoB = document.getElementById('cmpInfoB');
    const cmpError = document.getElementById('cmpError');
    const cmpError2 = document.getElementById('cmpError2');
    const cmpStep2 = document.getElementById('cmpStep2');
    const cmpStep3 = document.getElementById('cmpStep3');
    const cmpStartBtn = document.getElementById('cmpStartBtn');
    const cmpResetBtn = document.getElementById('cmpResetBtn');
    const cmpProgress = document.getElementById('cmpProgress');
    const cmpSummary = document.getElementById('cmpSummary');
    const cmpResult = document.getElementById('cmpResult');
    const cmpShowSame = document.getElementById('cmpShowSame');
    const cmpHighlightDanger = document.getElementById('cmpHighlightDanger');
    const cmpViewLine = document.getElementById('cmpViewLine');
    const cmpViewStructure = document.getElementById('cmpViewStructure');
    const cmpStructList = document.getElementById('cmpStructList');
    const cmpStructSummary = document.getElementById('cmpStructSummary');
    const cmpDownloadDocx = document.getElementById('cmpDownloadDocx');
    const cmpModeTabs = document.querySelectorAll('.cmp-mode-tab');
    const cmpViewRisks = document.getElementById('cmpViewRisks');
    const cmpRisksList = document.getElementById('cmpRisksList');
    const cmpRisksSummary = document.getElementById('cmpRisksSummary');
    const cmpRisksCriticalOnly = document.getElementById('cmpRisksCriticalOnly');
    const cmpDownloadRisksDocx = document.getElementById('cmpDownloadRisksDocx');

    // ===== Утилиты UI =====
    function showError(el, msg) { if (el) { el.textContent = msg; el.classList.add('visible'); } }
    function hideError(el) { if (el) el.classList.remove('visible'); }
    function showProgress(el, msg) { if (el) { el.textContent = msg; el.classList.add('visible'); } }
    function hideProgress(el) { if (el) el.classList.remove('visible'); }

    function formatSize(bytes) {
        if (bytes < 1024) return bytes + ' Б';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' КБ';
        return (bytes / 1024 / 1024).toFixed(2) + ' МБ';
    }

    function getLimit(key, fallback) {
        return (window.PRICING && PRICING.limits && PRICING.limits[key]) || fallback;
    }

    function getCost() {
        return (window.PRICING && PRICING.costs && PRICING.costs.comparePdf) || 3;
    }

    function getRiskCost() {
        return (window.PRICING && PRICING.costs && PRICING.costs.compareRisks) || 1;
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    // ===== Извлечение текста из PDF =====
    function reconstructPageText(items) {
        const lines = new Map();
        for (const it of items) {
            if (!it.str || it.str.trim() === '') continue;
            const y = Math.round(it.transform[5]);
            const x = it.transform[4];
            let key = y;
            for (const ky of lines.keys()) {
                if (Math.abs(ky - y) <= 2) { key = ky; break; }
            }
            if (!lines.has(key)) lines.set(key, []);
            lines.get(key).push({ x, str: it.str });
        }
        const sorted = [...lines.entries()].sort((a, b) => b[0] - a[0]);
        return sorted.map(([, parts]) =>
            parts.sort((a, b) => a.x - b.x)
                .map(p => p.str)
                .join(' ')
                .replace(/\s+/g, ' ')
                .trim()
        ).filter(Boolean).join('\n');
    }

    async function extractPdfText(source, onProgress) {
        const buffer = source instanceof Blob
            ? await source.arrayBuffer()
            : await source.arrayBuffer();
        const pdf = await window.pdfjsLib.getDocument({ data: buffer }).promise;
        const maxPages = getLimit('maxComparePages', 100);
        if (pdf.numPages > maxPages) {
            throw new Error('PDF содержит ' + pdf.numPages + ' стр. Максимум ' + maxPages + '.');
        }
        const parts = [];
        for (let i = 1; i <= pdf.numPages; i++) {
            if (onProgress) onProgress('страница ' + i + ' из ' + pdf.numPages);
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            parts.push(reconstructPageText(content.items));
        }
        return parts.join('\n');
    }

    // Извлечение текста из DOCX
    async function extractDocxText(file) {
        const buffer = await file.arrayBuffer();
        const zip = new PizZip(buffer);
        const filesToCheck = [
            'word/document.xml',
            'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
            'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml'
        ];
        const out = [];
        for (const name of filesToCheck) {
            const entry = zip.file(name);
            if (!entry) continue;
            let xml = entry.asText();
            xml = xml
                .replace(/<\/w:p>/g, '\n')
                .replace(/<w:br\s*\/?>/g, '\n')
                .replace(/<w:tab\s*\/?>/g, '    ');
            xml = xml.replace(/<[^>]+>/g, '');
            xml = xml.replace(/&amp;/g, '&')
                     .replace(/&lt;/g, '<')
                     .replace(/&gt;/g, '>')
                     .replace(/&quot;/g, '"')
                     .replace(/&#39;/g, "'")
                     .replace(/&nbsp;/g, ' ');
            out.push(xml);
        }
        return out.join('\n\n');
    }

    // DOC → PDF через LibreOffice → извлечение pdf.js
    async function extractDocText(file, onProgress) {
        if (!window.__libreOfficeReady) {
            if (window.__libreOfficeError) {
                throw new Error('LibreOffice не запустился: ' + window.__libreOfficeError.message);
            }
            if (onProgress) onProgress('инициализация LibreOffice (20–60 сек)...');
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
        if (onProgress) onProgress('конвертация DOC → PDF...');
        const bytes = new Uint8Array(await file.arrayBuffer());
        const result = await window.LibreOfficeConverter.convert(
            bytes,
            { outputFormat: 'pdf' },
            file.name
        );
        const blob = new Blob([result.data], { type: 'application/pdf' });
        return await extractPdfText(blob, onProgress);
    }

    async function extractTextFromFile(file, onProgress) {
        const name = file.name.toLowerCase();
        if (name.endsWith('.pdf')) return await extractPdfText(file, onProgress);
        if (name.endsWith('.docx')) return await extractDocxText(file);
        if (name.endsWith('.doc')) return await extractDocText(file, onProgress);
        throw new Error('Поддерживаются только PDF, DOCX, DOC.');
    }

    // ===== Нормализация =====
    const NOISE_LINE_RES = [
        /^ФОРМА(\s+УТВЕРЖДЕНА)?$/i,
        /^КОПИЯ(\s+ВЕРНА)?$/i,
        /^М\.?\s*П\.?$/i,
        /^-?\s*\d+\s*-?$/,
        /^стр\.?\s*\d+(\s*из\s*\d+)?\.?$/i,
        /^лист\s*\d+$/i,
        /^подпись$/i,
        /^\d+\s*\/\s*\d+$/,
        /^г\.?\s*$/i
    ];

    function isNoiseLine(line) {
        const t = line.trim();
        if (!t) return false;
        for (const re of NOISE_LINE_RES) {
            if (re.test(t)) return true;
        }
        return false;
    }

    function isJunkLine(line) {
        const t = line.replace(/\s/g, '');
        if (t.length < 4) return false;
        let junk = 0, useful = 0;
        for (let i = 0; i < t.length; i++) {
            const c = t.charCodeAt(i);
            if (c >= 0xA1 && c <= 0xFF) junk++;
            else if (c >= 0x0400 && c <= 0x04FF) useful++;
            else if ((c >= 0x41 && c <= 0x5A) || (c >= 0x61 && c <= 0x7A)) useful++;
        }
        return junk > 8 && junk > useful;
    }

    function ultraNorm(s) {
        if (!s) return '';
        return String(s)
            .toLowerCase()
            .replace(/ё/g, 'е')
            .replace(/[^\p{L}\p{N}]/gu, '');
    }

    function normalize(text) {
        let s = text;
        s = s.replace(/[\u200B-\u200F\u202A-\u202E\u2060\uFEFF\u00AD]/g, '');
        s = s.replace(/[«»""„""‘’‚‛]/g, '"');
        s = s.replace(/[‘’‚‛]/g, "'");
        s = s.replace(/[–—−]/g, '-');
        s = s.replace(/…/g, '...');
        s = s.replace(/[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]/g, ' ');
        s = s.replace(/\s+([.,;:!?])/g, '$1');
        s = s.replace(/(\d)\s+(?=\d)/g, '$1');
        s = s.replace(/[ \t]+/g, ' ');
        s = s.replace(/\r\n?/g, '\n');
        s = s.replace(/([а-яёa-z])-\n([а-яёa-z])/gi, '$1$2');

        let lines = s.split('\n')
            .map(l => l.trim())
            .filter(l => !isNoiseLine(l) && !isJunkLine(l));

        let joined = lines.join(' ').replace(/\s+/g, ' ').trim();

        // Защита аббревиатур — точка → маркер \u0001
        joined = joined.replace(/\b(п|г|ст|т|д|е|см|рис|табл|руб|коп|тыс|млн|млрд|им|тов|кв|стр|с)\.\s/gi, '$1\u0001');

        // Маркер \u0002 — начало нового пункта
        joined = joined.replace(
            /(^|[.!?:;]\s+)(\d{1,3}(?:\.\d{1,3}){0,5})\.\s+(?=[А-ЯЁA-Z«"0-9])/g,
            function (m, before, num) { return before + '\u0002' + num + '. '; }
        );

        const parts = joined.split('\u0002');

        return parts
            .map(p => p.replace(/\u0001/g, '. ').replace(/\s+/g, ' ').trim())
            .filter(Boolean)
            .join('\n');
    }

    // ===== Drop-zones =====
    function bindDrop(dropZone, input, slot) {
        dropZone.addEventListener('click', () => input.click());
        dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) await handleFile(file, slot, dropZone, input);
        });
        input.addEventListener('change', async (e) => {
            const file = e.target.files[0];
            input.value = '';
            if (file) await handleFile(file, slot, dropZone, input);
        });
    }

    async function handleFile(file, slot, dropZone, input) {
        hideError(cmpError);
        const lowerName = file.name.toLowerCase();
        if (!lowerName.endsWith('.pdf') && !lowerName.endsWith('.docx') && !lowerName.endsWith('.doc')) {
            showError(cmpError, 'Поддерживаются PDF, DOCX, DOC.');
            return;
        }
        const maxMB = getLimit('maxCompareFileSizeMB', 20);
        if (file.size > maxMB * 1024 * 1024) {
            showError(cmpError, 'Файл «' + file.name + '» больше ' + maxMB + ' МБ.');
            return;
        }

        const info = slot === 'A' ? cmpInfoA : cmpInfoB;
        info.textContent = '⏳ Извлечение текста из «' + file.name + '»...';
        dropZone.classList.add('has-file');
        dropZone.querySelector('.drop-zone-text').innerHTML =
            '📕 <strong>' + file.name + '</strong> — заменить';

        try {
            const text = await extractTextFromFile(file, (msg) => {
                info.textContent = '⏳ «' + file.name + '»: ' + msg;
            });

            const entry = { file, name: file.name, size: file.size, text };
            if (slot === 'A') state.fileA = entry;
            else state.fileB = entry;

            const chars = text.length;
            const lines = text.split('\n').length;
            info.textContent = '✅ ' + file.name + ' — ' + formatSize(file.size) +
                ', ~' + lines + ' строк, ' + chars + ' симв.';
            refreshUI();
        } catch (err) {
            console.error(err);
            showError(cmpError, 'Ошибка: ' + err.message);
            info.textContent = '';
            dropZone.classList.remove('has-file');
            dropZone.querySelector('.drop-zone-text').innerHTML =
                'Перетащите файл сюда или <strong>выберите</strong>';
            if (slot === 'A') state.fileA = null;
            else state.fileB = null;
            refreshUI();
        }
    }

    function refreshUI() {
        const hasA = !!state.fileA;
        const hasB = !!state.fileB;
        const readyCompare = hasA && hasB;
        const readyRisks = hasA || hasB;

        cmpStep2.classList.toggle('active', readyRisks);
        cmpStartBtn.disabled = !readyRisks;

        if (readyCompare) {
            state.compareMode = 'compare';
            cmpStartBtn.innerHTML = '⚖️ Сравнить <span class="cost-badge" id="cmpCostBadge">💎 ' + getCost() + '</span>';
        } else if (readyRisks) {
            state.compareMode = 'risks';
            cmpStartBtn.innerHTML = '⚠️ Посмотреть риски <span class="cost-badge" id="cmpCostBadge">💎 ' + getRiskCost() + '</span>';
        }

        if (readyRisks) {
            cmpStep3.style.display = 'none';
            cmpStep3.classList.remove('active', 'done');
            cmpResult.innerHTML = '';
        }
    }

    bindDrop(cmpDropA, cmpInputA, 'A');
    bindDrop(cmpDropB, cmpInputB, 'B');

    // ===== Паттерны опасных фрагментов =====
    const DANGER_PATTERNS = [
        /\d[\d\s]*\d\s*(?:руб|₽|коп|дней|дня|день|месяц|лет|год|года|%|процент|кВт)/gi,
        /\d{1,2}\.\d{1,2}\.\d{2,4}/g,
        /(?:ООО|ИП|АО|ПАО|ЗАО|ОАО)\s+["«][^"»]+["»]/g
    ];

    function containsDanger(s) {
        if (!s) return false;
        for (const re of DANGER_PATTERNS) {
            re.lastIndex = 0;
            if (re.test(s)) return true;
        }
        return false;
    }

    function dangerify(text, enabled) {
        const safe = escapeHtml(text);
        if (!enabled) return safe;
        let out = safe;
        for (const re of DANGER_PATTERNS) {
            out = out.replace(re, m => '<span class="cmp-danger">' + m + '</span>');
        }
        return out;
    }

    // ===== Паттерны рисков =====
    const RISK_PATTERNS = [
        {
            category: '💰 Финансовые обязательства',
            severity: 'high',
            re: /(?:штраф|неустойк|пен[ия]|пени)[^.!?]{0,250}?\d[\d\s]*\d\s*(?:руб|₽|процент|%)/gi,
            hint: 'Проверьте размер — несоразмерные штрафы можно оспорить в суде (ст. 333 ГК РФ).'
        },
        {
            category: '💰 Оплата — односторонняя',
            severity: 'high',
            re: /(?:в\s+одностороннем\s+порядке\s+измен|вправе\s+изменить\s+(?:цен|стоимость|размер))/gi,
            hint: 'Одностороннее изменение цены — невыгодное условие. Требуйте фиксации или согласования.'
        },
        {
            category: '⚖️ Ответственность снята',
            severity: 'high',
            re: /(?:не\s+нес[её]т\s+ответственности|освобождается\s+от\s+ответственности|ответственность\s+не\s+наступает)/gi,
            hint: 'Одна из сторон освобождена от ответственности. Проверьте, сбалансировано ли это.'
        },
        {
            category: '📝 Односторонние действия',
            severity: 'high',
            re: /(?:в\s+одностороннем\s+порядке|вправе\s+в\s+любое\s+время\s+отказаться|односторонне[ем]\s+расторж)/gi,
            hint: 'Односторонний отказ или изменение — риск для второй стороны. Требуйте симметрии.'
        },
        {
            category: '🔒 Конфиденциальность — штрафы',
            severity: 'high',
            re: /(?:штраф\s+за\s+разглашение|ответственность\s+за\s+разглашение|нарушением?\s+конфиденциальности)[^.!?]{0,200}?\d[\d\s]*\d\s*(?:руб|₽)/gi,
            hint: 'Большой штраф за разглашение. Проверьте, что объём обязательств реалистичен.'
        },
        {
            category: '⏰ Размытые сроки',
            severity: 'medium',
            re: /(?:в\s+разумный\s+срок|по\s+мере\s+необходимости|по\s+возможности|в\s+течение\s+разумного\s+времени|своевременно)/gi,
            hint: 'Размытые формулировки сроков — риск толкования в пользу контрагента. Замените на конкретные даты.'
        },
        {
            category: '🔄 Автопролонгация',
            severity: 'medium',
            re: /(?:автоматически\s+продлевается|пролонгируется|считается\s+пролонгированным|ежегодно\s+пролонгируемым)/gi,
            hint: 'Договор продлевается автоматически. Убедитесь, что порядок отказа от продления вам удобен.'
        },
        {
            category: '🏛️ Подсудность',
            severity: 'medium',
            re: /(?:по\s+месту\s+нахождения\s+(?:Истца|Ответчика|Принципала|Агента|Заказчика|Исполнителя|Поставщика|Покупателя))/gi,
            hint: 'Подсудность по месту нахождения одной стороны может быть невыгодна другой.'
        },
        {
            category: '📅 Претензионный срок',
            severity: 'medium',
            re: /претензи[яи][^.!?]{0,150}?(?:в\s+течение\s+\d+\s+(?:дней|дня|рабочих|календарных))/gi,
            hint: 'Короткий претензионный срок — риск не успеть подготовить и направить ответ.'
        },
        {
            category: '🚫 Запреты',
            severity: 'medium',
            re: /(?:не\s+вправе|запрещается|не\s+допускается)[^.!?]{0,180}/gi,
            hint: 'Односторонние запреты — проверьте, что они сбалансированы для обеих сторон.'
        },
        {
            category: '📎 Эксклюзивность',
            severity: 'medium',
            re: /(?:исключительн[аоы]|эксклюзивн[аоы]|не\s+вправе\s+заключать\s+(?:аналогичные\s+)?договор[ыа]?\s+с\s+третьими)/gi,
            hint: 'Условия эксклюзивности — проверьте, что это оправдано коммерчески.'
        },
        {
            category: '🕒 Молчаливое согласование',
            severity: 'low',
            re: /(?:считается\s+(?:согласованным|принятым|одобренным)|если\s+(?:не\s+поступило|отсутствует)\s+(?:возражени|ответ))/gi,
            hint: 'Условия автоматического согласования — можно пропустить важное уведомление.'
        }
    ];

    function analyzeRisks(text) {
        const clauses = text.split('\n').filter(Boolean);
        const risks = [];
        const seen = new Set();

        for (const clause of clauses) {
            const numMatch = clause.match(/^(\d+(?:\.\d+)*)\./);
            const clauseNum = numMatch ? numMatch[1] : null;

            for (const p of RISK_PATTERNS) {
                p.re.lastIndex = 0;
                let m;
                while ((m = p.re.exec(clause)) !== null) {
                    const quote = m[0].trim();
                    if (quote.length < 8) continue;
                    const key = p.category + '::' + (clauseNum || '—') + '::' + ultraNorm(quote).slice(0, 100);
                    if (seen.has(key)) continue;
                    seen.add(key);

                    risks.push({
                        category: p.category,
                        severity: p.severity,
                        hint: p.hint,
                        quote: quote,
                        clauseNum: clauseNum,
                        clauseText: clause
                    });
                    if (risks.length >= 200) return risks;
                }
            }
        }

        const sevRank = { high: 0, medium: 1, low: 2 };
        risks.sort((a, b) => sevRank[a.severity] - sevRank[b.severity]);
        return risks;
    }

    function mergeRisks(risksA, risksB) {
        const map = new Map();
        const add = (r, where) => {
            const key = r.category + '::' + (r.clauseNum || '—') + '::' + ultraNorm(r.quote).slice(0, 100);
            if (map.has(key)) {
                const cur = map.get(key);
                if (cur.where !== where) cur.where = 'both';
            } else {
                map.set(key, { ...r, where });
            }
        };
        risksA.forEach(r => add(r, 'a'));
        risksB.forEach(r => add(r, 'b'));

        const out = [...map.values()];
        const sevRank = { high: 0, medium: 1, low: 2 };
        const whereRank = { both: 0, a: 1, b: 2 };
        out.sort((x, y) => {
            const ds = sevRank[x.severity] - sevRank[y.severity];
            if (ds !== 0) return ds;
            return whereRank[x.where] - whereRank[y.where];
        });
        return out;
    }

    // ===== Diff =====
    function wordDiffHasChanges(wordDiff) {
        for (const w of wordDiff) {
            if (w.added || w.removed) {
                const stripped = w.value.replace(/[\s\p{P}\p{S}]/gu, '');
                if (stripped === '') continue;
                return true;
            }
        }
        return false;
    }

    function flattenDiff(parts) {
        const out = [];
        for (const part of parts) {
            const type = part.added ? 'added' : (part.removed ? 'removed' : 'same');
            const lines = part.value.split('\n');
            if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop();
            for (const line of lines) {
                out.push({ type, text: line });
            }
        }
        return out;
    }

    function pairChanged(flat) {
        const result = [];
        let i = 0;
        while (i < flat.length) {
            if (flat[i].type === 'same') {
                result.push({ left: flat[i], right: flat[i], type: 'same' });
                i++;
                continue;
            }
            const removedBlock = [];
            const addedBlock = [];
            while (i < flat.length && flat[i].type === 'removed') { removedBlock.push(flat[i]); i++; }
            while (i < flat.length && flat[i].type === 'added') { addedBlock.push(flat[i]); i++; }
            if (removedBlock.length === 0) {
                while (i < flat.length && flat[i].type === 'added') { addedBlock.push(flat[i]); i++; }
            }

            const usedL = new Set();
            const usedR = new Set();

            for (let li = 0; li < removedBlock.length; li++) {
                if (usedL.has(li)) continue;
                const ultraL = ultraNorm(removedBlock[li].text);
                for (let ri = 0; ri < addedBlock.length; ri++) {
                    if (usedR.has(ri)) continue;
                    if (removedBlock[li].text === addedBlock[ri].text ||
                        (ultraL && ultraL === ultraNorm(addedBlock[ri].text))) {
                        result.push({
                            left: removedBlock[li],
                            right: addedBlock[ri],
                            type: 'same'
                        });
                        usedL.add(li);
                        usedR.add(ri);
                        break;
                    }
                }
            }

            const restL = [];
            const restR = [];
            for (let li = 0; li < removedBlock.length; li++) {
                if (!usedL.has(li)) restL.push(removedBlock[li]);
            }
            for (let ri = 0; ri < addedBlock.length; ri++) {
                if (!usedR.has(ri)) restR.push(addedBlock[ri]);
            }

            const n = Math.max(restL.length, restR.length);
            for (let j = 0; j < n; j++) {
                const L = restL[j] || null;
                const R = restR[j] || null;
                if (L && R) {
                    const ultraL = ultraNorm(L.text);
                    const ultraR = ultraNorm(R.text);
                    if (ultraL === ultraR) {
                        result.push({ left: L, right: R, type: 'same' });
                    } else {
                        const wordDiff = window.Diff.diffWords(L.text, R.text);
                        if (!wordDiffHasChanges(wordDiff)) {
                            result.push({ left: L, right: R, type: 'same' });
                        } else {
                            result.push({ left: L, right: R, type: 'changed', wordDiff });
                        }
                    }
                } else if (L) {
                    result.push({ left: L, right: null, type: 'removed' });
                } else {
                    result.push({ left: null, right: R, type: 'added' });
                }
            }
        }
        return result;
    }

    function renderChangedCell(text, side, wordDiff, highlightDanger) {
        let html = '';
        for (const w of wordDiff) {
            const isRemoved = w.removed;
            const isAdded = w.added;
            const show = side === 'left' ? !isAdded : !isRemoved;
            if (!show) continue;
            let chunk = escapeHtml(w.value);
            if (highlightDanger && (isRemoved || isAdded)) {
                for (const re of DANGER_PATTERNS) {
                    chunk = chunk.replace(re, m => '<span class="cmp-danger">' + m + '</span>');
                }
            }
            if (isRemoved) html += '<span class="cmp-word-removed">' + chunk + '</span>';
            else if (isAdded) html += '<span class="cmp-word-added">' + chunk + '</span>';
            else html += chunk;
        }
        return html || '&nbsp;';
    }

    function renderResult() {
        if (!state.diffParts) return;
        const showSame = cmpShowSame.checked;
        const danger = cmpHighlightDanger.checked;

        const flat = flattenDiff(state.diffParts);
        const paired = pairChanged(flat);

        const parts = [];
        for (const row of paired) {
            const hide = row.type === 'same' && !showSame;
            if (hide) continue;

            let leftHtml, rightHtml;
            if (row.type === 'same') {
                leftHtml = escapeHtml(row.left.text) || '&nbsp;';
                rightHtml = leftHtml;
                parts.push(
                    '<div class="cmp-cell left">' + leftHtml + '</div>' +
                    '<div class="cmp-cell right">' + rightHtml + '</div>'
                );
            } else if (row.type === 'changed') {
                leftHtml = renderChangedCell(row.left.text, 'left', row.wordDiff, danger);
                rightHtml = renderChangedCell(row.right.text, 'right', row.wordDiff, danger);
                parts.push(
                    '<div class="cmp-cell left changed-old">' + leftHtml + '</div>' +
                    '<div class="cmp-cell right changed-new">' + rightHtml + '</div>'
                );
            } else if (row.type === 'removed') {
                const txt = danger ? dangerify(row.left.text, true) : escapeHtml(row.left.text);
                parts.push(
                    '<div class="cmp-cell left removed">' + (txt || '&nbsp;') + '</div>' +
                    '<div class="cmp-cell right empty"></div>'
                );
            } else if (row.type === 'added') {
                const txt = danger ? dangerify(row.right.text, true) : escapeHtml(row.right.text);
                parts.push(
                    '<div class="cmp-cell left empty"></div>' +
                    '<div class="cmp-cell right added">' + (txt || '&nbsp;') + '</div>'
                );
            }
        }

        cmpResult.innerHTML = parts.join('');
        return paired;
    }

    function buildSummary(paired) {
        let same = 0, added = 0, removed = 0, changed = 0, danger = 0;
        for (const row of paired) {
            if (row.type === 'same') same++;
            else if (row.type === 'added') {
                added++;
                if (containsDanger(row.right.text)) danger++;
            } else if (row.type === 'removed') {
                removed++;
                if (containsDanger(row.left.text)) danger++;
            } else if (row.type === 'changed') {
                changed++;
                if (containsDanger(row.left.text) || containsDanger(row.right.text)) danger++;
            }
        }
        const totalChanges = added + removed + changed;
        let text = 'Изменений: ' + totalChanges +
            ' — добавлено: ' + added +
            ', удалено: ' + removed +
            ', изменено: ' + changed +
            ', без изменений: ' + same;
        if (danger > 0) text += '. ⚠️ Требуют внимания: ' + danger + ' (суммы, даты, сроки)';
        return text;
    }

    // ===== Структурный diff =====
    function isSectionHeading(num, text) {
        if (num.includes('.')) return false;
        const t = (text || '').trim();
        if (t.length > 70) return false;
        return t === t.toUpperCase() || /^[А-ЯЁA-Z][А-ЯЁA-Z\s]{3,}$/.test(t);
    }

    function splitIntoClauses(text) {
        const lines = text.split('\n');
        const clauses = [];
        let current = null;
        let order = 0;

        const NUM_RE = /^(\d+(?:\.\d+)*)\.?\s+(.+)/;

        for (const rawLine of lines) {
            const line = rawLine.trim();
            if (!line) continue;

            const m = line.match(NUM_RE);
            if (m) {
                const num = m[1];
                const txt = m[2];
                const isHeading = isSectionHeading(num, txt);
                if (current) clauses.push(current);
                current = { num, text: txt, isHeading, order: order++ };
            } else {
                if (current) {
                    current.text += ' ' + line;
                } else {
                    current = { num: null, text: line, isHeading: false, order: order++ };
                }
            }
        }
        if (current) clauses.push(current);
        return clauses;
    }

    function cmpClauseNum(a, b) {
        if (a === null) return -1;
        if (b === null) return 1;
        const pa = String(a).split('.').map(n => parseInt(n, 10) || 0);
        const pb = String(b).split('.').map(n => parseInt(n, 10) || 0);
        for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
            const va = pa[i] || 0;
            const vb = pb[i] || 0;
            if (va !== vb) return va - vb;
        }
        return 0;
    }

    function structuralDiff(textA, textB) {
        const clausesA = splitIntoClauses(textA);
        const clausesB = splitIntoClauses(textB);

        const mapA = new Map();
        const mapB = new Map();
        clausesA.forEach(c => { if (c.num) mapA.set(c.num, c); });
        clausesB.forEach(c => { if (c.num) mapB.set(c.num, c); });

        const allNums = new Set([...mapA.keys(), ...mapB.keys()]);
        const result = [];

        for (const num of allNums) {
            const a = mapA.get(num);
            const b = mapB.get(num);
            if (a && b) {
                const ultraA = ultraNorm(a.text);
                const ultraB = ultraNorm(b.text);
                if (ultraA === ultraB) {
                    result.push({ num, type: 'same', textA: a.text, textB: b.text, isHeading: a.isHeading || b.isHeading });
                } else {
                    const wd = window.Diff.diffWords(a.text, b.text);
                    if (!wordDiffHasChanges(wd)) {
                        result.push({ num, type: 'same', textA: a.text, textB: b.text, isHeading: a.isHeading || b.isHeading });
                    } else {
                        result.push({ num, type: 'changed', textA: a.text, textB: b.text, isHeading: false, wordDiff: wd });
                    }
                }
            } else if (a) {
                result.push({ num, type: 'removed', textA: a.text, textB: null, isHeading: a.isHeading });
            } else if (b) {
                result.push({ num, type: 'added', textA: null, textB: b.text, isHeading: b.isHeading });
            }
        }

        const preA = clausesA.find(c => !c.num);
        const preB = clausesB.find(c => !c.num);
        if (preA || preB) {
            const tA = preA ? preA.text : '';
            const tB = preB ? preB.text : '';
            if (ultraNorm(tA) === ultraNorm(tB)) {
                result.unshift({ num: '—', type: 'same', textA: tA, textB: tB, isHeading: false });
            } else if (tA && tB) {
                result.unshift({ num: '—', type: 'changed', textA: tA, textB: tB, isHeading: false });
            } else if (tA) {
                result.unshift({ num: '—', type: 'removed', textA: tA, textB: null, isHeading: false });
            } else {
                result.unshift({ num: '—', type: 'added', textA: null, textB: tB, isHeading: false });
            }
        }

        result.sort((x, y) => cmpClauseNum(x.num === '—' ? null : x.num, y.num === '—' ? null : y.num));
        return result;
    }

    function renderStructuralResult() {
        if (!cmpStructList) return;
        if (!state.structuralResult) {
            cmpStructList.innerHTML = '<div class="cmp-struct-empty">Нет данных</div>';
            return;
        }

        const showSame = cmpShowSame.checked;
        const danger = cmpHighlightDanger.checked;
        const items = state.structuralResult;

        let added = 0, removed = 0, changed = 0, same = 0;
        items.forEach(c => {
            if (c.type === 'added') added++;
            else if (c.type === 'removed') removed++;
            else if (c.type === 'changed') changed++;
            else if (c.type === 'same') same++;
        });

        if (cmpStructSummary) {
            cmpStructSummary.textContent =
                'Пунктов: ' + items.length +
                ' — изменено: ' + changed +
                ', добавлено: ' + added +
                ', удалено: ' + removed +
                ', без изменений: ' + same;
        }

        const TYPE_LABEL = {
            changed: 'Изменён',
            added: 'Добавлен',
            removed: 'Удалён'
        };

        let html = '';
        for (const c of items) {
            if (c.type === 'same' && !showSame) continue;

            const cls = c.type === 'same' ? 'same' : c.type;
            const numLabel = c.num === '—' ? 'Преамбула' : 'п. ' + c.num;
            const preview = (c.textB || c.textA || '').slice(0, 120);
            const typeBadge = c.type !== 'same'
                ? '<span class="cmp-struct-type ' + cls + '">' + TYPE_LABEL[c.type] + '</span>'
                : '';

            html += '<div class="cmp-struct-item' + (c.type !== 'same' ? ' expanded' : '') + '">' +
                '<div class="cmp-struct-item-head">' +
                '<span class="cmp-struct-num">' + escapeHtml(numLabel) + '</span>' +
                typeBadge +
                '<span class="cmp-struct-preview">' + escapeHtml(preview) + '…</span>' +
                '<span class="cmp-struct-expand">▶</span>' +
                '</div>' +
                '<div class="cmp-struct-body">';

            if (c.type === 'changed' || c.type === 'removed') {
                let txtA = escapeHtml(c.textA || '');
                if (danger) {
                    for (const re of DANGER_PATTERNS) {
                        txtA = txtA.replace(re, m => '<span class="cmp-danger">' + m + '</span>');
                    }
                }
                html += '<div class="cmp-struct-cell a"><span class="cmp-cell-label">Было (A):</span><br>' +
                    (txtA || '—') + '</div>';
            }
            if (c.type === 'changed' || c.type === 'added') {
                let txtB = escapeHtml(c.textB || '');
                if (danger) {
                    for (const re of DANGER_PATTERNS) {
                        txtB = txtB.replace(re, m => '<span class="cmp-danger">' + m + '</span>');
                    }
                }
                html += '<div class="cmp-struct-cell b"><span class="cmp-cell-label">Стало (Б):</span><br>' +
                    (txtB || '—') + '</div>';
            }
            if (c.type === 'same') {
                html += '<div class="cmp-struct-cell b">' + escapeHtml(c.textB || c.textA || '') + '</div>';
            }

            html += '</div></div>';
        }

        cmpStructList.innerHTML = html;

        cmpStructList.querySelectorAll('.cmp-struct-item-head').forEach(h => {
            h.addEventListener('click', () => {
                h.parentNode.classList.toggle('expanded');
            });
        });
    }

    // ===== Рендер рисков =====
    function highlightQuoteInClause(clauseText, quote) {
        const safe = escapeHtml(clauseText);
        const safeQuote = escapeHtml(quote);
        const idx = safe.indexOf(safeQuote);
        if (idx === -1) {
            const shortQuote = quote.slice(0, 40);
            const idx2 = safe.indexOf(escapeHtml(shortQuote));
            if (idx2 === -1) return safe;
            return safe.slice(0, idx2) +
                '<span class="cmp-danger">' + escapeHtml(shortQuote) + '</span>' +
                safe.slice(idx2 + escapeHtml(shortQuote).length);
        }
        return safe.slice(0, idx) +
            '<span class="cmp-danger">' + safeQuote + '</span>' +
            safe.slice(idx + safeQuote.length);
    }

    function renderRisksResult() {
        if (!cmpRisksList) return;
        if (!state.risksResult || state.risksResult.length === 0) {
            cmpRisksList.innerHTML =
                '<div class="cmp-risks-empty"><strong>✅ Явных юридических рисков не обнаружено</strong>' +
                'Автоматический анализ не нашёл типичных «красных флагов». ' +
                'Это не значит, что договор безопасен — рекомендуется вычитка юристом.</div>';
            if (cmpRisksSummary) cmpRisksSummary.textContent = 'Рисков не найдено';
            return;
        }

        const criticalOnly = cmpRisksCriticalOnly && cmpRisksCriticalOnly.checked;
        let risks = state.risksResult;
        if (criticalOnly) risks = risks.filter(r => r.severity === 'high');

        if (risks.length === 0) {
            cmpRisksList.innerHTML =
                '<div class="cmp-risks-empty"><strong>Критичных рисков нет</strong>' +
                'Остальные риски скрыты фильтром. Снимите галочку, чтобы увидеть все.</div>';
            return;
        }

        let high = 0, medium = 0, low = 0;
        state.risksResult.forEach(r => {
            if (r.severity === 'high') high++;
            else if (r.severity === 'medium') medium++;
            else low++;
        });

        if (cmpRisksSummary) {
            cmpRisksSummary.textContent =
                'Найдено рисков: ' + state.risksResult.length +
                ' — критичных: ' + high +
                ', средних: ' + medium +
                ', низких: ' + low;
        }

        const SEV_LABEL = { high: 'Критично', medium: 'Средне', low: 'Низко' };
        const WHERE_LABEL = { both: 'в обоих', a: 'только в A', b: 'только в Б' };

        let html = '';
        for (const r of risks) {
            const whereCls = r.where === 'both' ? 'both' : r.where;
            const clauseLabel = r.clauseNum ? 'п. ' + r.clauseNum : 'Преамбула';
            const clauseHtml = highlightQuoteInClause(r.clauseText, r.quote);
            const whereHtml = (r.where && WHERE_LABEL[r.where])
                ? '<span class="cmp-risk-where ' + whereCls + '">' + WHERE_LABEL[r.where] + '</span>'
                : '';

            html += '<div class="cmp-risk-item ' + r.severity + '">' +
                '<div class="cmp-risk-head">' +
                '<span class="cmp-risk-clause">' + escapeHtml(clauseLabel) + '</span>' +
                '<span class="cmp-risk-category">' + escapeHtml(r.category) + '</span>' +
                '<span class="cmp-risk-severity ' + r.severity + '">' + SEV_LABEL[r.severity] + '</span>' +
                whereHtml +
                '</div>' +
                '<div class="cmp-risk-body">' +
                '<div class="cmp-risk-clause-text">' + clauseHtml + '</div>' +
                '<div class="cmp-risk-hint"><strong>Что делать:</strong> ' + escapeHtml(r.hint) + '</div>' +
                '</div></div>';
        }
        cmpRisksList.innerHTML = html;
    }

    // ===== Переключение режимов =====
    function switchCompareMode(mode) {
        state.mode = mode;
        cmpModeTabs.forEach(t => t.classList.toggle('active', t.dataset.cmpMode === mode));
        if (cmpViewLine) cmpViewLine.style.display = mode === 'line' ? 'block' : 'none';
        if (cmpViewStructure) cmpViewStructure.style.display = mode === 'structure' ? 'block' : 'none';
        if (cmpViewRisks) cmpViewRisks.style.display = mode === 'risks' ? 'block' : 'none';
        if (mode === 'risks') renderRisksResult();
    }

    cmpModeTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            switchCompareMode(tab.dataset.cmpMode);
        });
    });

    // ===== Запуск =====
    cmpStartBtn.addEventListener('click', async () => {
        if (state.running) return;
        hideError(cmpError2); hideProgress(cmpProgress);

        if (!state.fileA && !state.fileB) return;

        const isRisksOnly = state.compareMode === 'risks';

        if (!isRisksOnly && (!state.fileA || !state.fileB)) return;

        if (!window.Diff || !window.Diff.diffLines) {
            showError(cmpError2, 'Библиотека сравнения не загружена. Обновите страницу.');
            return;
        }

        const cost = isRisksOnly ? getRiskCost() : getCost();
        if (window.Credits) {
            const canPay = await Credits.canPerform(cost);
            if (!canPay) { Credits.openPurchaseModal(); return; }
        }

        state.running = true;
        cmpStartBtn.disabled = true;
        cmpStartBtn.innerHTML = '<span class="spinner"></span> ' + (isRisksOnly ? 'Анализ...' : 'Сравнение...');
        showProgress(cmpProgress, 'Нормализация текста...');
        await new Promise(r => setTimeout(r, 20));

        try {
            if (isRisksOnly) {
                const srcFile = state.fileA || state.fileB;
                const norm = normalize(srcFile.text);

                showProgress(cmpProgress, 'Анализ юридических рисков...');
                await new Promise(r => setTimeout(r, 20));
                const risks = analyzeRisks(norm);
                state.risksResult = risks.map(r => Object.assign({}, r, { where: null }));

                state.diffParts = null;
                state.structuralResult = null;

                cmpSummary.textContent = 'Анализ рисков в файле «' + srcFile.name + '»';
                renderRisksResult();
                switchCompareMode('risks');

                cmpModeTabs.forEach(t => {
                    const m = t.dataset.cmpMode;
                    t.style.display = (m === 'line' || m === 'structure') ? 'none' : '';
                });
                if (cmpDownloadDocx) cmpDownloadDocx.style.display = 'none';

                cmpStep3.style.display = 'block';
                cmpStep3.classList.add('active');
                hideProgress(cmpProgress);

                if (window.Credits && cost > 0) {
                    await Credits.spend(cost, 'compareRisks');
                }

                setTimeout(() => cmpStep3.scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
                return;
            }

            const normA = normalize(state.fileA.text);
            const normB = normalize(state.fileB.text);

            showProgress(cmpProgress, 'Вычисление различий...');
            await new Promise(r => setTimeout(r, 20));

            const parts = window.Diff.diffLines(normA, normB);
            state.diffParts = parts;
            state.structuralResult = structuralDiff(normA, normB);

            showProgress(cmpProgress, 'Анализ юридических рисков...');
            await new Promise(r => setTimeout(r, 20));
            const risksA = analyzeRisks(normA);
            const risksB = analyzeRisks(normB);
            state.risksResult = mergeRisks(risksA, risksB);

            cmpModeTabs.forEach(t => { t.style.display = ''; });
            if (cmpDownloadDocx) cmpDownloadDocx.style.display = '';

            const paired = renderResult();
            cmpSummary.textContent = buildSummary(paired);
            renderStructuralResult();
            renderRisksResult();
            switchCompareMode(state.mode);

            cmpStep3.style.display = 'block';
            cmpStep3.classList.add('active');
            hideProgress(cmpProgress);

            if (window.Credits && cost > 0) {
                await Credits.spend(cost, 'comparePdf');
            }

            setTimeout(() => cmpStep3.scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
        } catch (err) {
            console.error(err);
            showError(cmpError2, 'Ошибка: ' + err.message);
            hideProgress(cmpProgress);
        } finally {
            state.running = false;
            cmpStartBtn.disabled = false;
            refreshUI();
        }
    });

    // ===== Реактивные переключатели =====
    cmpShowSame.addEventListener('change', () => {
        if (state.diffParts) {
            const paired = renderResult();
            cmpSummary.textContent = buildSummary(paired);
        }
        renderStructuralResult();
    });
    cmpHighlightDanger.addEventListener('change', () => {
        if (state.diffParts) renderResult();
        renderStructuralResult();
    });

    if (cmpRisksCriticalOnly) {
        cmpRisksCriticalOnly.addEventListener('change', () => {
            renderRisksResult();
        });
    }

    // ===== Экспорт отчёта DOCX (изменения) =====
    if (cmpDownloadDocx) {
        cmpDownloadDocx.addEventListener('click', downloadDocxReport);
    }

    async function downloadDocxReport() {
        if (!state.structuralResult) {
            alert('Сначала выполните сравнение.');
            return;
        }
        if (!window.docx) {
            alert('Библиотека DOCX не загружена. Обновите страницу.');
            return;
        }

        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
            AlignmentType, WidthType, BorderStyle } = window.docx;

        const danger = cmpHighlightDanger.checked;
        const changesOnly = state.structuralResult.filter(c => c.type !== 'same');

        function markDanger(text) {
            if (!danger || !text) return text;
            let out = text;
            for (const re of DANGER_PATTERNS) {
                out = out.replace(re, m => '⚑ ' + m + ' ⚑');
            }
            return out;
        }

        const children = [];

        children.push(new Paragraph({
            children: [new TextRun({ text: 'АКТ РАЗНОГЛАСИЙ', bold: true, size: 32 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: 'Отчёт о сравнении двух редакций договора', italics: true, color: '666666' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
        }));

        children.push(new Paragraph({
            children: [
                new TextRun({ text: 'Дата формирования: ', bold: true }),
                new TextRun({ text: new Date().toLocaleDateString('ru-RU') })
            ],
            spacing: { after: 100 }
        }));
        children.push(new Paragraph({
            children: [
                new TextRun({ text: 'Оригинал (A): ', bold: true }),
                new TextRun({ text: state.fileA ? state.fileA.name : '—' })
            ],
            spacing: { after: 100 }
        }));
        children.push(new Paragraph({
            children: [
                new TextRun({ text: 'Редакция (Б): ', bold: true }),
                new TextRun({ text: state.fileB ? state.fileB.name : '—' })
            ],
            spacing: { after: 300 }
        }));

        let added = 0, removed = 0, changed = 0;
        state.structuralResult.forEach(c => {
            if (c.type === 'added') added++;
            else if (c.type === 'removed') removed++;
            else if (c.type === 'changed') changed++;
        });
        children.push(new Paragraph({
            children: [
                new TextRun({ text: 'Итого изменений: ', bold: true }),
                new TextRun({ text: String(changed + added + removed) })
            ]
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: 'Изменено: ' + changed + ', добавлено: ' + added + ', удалено: ' + removed })],
            spacing: { after: 400 }
        }));

        if (changesOnly.length === 0) {
            children.push(new Paragraph({
                children: [new TextRun({ text: 'Различий по пунктам не обнаружено.', italics: true, color: '888888' })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 400, after: 400 }
            }));
        } else {
            children.push(new Paragraph({
                children: [new TextRun({ text: 'Детализация изменений', bold: true, size: 26 })],
                spacing: { after: 200 }
            }));

            const TYPE_LABEL = { changed: 'Изменён', added: 'Добавлен', removed: 'Удалён' };
            const border = { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' };
            const cellBorders = { top: border, bottom: border, left: border, right: border };

            const rows = [
                new TableRow({
                    tableHeader: true,
                    children: [
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Пункт', bold: true })] })], borders: cellBorders, shading: { fill: 'F0F0F0' } }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Тип', bold: true })] })], borders: cellBorders, shading: { fill: 'F0F0F0' } }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Было (A)', bold: true })] })], borders: cellBorders, shading: { fill: 'F0F0F0' } }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Стало (Б)', bold: true })] })], borders: cellBorders, shading: { fill: 'F0F0F0' } })
                    ]
                })
            ];

            for (const c of changesOnly) {
                const numLabel = c.num === '—' ? 'Преамбула' : c.num;
                rows.push(new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph(numLabel)], borders: cellBorders }),
                        new TableCell({ children: [new Paragraph(TYPE_LABEL[c.type] || c.type)], borders: cellBorders }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: c.textA ? markDanger(c.textA) : '—' })] })], borders: cellBorders }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: c.textB ? markDanger(c.textB) : '—' })] })], borders: cellBorders })
                    ]
                }));
            }

            children.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
        }

        children.push(new Paragraph({
            children: [new TextRun({ text: 'Отчёт сформирован автоматически. При необходимости — проверьте вручную.', italics: true, color: '888888', size: 18 })],
            spacing: { before: 600 }
        }));

        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
                children
            }]
        });

        const blob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'akt_raznoglasii_' + new Date().toISOString().slice(0, 10) + '.docx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // ===== Экспорт отчёта по рискам =====
    if (cmpDownloadRisksDocx) {
        cmpDownloadRisksDocx.addEventListener('click', downloadRisksDocx);
    }

    async function downloadRisksDocx() {
        if (!state.risksResult || state.risksResult.length === 0) {
            alert('Нет рисков для экспорта.');
            return;
        }
        if (!window.docx) {
            alert('Библиотека DOCX не загружена.');
            return;
        }

        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
            AlignmentType, WidthType, BorderStyle } = window.docx;

        const children = [];

        children.push(new Paragraph({
            children: [new TextRun({ text: 'ОТЧЁТ О ЮРИДИЧЕСКИХ РИСКАХ', bold: true, size: 32 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: 'Автоматический анализ на типичные «красные флаги»', italics: true, color: '666666' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
        }));

        children.push(new Paragraph({
            children: [
                new TextRun({ text: 'Дата формирования: ', bold: true }),
                new TextRun({ text: new Date().toLocaleDateString('ru-RU') })
            ],
            spacing: { after: 100 }
        }));
        const srcFile = state.fileB || state.fileA;
        children.push(new Paragraph({
            children: [
                new TextRun({ text: 'Файл: ', bold: true }),
                new TextRun({ text: srcFile ? srcFile.name : '—' })
            ],
            spacing: { after: 300 }
        }));

        let high = 0, medium = 0, low = 0;
        state.risksResult.forEach(r => {
            if (r.severity === 'high') high++;
            else if (r.severity === 'medium') medium++;
            else low++;
        });

        children.push(new Paragraph({
            children: [
                new TextRun({ text: 'Всего найдено рисков: ', bold: true }),
                new TextRun({ text: String(state.risksResult.length) })
            ]
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: 'Критичных: ' + high + ', средних: ' + medium + ', низких: ' + low })],
            spacing: { after: 400 }
        }));

        const border = { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' };
        const cellBorders = { top: border, bottom: border, left: border, right: border };
        const SEV_LABEL = { high: 'Критично', medium: 'Средне', low: 'Низко' };
        const WHERE_LABEL = { both: 'В обоих', a: 'Только в A', b: 'Только в Б' };

        const rows = [
            new TableRow({
                tableHeader: true,
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Пункт', bold: true })] })], borders: cellBorders, shading: { fill: 'F0F0F0' } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Категория', bold: true })] })], borders: cellBorders, shading: { fill: 'F0F0F0' } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Уровень', bold: true })] })], borders: cellBorders, shading: { fill: 'F0F0F0' } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Фрагмент', bold: true })] })], borders: cellBorders, shading: { fill: 'F0F0F0' } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Рекомендация', bold: true })] })], borders: cellBorders, shading: { fill: 'F0F0F0' } })
                ]
            })
        ];

        for (const r of state.risksResult) {
            const numLabel = r.clauseNum ? 'п. ' + r.clauseNum : 'Преамбула';
            const whereText = (r.where && WHERE_LABEL[r.where]) ? ' (' + WHERE_LABEL[r.where] + ')' : '';
            rows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph(numLabel)], borders: cellBorders }),
                    new TableCell({ children: [new Paragraph(r.category + whereText)], borders: cellBorders }),
                    new TableCell({ children: [new Paragraph(SEV_LABEL[r.severity])], borders: cellBorders }),
                    new TableCell({ children: [new Paragraph(r.quote)], borders: cellBorders }),
                    new TableCell({ children: [new Paragraph(r.hint)], borders: cellBorders })
                ]
            }));
        }

        children.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));

        children.push(new Paragraph({
            children: [new TextRun({
                text: '⚠️ Это автоматический анализ по формальным признакам. Он не заменяет профессиональную юридическую оценку. Обязательно проверьте договор у юриста перед подписанием.',
                italics: true, color: '888888', size: 18
            })],
            spacing: { before: 600 }
        }));

        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
                children
            }]
        });

        const blob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'riski_' + new Date().toISOString().slice(0, 10) + '.docx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // ===== Сброс =====
    cmpResetBtn.addEventListener('click', () => {
        state.fileA = null;
        state.fileB = null;
        state.diffParts = null;
        state.structuralResult = null;
        state.risksResult = null;
        state.compareMode = 'compare';
        cmpInputA.value = '';
        cmpInputB.value = '';
        cmpDropA.classList.remove('has-file');
        cmpDropB.classList.remove('has-file');
        cmpDropA.querySelector('.drop-zone-text').innerHTML = 'Перетащите файл сюда или <strong>выберите</strong>';
        cmpDropB.querySelector('.drop-zone-text').innerHTML = 'Перетащите файл сюда или <strong>выберите</strong>';
        cmpInfoA.textContent = '';
        cmpInfoB.textContent = '';
        cmpResult.innerHTML = '';
        if (cmpStructList) cmpStructList.innerHTML = '';
        if (cmpStructSummary) cmpStructSummary.textContent = '—';
        if (cmpRisksList) cmpRisksList.innerHTML = '';
        if (cmpRisksSummary) cmpRisksSummary.textContent = '—';
        if (cmpRisksCriticalOnly) cmpRisksCriticalOnly.checked = false;
        cmpSummary.textContent = '—';

        cmpModeTabs.forEach(t => { t.style.display = ''; });
        if (cmpDownloadDocx) cmpDownloadDocx.style.display = '';

        cmpStep2.classList.remove('active');
        cmpStep3.style.display = 'none';
        cmpStep3.classList.remove('active', 'done');
        hideError(cmpError); hideError(cmpError2); hideProgress(cmpProgress);
        refreshUI();
    });

    // Добавляем стоимость в pricing, если её нет
    if (window.PRICING && window.PRICING.costs) {
        if (!window.PRICING.costs.comparePdf) window.PRICING.costs.comparePdf = 3;
        if (!window.PRICING.costs.compareRisks) window.PRICING.costs.compareRisks = 1;
    }
    if (window.PRICING && window.PRICING.limits) {
        if (!window.PRICING.limits.maxCompareFileSizeMB) window.PRICING.limits.maxCompareFileSizeMB = 20;
        if (!window.PRICING.limits.maxComparePages) window.PRICING.limits.maxComparePages = 100;
    }

})();