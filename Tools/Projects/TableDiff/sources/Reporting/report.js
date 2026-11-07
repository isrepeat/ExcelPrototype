const searchInput = document.querySelector('.report-toolbar input[type="search"]');
const changeFilter = document.querySelector('.report-toolbar select');
const toolbar = document.querySelector('.report-toolbar');
const toolbarControls = document.getElementById('toolbar-controls');
const showUnchangedFormulas = document.getElementById('show-unchanged-formulas');
const pageSections = [...document.querySelectorAll('main > section')];
const hideUnchangedColumns = document.getElementById('hide-unchanged-columns');
const pageSize = 100;
let liveColumnResize = document.getElementById('live-column-resize').checked;
const kinds = ['change', 'add', 'delete'];
const states = new Map();
const widthStoreKey = 'excel-diff-column-widths-v1';
let savedWidths = Object.create(null);
let activeState = null;
let searchTimer = null;
let filterVersion = 0;
let boundaryUpdatePending = false;
let cancelActiveResize = null;
let columnScale = 1;
const columnScaleInput = document.getElementById('column-scale');
const columnScaleValue = document.getElementById('column-scale-value');

columnScaleInput.addEventListener('input', () => {
    cancelActiveResize?.();
    columnScale = Number(columnScaleInput.value) / 100;
    columnScaleValue.value = columnScaleInput.value + '%';
    const section = activeState?.section;
    if (!section) {
        return;
    }
    const columns = dataColumns(section);
    const headers = [...section.querySelectorAll('th[data-column-name]')];
    const handles = [...section.querySelectorAll('.column-resizer')];
    let total = 134 + section.querySelectorAll('col[data-collapsed]').length * 24;
    columns.forEach((column, index) => {
        const width = storedWidth(section, headers[index].dataset.columnName) * columnScale;
        column.style.width = width + 'px';
        total += width;
        handles[index].setAttribute('aria-valuemin', String(60 * columnScale));
        handles[index].setAttribute('aria-valuemax', String(600 * columnScale));
        handles[index].setAttribute('aria-valuenow', String(width));
    });
    section.querySelector('.diff-table').style.width = total + 'px';
    scheduleBoundaryUpdate();
});

let navigationFrame = null;

function cancelNavigation() {
    cancelAnimationFrame(navigationFrame);
    navigationFrame = null;
}

window.addEventListener('wheel', cancelNavigation, { passive: true });
window.addEventListener('touchstart', cancelNavigation, { passive: true });
window.addEventListener('pointerdown', cancelNavigation, { passive: true });
window.addEventListener('keydown', cancelNavigation);

function scrollToTableEdge(bottom) {
    cancelNavigation();
    const table = activeState?.section.querySelector('.diff-table');
    if (!table) {
        return;
    }
    const bounds = table.getBoundingClientRect();
    const top = bottom
        ? window.scrollY + bounds.bottom - window.innerHeight + pageFooter.offsetHeight + 8
        : window.scrollY + bounds.top - toolbar.offsetHeight;
    const target = Math.max(0, Math.min(top, document.documentElement.scrollHeight - window.innerHeight));
    const left = window.scrollX;
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
        window.scrollTo({ left, top: target, behavior: 'instant' });
        return;
    }
    const start = window.scrollY;
    const started = performance.now();
    const duration = 250;
    function animate(now) {
        const progress = Math.min(1, (now - started) / duration);
        const eased = 1 - Math.pow(1 - progress, 3);
        window.scrollTo({ left, top: start + (target - start) * eased, behavior: 'instant' });
        navigationFrame = progress < 1 ? requestAnimationFrame(animate) : null;
    }
    navigationFrame = requestAnimationFrame(animate);
}

document.getElementById('table-top').addEventListener('click', () => scrollToTableEdge(false));
document.getElementById('table-bottom').addEventListener('click', () => scrollToTableEdge(true));


try {
    const stored = JSON.parse(localStorage.getItem(widthStoreKey) || '{}');
    if (stored && typeof stored === 'object' && !Array.isArray(stored)) {
        savedWidths = Object.assign(Object.create(null), stored);
    }
}
catch (error) {
    savedWidths = Object.create(null);
}

const pageFooter = document.createElement('footer');
pageFooter.className = 'page-footer';
pageFooter.setAttribute('aria-label', 'Листы отчёта');
const footerTabs = document.createElement('div');
footerTabs.className = 'footer-tabs';
pageFooter.append(footerTabs);
document.body.append(pageFooter);

const resetWidths = document.createElement('button');
resetWidths.type = 'button';
resetWidths.textContent = 'Сбросить ширины';
resetWidths.hidden = true;
toolbarControls.append(resetWidths);

const columnFilter = document.createElement('details');
const columnFilterSummary = document.createElement('summary');
columnFilterSummary.textContent = 'Изменённые поля';
const columnFilterOptions = document.createElement('div');
const exportChanges = document.createElement('button');
exportChanges.type = 'button';
exportChanges.textContent = 'Скачать XML';
exportChanges.disabled = true;
columnFilter.append(columnFilterSummary, columnFilterOptions, exportChanges);
toolbarControls.append(columnFilter);

exportChanges.addEventListener('click', () => {
    const state = activeState;
    const xml = document.implementation.createDocument(null, 'tableDiff');
    const root = xml.documentElement;
    root.setAttribute('version', '1');
    root.setAttribute('sheet', sheetName(state.section));

    for (const rowIndex of state.filtered) {
        const row = state.data.rows[rowIndex];
        const record = xml.createElement('row');
        record.setAttribute('oldRow', String(row[1]));
        record.setAttribute('newRow', String(row[2]));
        record.setAttribute('ambiguous', String(Boolean(row[3])));
        for (const index of state.selectedColumns) {
            const before = row[4]?.[index] ?? 0;
            const after = row[5]?.[index] ?? 0;
            if (before === after) {
                continue;
            }
            const field = xml.createElement('field');
            field.setAttribute('name', state.data.columns[index][1]);
            const oldValue = xml.createElement('old');
            oldValue.textContent = state.data.strings[before];
            const newValue = xml.createElement('new');
            newValue.textContent = state.data.strings[after];
            field.append(oldValue, newValue);
            record.append(field);
        }
        root.append(record);
    }

    const content = new XMLSerializer().serializeToString(xml).replaceAll('\r', '&#13;');
    const parsed = new DOMParser().parseFromString(content, 'application/xml');
    if (parsed.documentElement.localName !== 'tableDiff') {
        window.alert('В срезе есть символы, недопустимые в XML.');
        return;
    }
    const blob = new Blob(['<?xml version="1.0" encoding="UTF-8"?>\n', content], { type: 'application/xml' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'table-diff-slice.xml';
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
});

const settingsActions = document.createElement('div');
settingsActions.className = 'settings-actions';
const exportWidths = document.createElement('button');
exportWidths.type = 'button';
exportWidths.textContent = 'Экспорт JSON';
const importWidths = document.createElement('button');
importWidths.type = 'button';
importWidths.textContent = 'Импорт JSON';
const settingsMessage = document.createElement('span');
settingsMessage.setAttribute('role', 'status');
const settingsFile = document.createElement('input');
settingsFile.type = 'file';
settingsFile.accept = '.json,application/json';
settingsFile.hidden = true;
settingsActions.append(exportWidths, importWidths, settingsMessage, settingsFile);
document.querySelector('.report-summary').append(settingsActions);

function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, character => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
    })[character]);
}

function getState(section) {
    if (!states.has(section.id)) {
        const source = document.getElementById('data-' + section.id);
        const data = JSON.parse(source.textContent);
        source.remove();
        states.set(section.id, {
            section,
            data,
            page: 0,
            filtered: null,
            query: null,
            kind: null,
            selectedColumns: new Set(),
            columnFilterKey: null,
            searchCache: new Array(data.rows.length),
            lowerStrings: null
        });
    }
    return states.get(section.id);
}

function sheetName(section) {
    return section.dataset.sheetName;
}

function columnNames(section) {
    return getState(section).data.columns.map(column => column[1]);
}

function dataColumns(section) {
    return [...section.querySelectorAll('.diff-table col[data-column-index]')];
}

function storedWidth(section, name) {
    const value = Number(savedWidths[sheetName(section)]?.[name]);
    return Number.isFinite(value) && value >= 60 && value <= 600 ? value : 170;
}

function saveWidths() {
    try {
        localStorage.setItem(widthStoreKey, JSON.stringify(savedWidths));
        settingsMessage.textContent = '';
    }
    catch (error) {
        settingsMessage.textContent = 'Браузер запретил сохранение размеров';
    }
}

function updateResetButton() {
    const section = activeState?.section;
    const values = section ? savedWidths[sheetName(section)] || {} : {};
    resetWidths.hidden = !section || !columnNames(section).some(name => Object.hasOwn(values, name));
}

function rememberWidth(section, name, width) {
    width /= columnScale;
    const sheet = sheetName(section);
    if (width === 170 && !Object.hasOwn(savedWidths, sheet)) {
        return;
    }
    if (!Object.hasOwn(savedWidths, sheet)) {
        savedWidths[sheet] = Object.create(null);
    }
    savedWidths[sheet][name] = width;
    saveWidths();
    updateResetButton();
}

function applyWidth(section, index, width) {
    const columns = dataColumns(section);
    if (!columns[index]) {
        return;
    }
    columns[index].style.width = width + 'px';
    section.querySelector('.diff-table').style.width =
        (134 + columns.reduce((total, column) => total + parseFloat(column.style.width), 0)
            + section.querySelectorAll('col[data-collapsed]').length * 24) + 'px';
    const handle = section.querySelectorAll('.column-resizer')[index];
    handle?.setAttribute('aria-valuenow', String(width));
    scheduleBoundaryUpdate();
}

function columnLetter(number) {
    let result = '';
    while (number > 0) {
        number--;
        result = String.fromCharCode(65 + number % 26) + result;
        number = Math.floor(number / 26);
    }
    return result;
}

function cellValue(value) {
    return '<span class="cell-value">' + escapeHtml(value) + '</span>';
}

function renderRow(state, row) {
    const kind = kinds[row[0]];
    const beforeNumber = row[1] || '—';
    const afterNumber = row[2] || '—';
    const location = beforeNumber === afterNumber ? afterNumber : beforeNumber + ' → ' + afterNumber;
    const status = ['~', '+', '−'][row[0]];
    const label = ['Изменена строка', 'Добавлена строка', 'Удалена строка'][row[0]];
    const warning = escapeHtml(row[6] ?? 'Несколько записей имеют одинаковую близость: проверьте сопоставление.');
    let html = '<tr class="diff-row ' + kind + '" data-kind="' + kind + '"><td class="status-column" aria-label="' + label + '">' + status;

    if (row[3]) {
        html += '<span class="ambiguity" title="' + warning + '" aria-label="' + warning + '">?</span>';
    }
    html += '</td><td class="row-column">' + location + '</td>';

    state.displayColumns.forEach(entry => {
        if (entry.collapsed) {
            html += '<td class="collapsed-columns" title="' + escapeHtml(entry.title) + '">…</td>';
            return;
        }
        const index = entry.index;
        const column = state.data.columns[index];
        const x = state.data.strings[row[4]?.[index] ?? 0];
        const y = state.data.strings[row[5]?.[index] ?? 0];
        const unchanged = (kind === 'change' && x === y) || (!x && !y);
        const status = unchanged ? 'unchanged' : x && y ? 'replacement' : y ? 'insertion' : 'removal';
        const oldAddress = row[1] && column[2] >= 0 ? columnLetter(column[2] + 1) + row[1] : '—';
        const newAddress = row[2] && column[3] >= 0 ? columnLetter(column[3] + 1) + row[2] : '—';
        html += '<td class="' + status + '" title="Было: ' + oldAddress + ' · Стало: ' + newAddress + '">';

        if (unchanged) {
            if (x.startsWith('=')) {
                html += '<span class="formula-marker">= без изменений</span><span class="unchanged-formula">' + cellValue(x) + '</span>';
            }
            else {
                html += x ? cellValue(x) : '<span class="empty-cell">—</span>';
            }
        }
        else {
            if (x) {
                html += '<del aria-label="Было">' + cellValue(x) + '</del>';
            }
            if (y) {
                html += '<ins aria-label="Стало">' + cellValue(y) + '</ins>';
            }
        }
        html += '</td>';
    });
    return html + '</tr>';
}

function displayColumns(state, rowIndices) {
    const columns = state.data.columns;
    if (!hideUnchangedColumns.checked || rowIndices.length === 0) {
        return columns.map((column, index) => ({ index }));
    }

    const changed = columns.map((column, index) =>
        column[2] < 0 || column[3] < 0 || rowIndices.some(rowIndex => {
            const row = state.data.rows[rowIndex];
            return (row[4]?.[index] ?? 0) !== (row[5]?.[index] ?? 0);
        })
    );
    const result = [];
    for (let index = 0; index < columns.length;) {
        if (changed[index]) {
            result.push({ index });
            index++;
            continue;
        }

        const start = index;
        while (index < columns.length && !changed[index]) {
            index++;
        }
        const count = index - start;
        result.push({
            collapsed: true,
            start,
            count,
            title: 'Неизменённые колонки (' + count + '): ' + columns.slice(start, index).map(column => column[1]).join(' · ')
        });
    }
    return result;
}

function mountTable(state, layout = null) {
    cancelActiveResize?.();
    boundaryObserver.disconnect();
    const section = state.section;
    state.displayColumns = layout || state.data.columns.map((column, index) => ({ index }));
    state.layoutKey = state.displayColumns.map(entry => entry.collapsed ? 'g' + entry.start + ':' + entry.count : String(entry.index)).join(',');
    const widths = state.displayColumns.map(entry => entry.collapsed ? 24 : storedWidth(section, state.data.columns[entry.index][1]) * columnScale);
    let html = '<table class="diff-table" aria-label="' + escapeHtml(sheetName(section)) + '" style="width:'
        + (134 + widths.reduce((sum, width) => sum + width, 0)) + 'px"><colgroup><col style="width:44px"><col style="width:90px">';

    state.displayColumns.forEach((entry, index) => {
        html += entry.collapsed
            ? '<col data-collapsed="true" style="width:24px">'
            : '<col data-column-index="' + state.data.columns[entry.index][0] + '" style="width:' + widths[index] + 'px">';
    });
    html += '</colgroup><thead><tr><th class="status-column" scope="col">Δ</th><th class="row-column" scope="col">Строка</th>';

    for (const entry of state.displayColumns) {
        if (entry.collapsed) {
            html += '<th class="collapsed-columns" scope="col" title="' + escapeHtml(entry.title)
                + '" aria-label="' + escapeHtml(entry.title) + '">…</th>';
        }
        else {
            const column = state.data.columns[entry.index];
            const status = column[2] < 0 ? '+ ' : column[3] < 0 ? '− ' : '';
            const title = column[2] < 0 ? 'Добавлена колонка' : column[3] < 0 ? 'Удалена колонка' : '';
            html += '<th scope="col" title="' + title + '" data-column-name="' + escapeHtml(column[1]) + '">'
                + status + escapeHtml(column[1]) + '</th>';
        }
    }
    html += '</tr></thead><tbody></tbody></table>';
    section.querySelector('.table-host').innerHTML = html;
    initializeResizers(section);
    boundaryObserver.observe(section.querySelector('.diff-table'));
}

function renderPage(state, scroll = false) {
    if (state !== activeState || state.filtered === null) {
        return;
    }
    const count = state.filtered.length;
    const pages = Math.max(1, Math.ceil(count / pageSize));
    state.page = Math.min(Math.max(0, state.page), pages - 1);
    const start = state.page * pageSize;
    const visible = state.filtered.slice(start, start + pageSize);
    const layout = displayColumns(state, visible);
    const layoutKey = layout.map(entry => entry.collapsed ? 'g' + entry.start + ':' + entry.count : String(entry.index)).join(',');
    if (layoutKey !== state.layoutKey) {
        mountTable(state, layout);
    }
    state.section.querySelector('tbody').innerHTML = visible.map(index => renderRow(state, state.data.rows[index])).join('');
    state.section.querySelector('.page-status').textContent = count
        ? (start + 1) + '–' + (start + visible.length) + ' из ' + count + ' · страница ' + (state.page + 1) + '/' + pages
        : '0 из ' + state.data.rows.length;
    const message = state.section.querySelector('.empty-message');
    message.textContent = state.data.rows.length ? 'Нет строк, соответствующих фильтру.' : 'Изменений данных и формул нет.';
    message.hidden = count > 0;

    for (const button of state.section.querySelectorAll('[data-page-action]')) {
        const backwards = ['first', 'previous'].includes(button.dataset.pageAction);
        button.disabled = backwards ? state.page === 0 : state.page === pages - 1;
    }
    scheduleBoundaryUpdate();
    if (scroll) {
        window.scrollTo(window.scrollX, 0);
    }
}

function rowSearchText(state, index) {
    if (state.searchCache[index] !== undefined) {
        return state.searchCache[index];
    }
    if (!state.lowerStrings) {
        state.lowerStrings = state.data.strings.map(value => value.toLowerCase());
    }
    const row = state.data.rows[index];
    const strings = new Set([...(row[4] || []), ...(row[5] || [])]);
    const parts = [...strings].map(id => state.lowerStrings[id]);
    parts.push(String(row[1] || ''), String(row[2] || ''));

    state.data.columns.forEach((column, position) => {
        if (row[0] !== 0 || (row[4]?.[position] ?? 0) !== (row[5]?.[position] ?? 0)) {
            parts.push(column[1].toLowerCase());
        }
    });
    state.searchCache[index] = parts.join('\u0000');
    return state.searchCache[index];
}

async function filterCurrent() {
    clearTimeout(searchTimer);
    const state = activeState;
    if (!state) {
        return;
    }
    const request = ++filterVersion;
    const query = searchInput.value.toLowerCase();
    const kind = changeFilter.value;
    const selectedColumns = [...state.selectedColumns];
    const columnFilterKey = selectedColumns.join(',');
    if (state.filtered !== null && state.query === query && state.kind === kind
        && state.columnFilterKey === columnFilterKey) {
        exportChanges.disabled = selectedColumns.length === 0;
        renderPage(state);
        return;
    }

    exportChanges.disabled = true;
    state.section.querySelector('.page-status').textContent = 'Поиск…';
    const matched = [];
    for (let index = 0; index < state.data.rows.length; index++) {
        const row = state.data.rows[index];
        const matchesColumns = selectedColumns.length === 0 || (row[0] === 0 && selectedColumns.some(column =>
            (row[4]?.[column] ?? 0) !== (row[5]?.[column] ?? 0)
        ));
        if (matchesColumns && (!kind || kinds[row[0]] === kind) && (!query || rowSearchText(state, index).includes(query))) {
            matched.push(index);
        }
        if (query && index % 500 === 499) {
            await new Promise(resolve => setTimeout(resolve, 0));
            if (state !== activeState || request !== filterVersion) {
                return;
            }
        }
    }
    if (state !== activeState || request !== filterVersion) {
        return;
    }
    state.filtered = matched;
    state.query = query;
    state.kind = kind;
    state.columnFilterKey = columnFilterKey;
    exportChanges.disabled = selectedColumns.length === 0;
    state.page = 0;
    renderPage(state);
}

function selectPage(section, updateHistory = true) {
    document.querySelector('.source-files').textContent = section.dataset.sourceFiles;
    document.title = section.dataset.pageTitle;
    cancelActiveResize?.();
    filterVersion++;
    boundaryObserver.disconnect();
    if (activeState) {
        activeState.section.querySelector('.table-host').replaceChildren();
    }
    for (const item of pageSections) {
        item.hidden = item !== section;
    }
    activeState = getState(section);
    columnFilterOptions.replaceChildren();
    columnFilter.open = false;
    const state = activeState;
    state.data.columns.forEach((column, index) => {
        const option = document.createElement('div');
        const label = document.createElement('label');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = state.selectedColumns.has(index);
        checkbox.addEventListener('change', () => {
            if (checkbox.checked) {
                state.selectedColumns.add(index);
            }
            else {
                state.selectedColumns.delete(index);
            }
            columnFilterSummary.textContent = 'Изменённые поля (' + state.selectedColumns.size + ')';
            filterCurrent();
        });
        label.append(checkbox, ' ' + column[1]);
        option.append(label);
        columnFilterOptions.append(option);
    });
    columnFilterSummary.textContent = 'Изменённые поля (' + state.selectedColumns.size + ')';
    mountTable(activeState);

    for (const button of footerTabs.querySelectorAll('button')) {
        const active = button.dataset.sectionId === section.id;
        button.classList.toggle('active', active);
        button.setAttribute('aria-pressed', String(active));
    }
    if (updateHistory) {
        history.replaceState(null, '', '#' + section.id);
    }
    window.scrollTo(0, 0);
    updateResetButton();
    filterCurrent();
}

for (const section of pageSections) {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = sheetName(section);
    button.dataset.sectionId = section.id;
    button.onclick = () => selectPage(section);
    footerTabs.append(button);

    section.querySelector('.pagination').addEventListener('click', event => {
        const button = event.target.closest('[data-page-action]');
        if (!button || !activeState || section !== activeState.section || activeState.filtered === null) {
            return;
        }
        const action = button.dataset.pageAction;
        const last = Math.max(0, Math.ceil(activeState.filtered.length / pageSize) - 1);
        activeState.page = action === 'first' ? 0 : action === 'last' ? last
            : activeState.page + (action === 'previous' ? -1 : 1);
        renderPage(activeState, true);
    });
}

searchInput.addEventListener('input', () => {
    exportChanges.disabled = true;
    filterVersion++;
    clearTimeout(searchTimer);
    searchTimer = setTimeout(filterCurrent, 200);
});
changeFilter.addEventListener('change', filterCurrent);
hideUnchangedColumns.addEventListener('change', () => {
    cancelActiveResize?.();
    if (activeState) {
        renderPage(activeState);
    }
});

showUnchangedFormulas.addEventListener('change', () => {
    document.body.classList.toggle('show-unchanged-formulas', showUnchangedFormulas.checked);
});
document.getElementById('live-column-resize').addEventListener('change', event => {
    cancelActiveResize?.();
    liveColumnResize = event.target.checked;
});

function updateResizeBoundaries() {
    if (!activeState) {
        return;
    }
    const table = activeState.section.querySelector('.diff-table');
    if (!table) {
        return;
    }
    const bottom = table.getBoundingClientRect().bottom;
    const headers = [...table.querySelectorAll('th[data-column-name]')];
    const geometry = headers.map(header => {
        const bounds = header.getBoundingClientRect();
        return {
            height: Math.max(bounds.height, bottom - bounds.top),
            hidden: bounds.right <= 134
        };
    });
    headers.forEach((header, index) => {
        const handle = header.querySelector('.column-resizer');
        handle.style.height = geometry[index].height + 'px';
        handle.hidden = geometry[index].hidden;
    });
}

function scheduleBoundaryUpdate() {
    if (boundaryUpdatePending) {
        return;
    }
    boundaryUpdatePending = true;
    requestAnimationFrame(() => {
        boundaryUpdatePending = false;
        updateResizeBoundaries();
    });
}

const boundaryObserver = new ResizeObserver(scheduleBoundaryUpdate);
window.addEventListener('scroll', scheduleBoundaryUpdate, { passive: true });
window.addEventListener('resize', scheduleBoundaryUpdate);

function initializeResizers(section) {
    section.querySelectorAll('th[data-column-name]').forEach((header, index) => {
        const handle = document.createElement('span');
        handle.className = 'column-resizer';
        handle.tabIndex = 0;
        handle.setAttribute('role', 'separator');
        handle.setAttribute('aria-orientation', 'vertical');
        handle.setAttribute('aria-label', 'Ширина колонки ' + header.dataset.columnName);
        handle.setAttribute('aria-valuemin', String(60 * columnScale));
        handle.setAttribute('aria-valuemax', String(600 * columnScale));
        handle.setAttribute('aria-valuenow', String(parseFloat(dataColumns(section)[index].style.width)));
        handle.title = 'Перетащите границу. Двойной щелчок — сброс ширины.';
        header.append(handle);
        let drag = null;

        function setWidth(value) {
            const width = Math.min(600 * columnScale, Math.max(60 * columnScale, Math.round(value)));
            applyWidth(section, index, width);
            return width;
        }

        function finish(cancelled) {
            if (!drag) {
                return;
            }
            const active = drag;
            drag = null;
            cancelActiveResize = null;
            active.guide?.remove();
            document.body.classList.remove('column-resizing');
            if (cancelled && liveColumnResize) {
                setWidth(active.width);
            }
            else if (!cancelled) {
                rememberWidth(section, header.dataset.columnName, setWidth(active.proposed));
            }
            if (handle.hasPointerCapture(active.pointerId)) {
                handle.releasePointerCapture(active.pointerId);
            }
        }

        handle.addEventListener('pointerdown', event => {
            if (event.button !== 0 || drag) {
                return;
            }
            event.preventDefault();
            handle.focus({ preventScroll: true });
            const table = section.querySelector('.diff-table').getBoundingClientRect();
            const bounds = header.getBoundingClientRect();
            let guide = null;
            if (!liveColumnResize) {
                guide = document.createElement('div');
                guide.className = 'resize-guide';
                guide.style.left = (bounds.right - 0.5) + 'px';
                guide.style.top = bounds.top + 'px';
                guide.style.height = Math.max(bounds.height, Math.min(table.bottom, window.innerHeight - pageFooter.offsetHeight) - bounds.top) + 'px';
                document.body.append(guide);
            }
            const width = parseFloat(dataColumns(section)[index].style.width);
            drag = { pointerId: event.pointerId, x: event.clientX, scrollX: window.scrollX, width, proposed: width, guide };
            cancelActiveResize = () => finish(true);
            handle.setPointerCapture(event.pointerId);
            document.body.classList.add('column-resizing');
        });

        handle.addEventListener('pointermove', event => {
            if (drag && event.pointerId === drag.pointerId) {
                drag.proposed = Math.min(600 * columnScale, Math.max(60 * columnScale,
                    Math.round(drag.width + event.clientX - drag.x + window.scrollX - drag.scrollX)));
                if (liveColumnResize) {
                    setWidth(drag.proposed);
                }
                else {
                    drag.guide.style.transform = 'translateX(' + (drag.proposed - drag.width - window.scrollX + drag.scrollX) + 'px)';
                }
            }
        });
        handle.addEventListener('pointerup', event => {
            if (drag && event.pointerId === drag.pointerId) {
                finish(false);
            }
        });
        handle.addEventListener('pointercancel', () => finish(true));
        handle.addEventListener('lostpointercapture', () => finish(true));
        handle.addEventListener('dblclick', event => {
            event.preventDefault();
            rememberWidth(section, header.dataset.columnName, setWidth(170 * columnScale));
        });
        handle.addEventListener('keydown', event => {
            const current = parseFloat(dataColumns(section)[index].style.width);
            const step = event.shiftKey ? 50 : 10;
            let target;
            if (event.key === 'ArrowLeft') {
                target = current - step;
            }
            else if (event.key === 'ArrowRight') {
                target = current + step;
            }
            else if (event.key === 'Home') {
                target = 60 * columnScale;
            }
            else if (event.key === 'End') {
                target = 600 * columnScale;
            }
            else if (event.key === 'Escape' && drag) {
                finish(true);
            }
            else {
                return;
            }
            event.preventDefault();
            if (target !== undefined) {
                rememberWidth(section, header.dataset.columnName, setWidth(target));
            }
        });
    });
}

resetWidths.onclick = () => {
    if (!activeState) {
        return;
    }
    cancelActiveResize?.();
    const section = activeState.section;
    delete savedWidths[sheetName(section)];
    saveWidths();
    boundaryObserver.disconnect();
    mountTable(activeState);
    renderPage(activeState);
    updateResetButton();
};

function updateLayoutSpace() {
    document.body.style.paddingBottom = (pageFooter.offsetHeight + 8) + 'px';
    document.documentElement.style.setProperty('--toolbar-height', toolbar.offsetHeight + 'px');
    scheduleBoundaryUpdate();
}
const layoutObserver = new ResizeObserver(updateLayoutSpace);
layoutObserver.observe(pageFooter);
layoutObserver.observe(toolbar);
updateLayoutSpace();

exportWidths.onclick = () => {
    const allWidths = Object.create(null);
    for (const section of pageSections) {
        const columns = Object.create(null);
        for (const name of columnNames(section)) {
            columns[name] = storedWidth(section, name);
        }
        allWidths[sheetName(section)] = columns;
    }
    const payload = { version: 1, columnWidths: allWidths };
    const url = URL.createObjectURL(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = 'excel-column-widths.json';
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    settingsMessage.textContent = 'Настройки экспортированы';
};
importWidths.onclick = () => settingsFile.click();
settingsFile.onchange = async () => {
    const file = settingsFile.files[0];
    if (!file) {
        return;
    }
    try {
        if (file.size > 1048576) {
            throw new Error('Файл настроек слишком большой');
        }
        const payload = JSON.parse(await file.text());
        if (!payload || payload.version !== 1 || !payload.columnWidths
            || typeof payload.columnWidths !== 'object' || Array.isArray(payload.columnWidths)) {
            throw new Error('Неверный формат настроек');
        }
        const validated = Object.create(null);
        for (const [sheet, columns] of Object.entries(payload.columnWidths)) {
            if (!columns || typeof columns !== 'object' || Array.isArray(columns)) {
                throw new Error('Неверные настройки листа');
            }
            validated[sheet] = Object.create(null);
            for (const [name, width] of Object.entries(columns)) {
                if (typeof width !== 'number' || !Number.isFinite(width) || width < 60 || width > 600) {
                    throw new Error('Ширина должна быть числом от 60 до 600');
                }
                validated[sheet][name] = width;
            }
        }
        cancelActiveResize?.();
        savedWidths = validated;
        saveWidths();
        if (activeState) {
            boundaryObserver.disconnect();
            mountTable(activeState);
            renderPage(activeState);
        }
        updateResetButton();
        settingsMessage.textContent = 'Настройки импортированы';
    }
    catch (error) {
        settingsMessage.textContent = 'Ошибка импорта: ' + error.message;
    }
    finally {
        settingsFile.value = '';
    }
};

const initial = pageSections.find(section => '#' + section.id === location.hash) || pageSections[0];
if (initial) {
    selectPage(initial, false);
}
window.addEventListener('load', () => window.scrollTo(0, 0), { once: true });