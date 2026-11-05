const searchInput = document.querySelector('input');
const changeFilter = document.querySelector('select');

function filter() {
    for (const article of document.querySelectorAll('article')) {
        const matchesText = article.textContent.toLowerCase().includes(searchInput.value.toLowerCase());
        const matchesKind = !changeFilter.value || article.dataset.kind === changeFilter.value;
        article.hidden = !(matchesText && matchesKind);
    }
}

searchInput.oninput = filter;
changeFilter.onchange = filter;

const dialog = document.createElement('dialog');
dialog.className = 'cell-dialog';
dialog.innerHTML =
    '<div class="dialog-top"><h3 id="cell-dialog-title"></h3>' +
    '<button type="button" class="dialog-close" aria-label="Закрыть">×</button></div>' +
    '<p class="dialog-context"></p><div class="dialog-body"></div>';
dialog.setAttribute('aria-labelledby', 'cell-dialog-title');
document.body.append(dialog);

dialog.querySelector('.dialog-close').onclick = () => dialog.close();
dialog.addEventListener('click', event => {
    const bounds = dialog.getBoundingClientRect();
    const outsideBounds =
        event.clientX < bounds.left ||
        event.clientX > bounds.right ||
        event.clientY < bounds.top ||
        event.clientY > bounds.bottom;

    if (event.target === dialog && outsideBounds) {
        dialog.close();
    }
});

function openCell(cell) {
    const article = cell.closest('article');
    dialog.querySelector('h3').textContent = cell.querySelector('.cell-heading b').textContent;
    dialog.querySelector('.dialog-context').textContent =
        article.querySelector('header').textContent +
        ' · ' +
        article.closest('section').querySelector('h2').textContent;

    const body = dialog.querySelector('.dialog-body');
    body.replaceChildren();

    for (const child of cell.children) {
        if (child.tagName !== 'TEMPLATE') {
            body.append(child.cloneNode(true));
        }
    }

    const original = cell.querySelector('template.original-value');
    if (original) {
        body.querySelector('.current').replaceChildren(original.content.cloneNode(true));
    }

    for (const detail of body.querySelectorAll('details')) {
        detail.open = true;
    }

    dialog.showModal();
}

for (const cell of document.querySelectorAll('.cell')) {
    cell.setAttribute('role', 'button');
    cell.tabIndex = 0;
    cell.setAttribute('aria-haspopup', 'dialog');
    cell.setAttribute(
        'aria-label',
        cell.querySelector('.cell-heading b').textContent + ' — открыть подробности'
    );

    cell.addEventListener('click', event => {
        event.preventDefault();
        openCell(cell);
    });

    cell.addEventListener('keydown', event => {
        if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            openCell(cell);
        }
    });
}

const pageFooter = document.createElement('footer');
pageFooter.className = 'page-footer';
pageFooter.setAttribute('aria-label', 'Листы отчёта');

const pageSections = [...document.querySelectorAll('main>section')];
const pageNames = pageSections.map(section => {
    const name = section.querySelector('h2').firstChild.textContent.trim();
    return [name, name];
});

function selectPage(section, button) {
    for (const item of pageSections) {
        item.hidden = item !== section;
    }

    for (const tab of pageFooter.querySelectorAll('button')) {
        const active = tab === button;
        tab.classList.toggle('active', active);
        tab.setAttribute('aria-pressed', String(active));
    }

    history.replaceState(null, '', '#' + section.id);
    window.scrollTo(0, 0);
}

for (const [sheet, label] of pageNames) {
    const section = pageSections.find(
        section => section.querySelector('h2').firstChild.textContent.trim() === sheet
    );

    if (!section) {
        continue;
    }

    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = label;
    button.onclick = () => selectPage(section, button);
    pageFooter.append(button);
}

document.body.append(pageFooter);

const initial =
    pageSections.find(section =>
        '#' + section.id === location.hash &&
        pageNames.some(([name]) => section.querySelector('h2').firstChild.textContent.trim() === name)
    ) ||
    pageSections.find(section => section.querySelector('h2').firstChild.textContent.trim() === 'Список') ||
    pageSections.find(section =>
        pageNames.some(([name]) => section.querySelector('h2').firstChild.textContent.trim() === name)
    );

if (initial) {
    const index = pageNames.findIndex(
        ([name]) => initial.querySelector('h2').firstChild.textContent.trim() === name
    );
    const button = [...pageFooter.querySelectorAll('button')].find(
        button => button.textContent === pageNames[index][1]
    );

    selectPage(initial, button);
}

const sizing = document.createElement('details');
sizing.className = 'column-settings';
sizing.innerHTML =
    '<summary>Ширина столбцов</summary><div class="column-controls"></div>' +
    '<button type="button" class="reset-widths">Сбросить размеры листа</button>';
document.querySelector('nav').before(sizing);

const widthStoreKey = 'excel-diff-column-widths-v1';
let savedWidths = Object.create(null);

try {
    savedWidths = JSON.parse(localStorage.getItem(widthStoreKey) || '{}');
    if (!savedWidths || typeof savedWidths !== 'object') {
        savedWidths = {};
    }
}
catch (error) {
    savedWidths = {};
}

function sheetName(section) {
    return section.querySelector('h2').firstChild.textContent.trim();
}

function columnNames(section) {
    return [...(section.querySelector('article .row-cells')?.children || [])].map(
        cell => cell.querySelector('.cell-heading b').textContent
    );
}

function applyWidth(section, index, width) {
    for (const article of section.querySelectorAll('article')) {
        const cell = article.querySelector('.row-cells')?.children[index];
        if (cell) {
            cell.style.width = width + 'px';
            cell.style.flexBasis = width + 'px';
        }
    }
}

for (const section of pageSections) {
    const values = savedWidths[sheetName(section)] || {};
    columnNames(section).forEach((name, index) => {
        const value = Number(values[name]);
        if (Number.isFinite(value) && value >= 60 && value <= 600) {
            applyWidth(section, index, value);
        }
    });
}

function currentSection() {
    return pageSections.find(section => !section.hidden);
}

function saveWidths() {
    try {
        localStorage.setItem(widthStoreKey, JSON.stringify(savedWidths));
        sizing.removeAttribute('data-save-error');
    }
    catch (error) {
        sizing.setAttribute('data-save-error', 'Браузер запретил сохранение размеров');
    }
}

function renderSizing() {
    const section = currentSection();
    const controls = sizing.querySelector('.column-controls');
    controls.replaceChildren();

    if (!section) {
        return;
    }

    const sheet = sheetName(section);
    const values = savedWidths[sheet] || {};

    columnNames(section).forEach((name, index) => {
        const label = document.createElement('label');
        const title = document.createElement('span');
        title.textContent = name;

        const slider = document.createElement('input');
        slider.type = 'range';
        slider.min = '60';
        slider.max = '600';
        slider.step = '1';
        slider.value = values[name] || 170;
        slider.setAttribute('aria-label', 'Ширина ' + name);

        const number = document.createElement('input');
        number.type = 'number';
        number.min = '60';
        number.max = '600';
        number.step = '1';
        number.value = slider.value;
        number.setAttribute('aria-label', 'Ширина ' + name + ' в пикселях');

        function update(value) {
            const width = Math.min(600, Math.max(60, Number(value)));
            if (!Number.isFinite(width)) {
                return;
            }

            slider.value = width;
            number.value = width;
            applyWidth(section, index, width);

            if (!Object.hasOwn(savedWidths, sheet)) {
                savedWidths[sheet] = Object.create(null);
            }

            savedWidths[sheet][name] = width;
            saveWidths();
        }

        slider.oninput = () => update(slider.value);
        number.oninput = () => {
            if (number.value !== '' && Number(number.value) >= 60 && Number(number.value) <= 600) {
                update(number.value);
            }
        };
        number.onchange = () => update(number.value);

        label.append(title, slider, number);
        controls.append(label);
    });
}

sizing.querySelector('.reset-widths').onclick = () => {
    const section = currentSection();
    if (!section) {
        return;
    }

    delete savedWidths[sheetName(section)];
    columnNames(section).forEach((name, index) => applyWidth(section, index, 170));
    saveWidths();
    renderSizing();
};

pageFooter.addEventListener('click', event => {
    if (event.target.closest('.footer-tabs button')) {
        renderSizing();
    }
});
renderSizing();

const footerTabs = document.createElement('div');
footerTabs.className = 'footer-tabs';

for (const button of [...pageFooter.children]) {
    if (button.tagName === 'BUTTON') {
        footerTabs.append(button);
    }
}

pageFooter.prepend(footerTabs);

const footerSizing = sizing;
pageFooter.append(footerSizing);

function updateFooterSpace() {
    document.body.style.paddingBottom = (pageFooter.offsetHeight + 10) + 'px';
}

footerSizing.addEventListener('toggle', updateFooterSpace);
new ResizeObserver(updateFooterSpace).observe(pageFooter);
updateFooterSpace();

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
footerSizing.append(settingsActions);

exportWidths.onclick = () => {
    const allWidths = Object.create(null);

    for (const section of pageSections) {
        const sheet = sheetName(section);
        allWidths[sheet] = Object.create(null);

        const cells = section.querySelector('article .row-cells')?.children || [];
        columnNames(section).forEach((name, index) => {
            const width =
                Number.parseFloat(cells[index]?.style.width) ||
                Number(savedWidths[sheet]?.[name]) ||
                170;

            allWidths[sheet][name] = width;
        });
    }

    const payload = { version: 1, columnWidths: allWidths };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
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
        if (
            payload.version !== 1 ||
            !payload.columnWidths ||
            typeof payload.columnWidths !== 'object' ||
            Array.isArray(payload.columnWidths)
        ) {
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

        savedWidths = validated;

        for (const section of pageSections) {
            const values = savedWidths[sheetName(section)] || {};
            columnNames(section).forEach((name, index) => applyWidth(section, index, values[name] ?? 170));
        }

        saveWidths();
        renderSizing();
        updateFooterSpace();
        settingsMessage.textContent = 'Настройки импортированы';
    }
    catch (error) {
        settingsMessage.textContent = 'Ошибка импорта: ' + error.message;
    }
    finally {
        settingsFile.value = '';
    }
};