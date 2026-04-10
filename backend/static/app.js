const state = {
    workbooks: [],
    activeWorkbook: null,
    workbookSummary: null,
    activeSheet: null,
    activeSheetSummary: null,
    activeRangeData: null,
    selectedCell: null,
};

const workbookInfoElement = document.getElementById("workbook-info");
const sheetListElement = document.getElementById("sheet-list");
const activeSheetTitleElement = document.getElementById("active-sheet-title");
const activeSheetMetaElement = document.getElementById("active-sheet-meta");
const sheetSummaryElement = document.getElementById("sheet-summary");
const tableContainerElement = document.getElementById("table-container");
const cellDetailsElement = document.getElementById("cell-details");
const formulasListElement = document.getElementById("formulas-list");
const rangeMetaElement = document.getElementById("range-meta");
const rangeFormElement = document.getElementById("range-form");
const loadFormulasButtonElement = document.getElementById("load-formulas-button");

const startRowElement = document.getElementById("start-row");
const endRowElement = document.getElementById("end-row");
const startColumnElement = document.getElementById("start-column");
const endColumnElement = document.getElementById("end-column");
const includeEmptyElement = document.getElementById("include-empty");

async function fetchJson(url) {
    const response = await fetch(url);

    if (!response.ok) {
        let message = `HTTP ${response.status}`;

        try {
            const payload = await response.json();
            if (payload?.detail) {
                message = payload.detail;
            }
        } catch {
        }

        throw new Error(message);
    }

    return response.json();
}

function escapeHtml(value) {
    if (value === null || value === undefined) {
        return "";
    }

    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function columnNumberToName(columnNumber) {
    let result = "";
    let current = columnNumber;

    while (current > 0) {
        const remainder = (current - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        current = Math.floor((current - 1) / 26);
    }

    return result;
}

function renderWorkbookInfo() {
    if (!state.workbooks.length) {
        workbookInfoElement.innerHTML = '<div class="placeholder">Brak plików Excel.</div>';
        return;
    }

    const activeWorkbookText = state.activeWorkbook
        ? `<div><strong>Wybrany plik:</strong> ${escapeHtml(state.activeWorkbook)}</div>`
        : `<div><strong>Wybrany plik:</strong> brak</div>`;

    workbookInfoElement.innerHTML = `
        <div><strong>Liczba Exceli:</strong> ${state.workbooks.length}</div>
        ${activeWorkbookText}
    `;
}

function renderSheetList() {
    if (!state.activeWorkbook) {
        sheetListElement.innerHTML = '<div class="placeholder">Najpierw wybierz Excel.</div>';
        return;
    }

    const sheets = state.workbookSummary?.sheets ?? [];

    if (!sheets.length) {
        sheetListElement.innerHTML = '<div class="placeholder">Brak arkuszy.</div>';
        return;
    }

    sheetListElement.innerHTML = sheets
        .map((sheet) => {
            const isActive = state.activeSheet === sheet.title;
            const isDisabled = !sheet.hasFormulas;

            return `
                <button
                    class="sheet-button ${isActive ? "active" : ""}"
                    type="button"
                    data-sheet-name="${escapeHtml(sheet.title)}"
                    data-has-formulas="${sheet.hasFormulas}"
                    ${isDisabled ? 'disabled style="opacity:0.5; cursor:not-allowed;"' : ""}
                >
                    ${escapeHtml(sheet.title)}
                    ${sheet.hasFormulas ? "" : " (brak formuł)"}
                </button>
            `;
        })
        .join("");

    sheetListElement.querySelectorAll(".sheet-button").forEach((button) => {
        button.addEventListener("click", async () => {
            const sheetName = button.dataset.sheetName;
            const hasFormulas = button.dataset.hasFormulas === "true";

            if (!hasFormulas) {
                return;
            }

            await selectSheet(sheetName);
        });
    });
}

function renderSheetSummary() {
    if (!state.activeSheetSummary) {
        sheetSummaryElement.innerHTML = '<div class="placeholder">Brak danych.</div>';
        return;
    }

    const sheet = state.activeSheetSummary.sheet;

    sheetSummaryElement.innerHTML = `
        <div class="detail-item">
            <div class="detail-label">Tytuł</div>
            <div class="detail-value">${escapeHtml(sheet.title)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Wymiary</div>
            <div class="detail-value">${escapeHtml(sheet.dimensions)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Max row</div>
            <div class="detail-value">${escapeHtml(sheet.maxRow)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Max column</div>
            <div class="detail-value">${escapeHtml(sheet.maxColumn)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Ma formuły</div>
            <div class="detail-value">${escapeHtml(sheet.hasFormulas)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Liczba formuł</div>
            <div class="detail-value">${escapeHtml(sheet.formulaCellsCount)}</div>
        </div>
    `;
}

function renderCellDetails() {
    if (!state.selectedCell) {
        cellDetailsElement.innerHTML = '<div class="placeholder">Kliknij komórkę w tabeli.</div>';
        return;
    }

    const cell = state.selectedCell;

    cellDetailsElement.innerHTML = `
        <div class="detail-item">
            <div class="detail-label">Adres</div>
            <div class="detail-value">${escapeHtml(cell.address)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Requested address</div>
            <div class="detail-value">${escapeHtml(cell.requestedAddress)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Wartość</div>
            <div class="detail-value">${escapeHtml(cell.value)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Cached value</div>
            <div class="detail-value">${escapeHtml(cell.cachedValue)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Formuła</div>
            <div class="detail-value">${escapeHtml(cell.formula)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Data type</div>
            <div class="detail-value">${escapeHtml(cell.dataType)}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Number format</div>
            <div class="detail-value">${escapeHtml(cell.numberFormat)}</div>
        </div>
    `;
}

function renderTable() {
    if (!state.activeRangeData) {
        tableContainerElement.innerHTML = '<div class="placeholder">Wybierz arkusz, aby zobaczyć dane.</div>';
        rangeMetaElement.textContent = "";
        return;
    }

    const rangeData = state.activeRangeData;
    const cells = rangeData.cells;
    const range = rangeData.range;

    const cellsMap = new Map();
    cells.forEach((cell) => {
        cellsMap.set(`${cell.requestedRow}-${cell.requestedColumn}`, cell);
    });

    rangeMetaElement.textContent =
        `Zakres: rzędy ${range.startRow}-${range.endRow}, kolumny ${range.startColumn}-${range.endColumn}, zwrócone komórki: ${range.returnedCellsCount}`;

    let headerHtml = '<tr><th class="corner-cell">#</th>';
    for (let columnIndex = range.startColumn; columnIndex <= range.endColumn; columnIndex += 1) {
        headerHtml += `<th>${escapeHtml(columnNumberToName(columnIndex))}</th>`;
    }
    headerHtml += "</tr>";

    let bodyHtml = "";

    for (let rowIndex = range.startRow; rowIndex <= range.endRow; rowIndex += 1) {
        bodyHtml += `<tr><th class="row-header">${rowIndex}</th>`;

        for (let columnIndex = range.startColumn; columnIndex <= range.endColumn; columnIndex += 1) {
            const key = `${rowIndex}-${columnIndex}`;
            const cell = cellsMap.get(key);

            if (!cell) {
                bodyHtml += `<td class="cell-empty"></td>`;
                continue;
            }

            const classes = [cell.formula ? "cell-has-formula" : ""].join(" ").trim();
            const displayValue =
                cell.cachedValue !== null && cell.cachedValue !== undefined && cell.cachedValue !== ""
                    ? cell.cachedValue
                    : cell.value;

            const safeCellJson = encodeURIComponent(JSON.stringify(cell));

            bodyHtml += `
                <td class="${classes}">
                    <button
                        type="button"
                        class="cell-button"
                        data-cell="${safeCellJson}"
                    >
                        <span class="cell-address">${escapeHtml(cell.requestedAddress ?? cell.address)}</span>
                        <span class="cell-value">${escapeHtml(displayValue)}</span>
                    </button>
                </td>
            `;
        }

        bodyHtml += "</tr>";
    }

    tableContainerElement.innerHTML = `
        <table class="sheet-table">
            <thead>${headerHtml}</thead>
            <tbody>${bodyHtml}</tbody>
        </table>
    `;

    tableContainerElement.querySelectorAll(".cell-button").forEach((button) => {
        button.addEventListener("click", () => {
            const rawCell = button.getAttribute("data-cell");
            if (!rawCell) {
                return;
            }

            state.selectedCell = JSON.parse(decodeURIComponent(rawCell));
            renderCellDetails();
        });
    });
}

function renderFormulas(payload) {
    if (!payload?.cells?.length) {
        formulasListElement.innerHTML = '<div class="placeholder">Brak formuł.</div>';
        return;
    }

    formulasListElement.innerHTML = payload.cells
        .map((cell) => {
            return `
                <div class="formula-item">
                    <div><strong>${escapeHtml(cell.requestedAddress ?? cell.address)}</strong></div>
                    <div class="muted">${escapeHtml(cell.formula)}</div>
                    <div>Wynik: ${escapeHtml(cell.cachedValue)}</div>
                </div>
            `;
        })
        .join("");
}

function renderError(container, error) {
    container.innerHTML = `<div class="error-box">${escapeHtml(error.message || "Wystąpił błąd.")}</div>`;
}

function renderWorkbookButtons() {
    if (!state.workbooks.length) {
        workbookInfoElement.innerHTML = '<div class="placeholder">Brak plików Excel.</div>';
        return;
    }

    const workbookButtonsHtml = state.workbooks
        .map((workbook) => {
            const isActive = state.activeWorkbook === workbook.fileName;
            return `
                <button
                    class="sheet-button ${isActive ? "active" : ""}"
                    type="button"
                    data-workbook-name="${escapeHtml(workbook.fileName)}"
                >
                    ${escapeHtml(workbook.fileName)}
                </button>
            `;
        })
        .join("");

    workbookInfoElement.innerHTML = `
        <div><strong>Liczba Exceli:</strong> ${state.workbooks.length}</div>
        <div class="sheet-list" style="margin-top: 12px;">
            ${workbookButtonsHtml}
        </div>
    `;

    workbookInfoElement.querySelectorAll("[data-workbook-name]").forEach((button) => {
        button.addEventListener("click", async () => {
            const workbookName = button.dataset.workbookName;
            await selectWorkbook(workbookName);
        });
    });
}

async function loadWorkbooks() {
    try {
        const payload = await fetchJson("/api/workbooks");
        state.workbooks = payload.workbooks ?? [];
        renderWorkbookInfo();
        renderWorkbookButtons();
        renderSheetList();
    } catch (error) {
        renderError(workbookInfoElement, error);
        renderError(sheetListElement, error);
    }
}

async function selectWorkbook(workbookName) {
    state.activeWorkbook = workbookName;
    state.workbookSummary = null;
    state.activeSheet = null;
    state.activeSheetSummary = null;
    state.activeRangeData = null;
    state.selectedCell = null;

    renderWorkbookInfo();
    renderWorkbookButtons();
    renderCellDetails();
    renderTable();
    formulasListElement.innerHTML = '<div class="placeholder">Formuły nie zostały jeszcze wczytane.</div>';

    activeSheetTitleElement.textContent = workbookName;
    activeSheetMetaElement.textContent = "Ładowanie arkuszy...";

    try {
        const workbookSummary = await fetchJson(`/api/workbooks/${encodeURIComponent(workbookName)}/summary`);
        state.workbookSummary = workbookSummary;
        renderSheetList();
        activeSheetMetaElement.textContent = `Arkusze: ${workbookSummary.sheetCount}`;
    } catch (error) {
        renderError(sheetListElement, error);
        activeSheetMetaElement.textContent = "Błąd ładowania workbooka";
    }
}

async function selectSheet(sheetName) {
    if (!state.activeWorkbook) {
        renderError(sheetSummaryElement, new Error("Najpierw wybierz Excel."));
        return;
    }

    const selectedSheet = state.workbookSummary?.sheets?.find((sheet) => sheet.title === sheetName);

    if (!selectedSheet?.hasFormulas) {
        renderError(sheetSummaryElement, new Error("Ten sheet nie zawiera formuł i nie będzie wczytywany."));
        return;
    }

    state.activeSheet = sheetName;
    state.activeSheetSummary = null;
    state.activeRangeData = null;
    state.selectedCell = null;

    renderSheetList();
    renderCellDetails();
    renderTable();
    formulasListElement.innerHTML = '<div class="placeholder">Kliknij "Wczytaj formuły", aby zobaczyć formuły z arkusza.</div>';

    activeSheetTitleElement.textContent = `${state.activeWorkbook} / ${sheetName}`;
    activeSheetMetaElement.textContent = "Ładowanie summary...";

    try {
        const summary = await fetchJson(
            `/api/workbooks/${encodeURIComponent(state.activeWorkbook)}/sheet/${encodeURIComponent(sheetName)}/summary`
        );

        state.activeSheetSummary = summary;
        renderSheetSummary();

        activeSheetMetaElement.textContent =
            `${summary.sheet.dimensions} | formuły: ${summary.sheet.formulaCellsCount}`;

        startRowElement.value = "1";
        endRowElement.value = String(Math.min(summary.sheet.maxRow, 30));
        startColumnElement.value = "1";
        endColumnElement.value = String(Math.min(summary.sheet.maxColumn, 12));

        await loadRange();
    } catch (error) {
        renderError(sheetSummaryElement, error);
        activeSheetMetaElement.textContent = "Błąd ładowania";
    }
}

async function loadRange() {
    if (!state.activeWorkbook) {
        renderError(tableContainerElement, new Error("Najpierw wybierz Excel."));
        return;
    }

    if (!state.activeSheet) {
        renderError(tableContainerElement, new Error("Najpierw wybierz arkusz."));
        return;
    }

    const params = new URLSearchParams({
        start_row: startRowElement.value,
        end_row: endRowElement.value,
        start_column: startColumnElement.value,
        end_column: endColumnElement.value,
        include_empty: String(includeEmptyElement.checked),
    });

    tableContainerElement.innerHTML = '<div class="placeholder">Ładowanie zakresu...</div>';

    try {
        const payload = await fetchJson(
            `/api/workbooks/${encodeURIComponent(state.activeWorkbook)}/sheet/${encodeURIComponent(state.activeSheet)}/range?${params.toString()}`
        );

        state.activeRangeData = payload;
        state.selectedCell = null;
        renderCellDetails();
        renderTable();
    } catch (error) {
        renderError(tableContainerElement, error);
        rangeMetaElement.textContent = "";
    }
}

async function loadFormulas() {
    if (!state.activeWorkbook) {
        renderError(formulasListElement, new Error("Najpierw wybierz Excel."));
        return;
    }

    if (!state.activeSheet) {
        renderError(formulasListElement, new Error("Najpierw wybierz arkusz."));
        return;
    }

    formulasListElement.innerHTML = '<div class="placeholder">Ładowanie formuł...</div>';

    try {
        const payload = await fetchJson(
            `/api/workbooks/${encodeURIComponent(state.activeWorkbook)}/sheet/${encodeURIComponent(state.activeSheet)}/formulas?limit=100`
        );
        renderFormulas(payload);
    } catch (error) {
        renderError(formulasListElement, error);
    }
}

rangeFormElement.addEventListener("submit", async (event) => {
    event.preventDefault();
    await loadRange();
});

loadFormulasButtonElement.addEventListener("click", async () => {
    await loadFormulas();
});

loadWorkbooks();