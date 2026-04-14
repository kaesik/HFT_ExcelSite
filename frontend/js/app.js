import {
    getWorkbooks,
    getWorkbookSummary,
    getSheetSummary,
    getColoredCells,
    getFormulas,
} from "./api.js";
import {
    createState,
    resetWorkbookState,
    resetSheetState,
    getWorkbookCacheKey,
    getSheetCacheKey,
} from "./state.js";
import { createRenderer } from "./render.js";

const COLOR_ROLE_MAP = {
    "#FFFF00": "display",
    "#0070C0": "result",
    "#92D050": "input_1",
    "#00B050": "input_2",
    "#00B0F0": "dropdown",
};

function getElements() {
    return {
        workbookInfoElement: document.getElementById("workbook-info"),
        sheetListElement: document.getElementById("sheet-list"),
        activeSheetTitleElement: document.getElementById("active-sheet-title"),
        activeSheetMetaElement: document.getElementById("active-sheet-meta"),
        sheetSummaryElement: document.getElementById("sheet-summary"),
        candidatesContainerElement: document.getElementById("candidates-container"),
        cellDetailsElement: document.getElementById("cell-details"),
        formulasListElement: document.getElementById("formulas-list"),
        candidatesMetaElement: document.getElementById("candidates-meta"),
        loadFormulasButtonElement: document.getElementById("load-formulas-button"),
    };
}

function normalizeHexColor(color) {
    if (!color) {
        return "";
    }

    return String(color).trim().toUpperCase();
}

function hasSupportedColor(cell) {
    if (!cell) {
        return false;
    }

    if (cell.colorRole && Object.values(COLOR_ROLE_MAP).includes(cell.colorRole)) {
        return true;
    }

    const normalizedColor = normalizeHexColor(cell.fillColor);
    return Boolean(COLOR_ROLE_MAP[normalizedColor]);
}

function hasSupportedColorsInPayload(payload) {
    const cells = payload?.cells ?? [];
    return cells.some(hasSupportedColor);
}

function createApplication() {
    const state = createState();
    const elements = getElements();

    const actions = {
        selectWorkbook,
        selectSheet,
        loadFormulas,
    };

    const renderer = createRenderer(elements, state, actions);

    function bindFilterButtons() {
        document.querySelectorAll(".filter-button").forEach((button) => {
            button.addEventListener("click", () => {
                const filter = button.dataset.filter;
                state.activeFilter = filter;
                renderer.renderFilterButtons();
                renderer.renderCandidates();
            });
        });
    }

    async function loadWorkbooks() {
        try {
            if (elements.workbookInfoElement) {
                elements.workbookInfoElement.innerHTML = '<div class="muted">Ładowanie...</div>';
            }

            if (state.cache.workbooks) {
                state.workbooks = state.cache.workbooks.workbooks ?? [];
                renderer.renderWorkbookInfo();
                renderer.renderSheetList();
                return;
            }

            const payload = await getWorkbooks();
            state.cache.workbooks = payload;
            state.workbooks = payload.workbooks ?? [];

            renderer.renderWorkbookInfo();
            renderer.renderSheetList();
        } catch (error) {
            renderer.renderError(elements.workbookInfoElement, error);
            renderer.renderError(elements.sheetListElement, error);
            console.error("Błąd loadWorkbooks:", error);
        }
    }

    async function getVisibleSheetsForWorkbook(workbookName, workbookSummary) {
        const workbookCacheKey = getWorkbookCacheKey(workbookName);

        if (state.cache.visibleSheetsByWorkbook[workbookCacheKey]) {
            return state.cache.visibleSheetsByWorkbook[workbookCacheKey];
        }

        const sheetsWithFormulas = (workbookSummary?.sheets ?? []).filter((sheet) => sheet.hasFormulas);

        const checks = await Promise.all(
            sheetsWithFormulas.map(async (sheet) => {
                const sheetCacheKey = getSheetCacheKey(workbookName, sheet.title);

                let coloredPayload = state.cache.coloredCells[sheetCacheKey];

                if (!coloredPayload) {
                    try {
                        coloredPayload = await getColoredCells(workbookName, sheet.title);
                        state.cache.coloredCells[sheetCacheKey] = coloredPayload;
                    } catch (error) {
                        console.error(`Błąd sprawdzania kolorów dla arkusza ${sheet.title}:`, error);
                        return null;
                    }
                }

                if (hasSupportedColorsInPayload(coloredPayload)) {
                    return sheet;
                }

                return null;
            })
        );

        const visibleSheets = checks.filter(Boolean);
        state.cache.visibleSheetsByWorkbook[workbookCacheKey] = visibleSheets;

        return visibleSheets;
    }

    async function selectWorkbook(workbookName) {
        resetWorkbookState(state, workbookName);

        renderer.renderWorkbookInfo();
        renderer.renderSheetList();
        renderer.renderCellDetails();
        renderer.renderCandidates();
        renderer.renderInitialFormulasPlaceholder();
        renderer.renderFilterButtons();

        if (elements.sheetSummaryElement) {
            elements.sheetSummaryElement.innerHTML = '<div class="placeholder">Brak danych.</div>';
        }

        if (elements.activeSheetTitleElement) {
            elements.activeSheetTitleElement.textContent = workbookName;
        }

        if (elements.activeSheetMetaElement) {
            elements.activeSheetMetaElement.textContent = "Sprawdzanie arkuszy...";
        }

        try {
            const workbookCacheKey = getWorkbookCacheKey(workbookName);

            let workbookSummary = state.cache.workbookSummaries[workbookCacheKey];
            if (!workbookSummary) {
                workbookSummary = await getWorkbookSummary(workbookName);
                state.cache.workbookSummaries[workbookCacheKey] = workbookSummary;
            }

            state.workbookSummary = workbookSummary;
            state.visibleSheets = await getVisibleSheetsForWorkbook(workbookName, workbookSummary);

            renderer.renderSheetList();

            if (elements.activeSheetMetaElement) {
                elements.activeSheetMetaElement.textContent =
                    `Widoczne arkusze: ${state.visibleSheets.length} z ${workbookSummary.sheetCount}`;
            }
        } catch (error) {
            renderer.renderError(elements.sheetListElement, error);

            if (elements.activeSheetMetaElement) {
                elements.activeSheetMetaElement.textContent = "Błąd ładowania workbooka";
            }

            console.error("Błąd selectWorkbook:", error);
        }
    }

    async function selectSheet(sheetName) {
        if (!state.activeWorkbook) {
            renderer.renderError(elements.sheetSummaryElement, new Error("Najpierw wybierz Excel."));
            return;
        }

        const selectedSheet = (state.visibleSheets ?? []).find((sheet) => sheet.title === sheetName);

        if (!selectedSheet) {
            renderer.renderError(
                elements.sheetSummaryElement,
                new Error("Ten sheet nie zawiera obsługiwanych kolorów.")
            );
            return;
        }

        resetSheetState(state, sheetName);

        renderer.renderFilterButtons();
        renderer.renderSheetList();
        renderer.renderCellDetails();
        renderer.renderCandidates();
        renderer.renderSheetFormulasPlaceholder();

        if (elements.sheetSummaryElement) {
            elements.sheetSummaryElement.innerHTML = '<div class="placeholder">Ładowanie podsumowania arkusza...</div>';
        }

        if (elements.activeSheetTitleElement) {
            elements.activeSheetTitleElement.textContent = `${state.activeWorkbook} / ${sheetName}`;
        }

        if (elements.activeSheetMetaElement) {
            elements.activeSheetMetaElement.textContent = "Ładowanie danych arkusza...";
        }

        try {
            const sheetCacheKey = getSheetCacheKey(state.activeWorkbook, sheetName);

            let summary = state.cache.sheetSummaries[sheetCacheKey];
            let coloredPayload = state.cache.coloredCells[sheetCacheKey];

            if (!summary || !coloredPayload) {
                const results = await Promise.all([
                    summary ? Promise.resolve(summary) : getSheetSummary(state.activeWorkbook, sheetName),
                    coloredPayload ? Promise.resolve(coloredPayload) : getColoredCells(state.activeWorkbook, sheetName),
                ]);

                summary = results[0];
                coloredPayload = results[1];

                state.cache.sheetSummaries[sheetCacheKey] = summary;
                state.cache.coloredCells[sheetCacheKey] = coloredPayload;
            }

            state.activeSheetSummary = summary;
            state.activeCandidates = {
                inputs: coloredPayload.cells ?? [],
                results: [],
            };

            renderer.renderSheetSummary();
            renderer.renderCandidates();

            if (elements.activeSheetMetaElement) {
                elements.activeSheetMetaElement.textContent =
                    `${summary.sheet.dimensions} | formuły: ${summary.sheet.formulaCellsCount}`;
            }
        } catch (error) {
            renderer.renderError(elements.sheetSummaryElement, error);
            renderer.renderError(elements.candidatesContainerElement, error);

            if (elements.activeSheetMetaElement) {
                elements.activeSheetMetaElement.textContent = "Błąd ładowania";
            }

            console.error("Błąd selectSheet:", error);
        }
    }

    async function loadFormulas() {
        if (!state.activeWorkbook) {
            renderer.renderError(elements.formulasListElement, new Error("Najpierw wybierz Excel."));
            return;
        }

        if (!state.activeSheet) {
            renderer.renderError(elements.formulasListElement, new Error("Najpierw wybierz arkusz."));
            return;
        }

        const sheetCacheKey = getSheetCacheKey(state.activeWorkbook, state.activeSheet);

        try {
            if (state.cache.formulas[sheetCacheKey]) {
                renderer.renderFormulas(state.cache.formulas[sheetCacheKey]);
                return;
            }

            renderer.renderFormulasLoading();

            const payload = await getFormulas(state.activeWorkbook, state.activeSheet, 100);
            state.cache.formulas[sheetCacheKey] = payload;

            renderer.renderFormulas(payload);
        } catch (error) {
            renderer.renderError(elements.formulasListElement, error);
            console.error("Błąd loadFormulas:", error);
        }
    }

    if (elements.loadFormulasButtonElement) {
        elements.loadFormulasButtonElement.addEventListener("click", async () => {
            await loadFormulas();
        });
    }

    bindFilterButtons();
    loadWorkbooks();
}

document.addEventListener("DOMContentLoaded", () => {
    try {
        createApplication();
    } catch (error) {
        console.error("Błąd startu aplikacji:", error);

        const workbookInfoElement = document.getElementById("workbook-info");
        if (workbookInfoElement) {
            workbookInfoElement.innerHTML = `<div class="error-box">${error.message || "Błąd startu aplikacji"}</div>`;
        }
    }
});