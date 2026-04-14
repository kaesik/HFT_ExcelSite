const ALLOWED_COLOR_ROLES = new Set([
    "display",
    "result",
    "input_1",
    "input_2",
    "dropdown",
]);

const COLOR_ROLE_MAP = {
    "#FFFF00": "display",
    "#0070C0": "result",
    "#92D050": "input_1",
    "#00B050": "input_2",
    "#00B0F0": "dropdown",
};

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

function normalizeRoleLabel(role) {
    if (!role) {
        return "brak";
    }

    if (role === "input_1") {
        return "input 1";
    }

    if (role === "input_2") {
        return "input 2";
    }

    return role;
}

function normalizeHexColor(color) {
    if (!color) {
        return "";
    }

    return String(color).trim().toUpperCase();
}

function resolveColorRole(item) {
    if (item?.colorRole && ALLOWED_COLOR_ROLES.has(item.colorRole)) {
        return item.colorRole;
    }

    const normalizedColor = normalizeHexColor(item?.fillColor);
    return COLOR_ROLE_MAP[normalizedColor] ?? null;
}

function getDisplayValue(item) {
    if (item.cachedValue !== null && item.cachedValue !== undefined && item.cachedValue !== "") {
        return item.cachedValue;
    }

    if (item.value !== null && item.value !== undefined && item.value !== "") {
        return item.value;
    }

    return "—";
}

function isColoredCandidate(item) {
    return Boolean(resolveColorRole(item));
}

function getAllColoredCandidates(state) {
    if (!state.activeCandidates) {
        return [];
    }

    return (state.activeCandidates.inputs ?? [])
        .filter(isColoredCandidate)
        .map((item) => {
            const colorRole = resolveColorRole(item);

            if (item.colorRole === colorRole) {
                return item;
            }

            return {
                ...item,
                colorRole,
            };
        });
}

function getFilteredCandidates(state) {
    const allItems = getAllColoredCandidates(state);

    if (state.activeFilter === "all") {
        return allItems;
    }

    return allItems.filter((item) => item.colorRole === state.activeFilter);
}

export function createRenderer(elements, state, actions) {
    function renderError(container, error) {
        if (!container) {
            return;
        }

        container.innerHTML = `<div class="error-box">${escapeHtml(error.message || "Wystąpił błąd.")}</div>`;
    }

    function renderWorkbookInfo() {
        if (!elements.workbookInfoElement) {
            return;
        }

        if (!state.workbooks.length) {
            elements.workbookInfoElement.innerHTML = '<div class="placeholder">Brak plików Excel.</div>';
            return;
        }

        const workbookButtonsHtml = state.workbooks
            .map((workbook) => {
                const fileName = workbook.fileName;
                const isActive = state.activeWorkbook === fileName;

                return `
                    <button
                        class="sheet-button ${isActive ? "active" : ""}"
                        type="button"
                        data-workbook-name="${escapeHtml(fileName)}"
                    >
                        ${escapeHtml(fileName)}
                    </button>
                `;
            })
            .join("");

        elements.workbookInfoElement.innerHTML = `
            <div><strong>Liczba Exceli:</strong> ${state.workbooks.length}</div>
            <div class="sheet-list" style="margin-top: 12px;">
                ${workbookButtonsHtml}
            </div>
        `;

        elements.workbookInfoElement.querySelectorAll("[data-workbook-name]").forEach((button) => {
            button.addEventListener("click", async () => {
                const workbookName = button.dataset.workbookName;
                await actions.selectWorkbook(workbookName);
            });
        });
    }

    function renderSheetList() {
        if (!elements.sheetListElement) {
            return;
        }

        if (!state.activeWorkbook) {
            elements.sheetListElement.innerHTML = '<div class="placeholder">Najpierw wybierz Excel.</div>';
            return;
        }

        const sheets = state.visibleSheets ?? [];

        if (!sheets.length) {
            elements.sheetListElement.innerHTML = '<div class="placeholder">Brak arkuszy z obsługiwanymi kolorami.</div>';
            return;
        }

        elements.sheetListElement.innerHTML = sheets
            .map((sheet) => {
                const isActive = state.activeSheet === sheet.title;

                return `
                    <button
                        class="sheet-button ${isActive ? "active" : ""}"
                        type="button"
                        data-sheet-name="${escapeHtml(sheet.title)}"
                    >
                        ${escapeHtml(sheet.title)}
                    </button>
                `;
            })
            .join("");

        elements.sheetListElement.querySelectorAll(".sheet-button").forEach((button) => {
            button.addEventListener("click", async () => {
                const sheetName = button.dataset.sheetName;
                await actions.selectSheet(sheetName);
            });
        });
    }

    function renderSheetSummary() {
        if (!elements.sheetSummaryElement) {
            return;
        }

        if (!state.activeSheetSummary) {
            elements.sheetSummaryElement.innerHTML = '<div class="placeholder">Brak danych.</div>';
            return;
        }

        const sheet = state.activeSheetSummary.sheet;

        elements.sheetSummaryElement.innerHTML = `
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
        if (!elements.cellDetailsElement) {
            return;
        }

        if (!state.selectedCell) {
            elements.cellDetailsElement.innerHTML = '<div class="placeholder">Kliknij pole z listy.</div>';
            return;
        }

        const cell = state.selectedCell;
        const validationInfo = cell.validationInfo || null;
        const resolvedRole = resolveColorRole(cell);

        elements.cellDetailsElement.innerHTML = `
            <div class="detail-item">
                <div class="detail-label">Adres</div>
                <div class="detail-value">${escapeHtml(cell.address)}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">Source address</div>
                <div class="detail-value">${escapeHtml(cell.sourceAddress)}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">Label</div>
                <div class="detail-value">${escapeHtml(cell.label)}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">Rola</div>
                <div class="detail-value">${escapeHtml(normalizeRoleLabel(resolvedRole))}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">Kolor</div>
                <div class="detail-value">${escapeHtml(cell.fillColor)}</div>
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
                <div class="detail-label">Input type</div>
                <div class="detail-value">${escapeHtml(cell.inputType)}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">Has dropdown</div>
                <div class="detail-value">${escapeHtml(cell.hasDropdown)}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">Validation formula</div>
                <div class="detail-value">${escapeHtml(validationInfo?.formula1)}</div>
            </div>
        `;
    }

    function renderCandidates() {
        if (!elements.candidatesContainerElement || !elements.candidatesMetaElement) {
            return;
        }

        if (!state.activeCandidates) {
            elements.candidatesContainerElement.innerHTML = '<div class="placeholder">Wybierz arkusz, aby zobaczyć dane.</div>';
            elements.candidatesMetaElement.textContent = "";
            return;
        }

        const allColoredItems = getAllColoredCandidates(state);
        const filteredItems = getFilteredCandidates(state);

        elements.candidatesMetaElement.textContent =
            `Kolorowe pola: ${allColoredItems.length}, widoczne: ${filteredItems.length}`;

        if (!filteredItems.length) {
            elements.candidatesContainerElement.innerHTML = '<div class="placeholder">Brak kolorowych elementów dla wybranego filtra.</div>';
            return;
        }

        elements.candidatesContainerElement.innerHTML = filteredItems
            .map((item) => {
                const safeCellJson = encodeURIComponent(JSON.stringify(item));
                const displayValue = getDisplayValue(item);
                const isSelected =
                    state.selectedCell?.address === item.address &&
                    state.selectedCell?.sourceAddress === item.sourceAddress;

                return `
                    <div
                        class="candidate-card role-${escapeHtml(item.colorRole || "none")} ${isSelected ? "selected" : ""}"
                        data-cell="${safeCellJson}"
                    >
                        <div class="candidate-header">
                            <div class="candidate-address">${escapeHtml(item.address)}</div>
                            <div class="candidate-role">${escapeHtml(normalizeRoleLabel(item.colorRole))}</div>
                        </div>

                        <div class="candidate-label">${escapeHtml(item.label || "Brak etykiety")}</div>
                        <div class="candidate-value">${escapeHtml(displayValue)}</div>

                        <div class="candidate-meta">
                            <span class="meta-badge">kolor: ${escapeHtml(item.fillColor || "brak")}</span>
                            ${item.hasFormula ? '<span class="meta-badge">formuła</span>' : ""}
                            ${item.hasDropdown ? '<span class="meta-badge">dropdown</span>' : ""}
                            ${item.inputType ? `<span class="meta-badge">typ: ${escapeHtml(item.inputType)}</span>` : ""}
                        </div>
                    </div>
                `;
            })
            .join("");

        elements.candidatesContainerElement.querySelectorAll(".candidate-card").forEach((card) => {
            card.addEventListener("click", () => {
                const rawCell = card.getAttribute("data-cell");
                if (!rawCell) {
                    return;
                }

                state.selectedCell = JSON.parse(decodeURIComponent(rawCell));
                renderCellDetails();
                renderCandidates();
            });
        });
    }

    function renderFormulas(payload) {
        if (!elements.formulasListElement) {
            return;
        }

        if (!payload?.cells?.length) {
            elements.formulasListElement.innerHTML = '<div class="placeholder">Brak formuł.</div>';
            return;
        }

        elements.formulasListElement.innerHTML = payload.cells
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

    function renderFilterButtons() {
        document.querySelectorAll(".filter-button").forEach((button) => {
            button.classList.toggle("active", button.dataset.filter === state.activeFilter);
        });
    }

    function renderInitialFormulasPlaceholder() {
        if (!elements.formulasListElement) {
            return;
        }

        elements.formulasListElement.innerHTML = '<div class="placeholder">Formuły nie zostały jeszcze wczytane.</div>';
    }

    function renderSheetFormulasPlaceholder() {
        if (!elements.formulasListElement) {
            return;
        }

        elements.formulasListElement.innerHTML = '<div class="placeholder">Kliknij "Wczytaj formuły", aby zobaczyć formuły z arkusza.</div>';
    }

    function renderFormulasLoading() {
        if (!elements.formulasListElement) {
            return;
        }

        elements.formulasListElement.innerHTML = '<div class="placeholder">Ładowanie formuł...</div>';
    }

    return {
        renderError,
        renderWorkbookInfo,
        renderSheetList,
        renderSheetSummary,
        renderCellDetails,
        renderCandidates,
        renderFormulas,
        renderFilterButtons,
        renderInitialFormulasPlaceholder,
        renderSheetFormulasPlaceholder,
        renderFormulasLoading,
    };
}