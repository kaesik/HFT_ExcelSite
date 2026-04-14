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

export async function getWorkbooks() {
    return fetchJson("/api/workbooks");
}

export async function getWorkbookSummary(workbookName) {
    return fetchJson(`/api/workbooks/${encodeURIComponent(workbookName)}/summary`);
}

export async function getSheetSummary(workbookName, sheetName) {
    return fetchJson(
        `/api/workbooks/${encodeURIComponent(workbookName)}/sheet/${encodeURIComponent(sheetName)}/summary`
    );
}

export async function getColoredCells(workbookName, sheetName) {
    return fetchJson(
        `/api/workbooks/${encodeURIComponent(workbookName)}/sheet/${encodeURIComponent(sheetName)}/colored-cells`
    );
}

export async function getFormulas(workbookName, sheetName, limit = 100) {
    return fetchJson(
        `/api/workbooks/${encodeURIComponent(workbookName)}/sheet/${encodeURIComponent(sheetName)}/formulas?limit=${limit}`
    );
}