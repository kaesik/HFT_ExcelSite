export const initialState = {
    workbooks: [],
    activeWorkbook: null,
    workbookSummary: null,
    visibleSheets: [],
    activeSheet: null,
    activeSheetSummary: null,
    activeCandidates: null,
    selectedCell: null,
    activeFilter: "all",
    cache: {
        workbooks: null,
        workbookSummaries: {},
        sheetSummaries: {},
        coloredCells: {},
        formulas: {},
        visibleSheetsByWorkbook: {},
    },
};

export function createState() {
    return structuredClone(initialState);
}

export function resetWorkbookState(state, workbookName) {
    state.activeWorkbook = workbookName;
    state.workbookSummary = null;
    state.visibleSheets = [];
    state.activeSheet = null;
    state.activeSheetSummary = null;
    state.activeCandidates = null;
    state.selectedCell = null;
    state.activeFilter = "all";
}

export function resetSheetState(state, sheetName) {
    state.activeSheet = sheetName;
    state.activeSheetSummary = null;
    state.activeCandidates = null;
    state.selectedCell = null;
    state.activeFilter = "all";
}

export function getWorkbookCacheKey(workbookName) {
    return workbookName;
}

export function getSheetCacheKey(workbookName, sheetName) {
    return `${workbookName}::${sheetName}`;
}