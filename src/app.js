// Tauri API imports
const tauriApi = window.__TAURI__;
const invoke = tauriApi?.core?.invoke?.bind(tauriApi.core) ?? null;
const open = tauriApi?.dialog?.open?.bind(tauriApi.dialog) ?? null;
const listen = tauriApi?.event?.listen?.bind(tauriApi.event) ?? null;

// DOM refs
const welcomeScreen = document.getElementById('welcome-screen');
const documentView = document.getElementById('document-view');
const desk = document.getElementById('desk');
const deskContent = document.getElementById('desk-content');
const commentsPanel = document.getElementById('comments-panel');
const commentsList = document.getElementById('comments-list');
const recentFilesSection = document.getElementById('recent-files-section');
const recentFilesList = document.getElementById('recent-files-list');
const statusBanner = document.getElementById('status-banner');
const fileInfo = document.getElementById('file-info');
const themeIcon = document.getElementById('theme-icon');
const openInWordButton = document.getElementById('open-in-word-btn');
const commentsButton = document.getElementById('toggle-comments-btn');
const findBar = document.getElementById('find-bar');
const findInput = document.getElementById('find-input');
const findCount = document.getElementById('find-count');

const DOC_ZOOM_STORAGE_KEY = 'papyr-doc-zoom';
const DOC_ZOOM_DEFAULT = 1;
const DOC_ZOOM_MIN = 0.5;
const DOC_ZOOM_MAX = 2;
const DOC_ZOOM_STEP = 0.1;
const SUPPORTED_SPREADSHEET_EXTENSIONS = ['xlsx', 'xlsm', 'xlsb', 'xls', 'csv'];

// State
let currentDocument = null;
let currentWorkbook = null;
let currentFilePath = null;
let commentsVisible = false;
let findMatches = [];
let findIndex = -1;
let statusTimer = null;
let findDebounceTimer = null;
let currentDocumentZoom = DOC_ZOOM_DEFAULT;

// --- Initialization ---

document.addEventListener('DOMContentLoaded', async () => {
    if (!hasTauriApi()) {
        reportMissingTauriApi();
        return;
    }

    initializeDocumentZoom();
    setupEventListeners();
    setupKeyboardShortcuts();
    setupTauriListeners();
    setupDragAndDrop();

    await Promise.allSettled([
        initializeTheme(),
        loadRecentFiles(),
    ]);

    invoke('show_main_window').catch(() => {});
    void openLaunchDocument();
});

function setupEventListeners() {
    document.getElementById('open-file-btn').addEventListener('click', handleOpenFile);
    document.getElementById('welcome-open-btn').addEventListener('click', handleOpenFile);
    openInWordButton.addEventListener('click', handleOpenInMicrosoftWord);
    document.getElementById('theme-toggle-btn').addEventListener('click', toggleTheme);
    commentsButton.addEventListener('click', toggleComments);
    document.getElementById('close-comments-btn').addEventListener('click', toggleComments);
    document.getElementById('find-close-btn').addEventListener('click', closeFindBar);
    document.getElementById('find-next-btn').addEventListener('click', () => navigateFind(1));
    document.getElementById('find-prev-btn').addEventListener('click', () => navigateFind(-1));
    recentFilesList.addEventListener('click', handleRecentFilesClick);
    desk.addEventListener('click', handleDeskClick);
    commentsList.addEventListener('click', handleCommentListClick);
    findInput.addEventListener('input', scheduleFind);
    findInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') navigateFind(e.shiftKey ? -1 : 1);
        if (e.key === 'Escape') closeFindBar();
    });
}

function setupKeyboardShortcuts() {
    document.addEventListener('keydown', (e) => {
        if (isZoomInShortcut(e)) {
            e.preventDefault();
            adjustDocumentZoom(DOC_ZOOM_STEP);
            return;
        }
        if (isZoomOutShortcut(e)) {
            e.preventDefault();
            adjustDocumentZoom(-DOC_ZOOM_STEP);
            return;
        }
        if (e.ctrlKey && e.key === 'o') { e.preventDefault(); handleOpenFile(); }
        if (e.ctrlKey && e.key === 'd') { e.preventDefault(); toggleTheme(); }
        if (e.ctrlKey && e.key === ']') { e.preventDefault(); toggleComments(); }
        if (e.ctrlKey && e.key === 'f') { e.preventDefault(); openFindBar(); }
        if (e.ctrlKey && e.key === 'q') {
            e.preventDefault();
            void invoke('quit_app').catch((err) => showError('Could not quit application: ' + err));
        }
    });
}

async function setupTauriListeners() {
    if (!listen) return;

    try {
        await listen('tauri://drag-drop', async (event) => {
            const paths = event.payload?.paths || [];
            const filePath = paths.find(isSupportedFilePath);
            if (filePath) {
                await loadDocument(filePath);
            }
        });
    } catch (err) {
        console.log('Drag-drop listener setup:', err);
    }
}

function setupDragAndDrop() {
    document.addEventListener('dragenter', (e) => { e.preventDefault(); document.body.classList.add('drag-over'); });
    document.addEventListener('dragover', (e) => { e.preventDefault(); });
    document.addEventListener('dragleave', (e) => { e.preventDefault(); document.body.classList.remove('drag-over'); });
    document.addEventListener('drop', (e) => { e.preventDefault(); document.body.classList.remove('drag-over'); });
}

// --- File handling ---

async function handleOpenFile() {
    if (!open) {
        showError('Tauri dialog API is unavailable.');
        return;
    }

    try {
        const selected = await open({
            multiple: false,
            filters: [
                { name: 'Supported Documents', extensions: ['docx', ...SUPPORTED_SPREADSHEET_EXTENSIONS] },
                { name: 'Word Documents', extensions: ['docx'] },
                { name: 'Excel Workbooks', extensions: SUPPORTED_SPREADSHEET_EXTENSIONS }
            ]
        });
        if (selected) {
            const path = selected.path || selected;
            await loadDocument(path);
        }
    } catch (err) {
        showError('Could not open file: ' + err);
    }
}

async function loadDocument(path) {
    if (!invoke) {
        showError('Tauri command API is unavailable.');
        return;
    }

    currentFilePath = path;
    setWordButtonAvailability(false);
    fileInfo.textContent = 'Loading...';
    showStatus('Loading document...', 'loading', 1500);

    try {
        const openedFile = await invoke('open_file', { path });
        fileInfo.textContent = '';
        renderOpenedFile(openedFile);
        setWordButtonAvailability(openedFile?.type === 'docx');
        void loadRecentFiles();
        showStatus(getFileName(path) + ' opened', 'success', 1800);
    } catch (err) {
        setWordButtonAvailability(false);
        showError(err);
    }
}

async function handleOpenInMicrosoftWord() {
    if (!invoke || !currentFilePath || !currentDocument) return;

    setWordButtonAvailability(false);
    openInWordButton.setAttribute('aria-busy', 'true');

    try {
        await invoke('open_in_microsoft_word', { path: currentFilePath });
        showStatus(`${getFileName(currentFilePath)} opened in Microsoft Word`, 'success', 1800);
    } catch (err) {
        showError(err);
    } finally {
        openInWordButton.removeAttribute('aria-busy');
        setWordButtonAvailability(Boolean(currentDocument));
    }
}

function setWordButtonAvailability(isAvailable) {
    openInWordButton.disabled = !isAvailable;
}

function showError(msg) {
    const text = typeof msg === 'string' ? msg : (msg?.toString?.() || 'Unknown error');
    fileInfo.textContent = 'Error: ' + text;
    fileInfo.style.color = '#e74c3c';
    showStatus(text, 'error');
    setTimeout(() => { fileInfo.style.color = ''; }, 5000);
}

async function loadRecentFiles() {
    if (!recentFilesList) return;
    if (!invoke) {
        renderRecentFiles([]);
        return;
    }

    try {
        const result = await invoke('get_recent_files');
        const files = Array.isArray(result)
            ? result
            : Array.isArray(result?.files)
                ? result.files
                : Array.isArray(result?.recent_files)
                    ? result.recent_files
                    : [];
        renderRecentFiles(files);
    } catch (err) {
        renderRecentFiles([]);
        console.log('Could not load recent files:', err);
    }
}

async function openLaunchDocument() {
    if (!invoke) return;

    try {
        const launchPath = await invoke('get_launch_file_path');
        if (typeof launchPath === 'string' && launchPath.trim().length > 0) {
            await loadDocument(launchPath);
        }
    } catch (err) {
        console.log('Could not read launch document path:', err);
    }
}

function renderRecentFiles(files) {
    if (!recentFilesList) return;

    const normalized = (files || [])
        .map((entry) => normalizeRecentFile(entry))
        .filter(Boolean);

    recentFilesList.replaceChildren();

    if (normalized.length === 0) {
        recentFilesList.innerHTML = '<p class="recent-files-empty">No recent files yet.</p>';
        return;
    }

    const fragment = document.createDocumentFragment();
    normalized.slice(0, 8).forEach((file) => {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'recent-file-btn';
        button.dataset.path = file.path;
        button.innerHTML = `
            <span class="recent-file-name">${escapeHtml(file.name)}</span>
            <span class="recent-file-path">${escapeHtml(file.path)}</span>
        `;
        fragment.appendChild(button);
    });

    recentFilesList.appendChild(fragment);
}

async function handleRecentFilesClick(event) {
    const button = event.target.closest('.recent-file-btn');
    if (!button) return;
    const path = button.dataset.path;
    if (path) {
        await loadDocument(path);
    }
}

function normalizeRecentFile(entry) {
    if (!entry) return null;
    if (typeof entry === 'string') {
        return { path: entry, name: getFileName(entry) || entry };
    }
    if (typeof entry === 'object') {
        const path = entry.path || entry.filePath || entry.fullPath || entry.location;
        if (!path) return null;
        return {
            path,
            name: entry.name || entry.label || getFileName(path) || path,
        };
    }
    return null;
}

function showStatus(message, kind = 'info', timeoutMs = 2200) {
    if (!statusBanner) return;

    if (statusTimer) {
        clearTimeout(statusTimer);
        statusTimer = null;
    }

    statusBanner.textContent = message;
    statusBanner.className = 'status-banner visible status-' + kind;

    if (timeoutMs !== null && timeoutMs !== undefined) {
        statusTimer = setTimeout(() => {
            if (statusBanner) {
                statusBanner.className = 'status-banner';
                statusBanner.textContent = '';
            }
        }, timeoutMs);
    }
}

// --- Document zoom ---

function initializeDocumentZoom() {
    const storedZoom = Number.parseFloat(localStorage.getItem(DOC_ZOOM_STORAGE_KEY));
    applyDocumentZoom(Number.isFinite(storedZoom) ? storedZoom : DOC_ZOOM_DEFAULT, { announce: false });
}

// Ctrl+wheel zoom only applies to DOCX. The listener is non-passive (it calls
// preventDefault), which forces wheel events onto the main thread and makes
// scrolling janky — so it is attached only while a document is shown and
// detached for spreadsheets, keeping grid scrolling on the fast path.
function setDocumentZoomWheel(enabled) {
    desk.removeEventListener('wheel', handleDocumentZoomWheel);
    if (enabled) {
        desk.addEventListener('wheel', handleDocumentZoomWheel, { passive: false });
    }
}

function handleDocumentZoomWheel(event) {
    if (!event.ctrlKey || event.deltaY === 0) return;

    event.preventDefault();
    adjustDocumentZoom(event.deltaY < 0 ? DOC_ZOOM_STEP : -DOC_ZOOM_STEP);
}

function adjustDocumentZoom(delta) {
    applyDocumentZoom(currentDocumentZoom + delta);
}

function applyDocumentZoom(zoom, options = {}) {
    const { announce = true } = options;
    const nextZoom = clampDocumentZoom(zoom);
    if (nextZoom === currentDocumentZoom && desk.style.getPropertyValue('--doc-zoom')) {
        return;
    }

    currentDocumentZoom = nextZoom;
    desk.style.setProperty('--doc-zoom', String(nextZoom));
    localStorage.setItem(DOC_ZOOM_STORAGE_KEY, String(nextZoom));
    updateDocumentImages();

    if (announce) {
        showStatus('Zoom ' + Math.round(nextZoom * 100) + '%', 'info', 900);
    }
}

function clampDocumentZoom(zoom) {
    const roundedZoom = Math.round(zoom * 100) / 100;
    return Math.min(DOC_ZOOM_MAX, Math.max(DOC_ZOOM_MIN, roundedZoom));
}

function isZoomInShortcut(event) {
    if (!event.ctrlKey || event.altKey || event.metaKey) return false;
    return event.key === '=' || event.key === '+' || event.code === 'NumpadAdd';
}

function isZoomOutShortcut(event) {
    if (!event.ctrlKey || event.altKey || event.metaKey) return false;
    return event.key === '-' || event.key === '_' || event.code === 'NumpadSubtract';
}

// --- File rendering ---

function renderOpenedFile(openedFile) {
    if (openedFile?.type === 'docx' && openedFile.document) {
        renderDocument(openedFile.document);
        return;
    }

    if (openedFile?.type === 'xlsx' && openedFile.workbook) {
        renderWorkbook(openedFile.workbook);
        return;
    }

    showError('Unsupported file response from Papyr.');
}

function renderWorkbook(workbook) {
    currentDocument = null;
    currentWorkbook = workbook;
    welcomeScreen.style.display = 'none';
    documentView.style.display = 'flex';
    deskContent.replaceChildren();
    deskContent.classList.add('spreadsheet-content');
    desk.classList.add('desk--spreadsheet');
    setDocumentZoomWheel(false);

    // Seed the per-workbook sheet cache so revisiting a tab needs no re-parse.
    if (workbook.active_sheet) {
        const cache = workbook._sheetCache || (workbook._sheetCache = new Map());
        if (!cache.has(workbook.active_sheet_index)) {
            cache.set(workbook.active_sheet_index, workbook.active_sheet);
        }
    }
    setCommentsAvailability(false);
    setCommentsVisibility(false);
    renderComments([]);
    closeFindBar();

    const sheet = workbook.active_sheet;
    const container = document.createElement('section');
    container.className = 'spreadsheet-view';

    container.appendChild(renderWorkbookHeader(workbook, sheet));
    container.appendChild(renderSheetTabs(workbook));
    container.appendChild(sheet ? renderSheetGrid(sheet, workbook.limits) : renderEmptySheetState());

    deskContent.appendChild(container);

    const filename = getFileName(currentFilePath) || 'Papyr';
    document.title = filename + ' - Papyr';
}

function renderWorkbookHeader(workbook, sheet) {
    const header = document.createElement('div');
    header.className = 'spreadsheet-header';

    const title = document.createElement('div');
    title.className = 'spreadsheet-title';
    title.textContent = getFileName(currentFilePath) || 'Workbook';

    const meta = document.createElement('div');
    meta.className = 'spreadsheet-meta';
    const sheetCount = workbook.sheets?.length || 0;
    const rowCount = sheet?.row_count || 0;
    const columnCount = sheet?.column_count || 0;
    meta.textContent = `${sheetCount} sheet${sheetCount === 1 ? '' : 's'} · ${rowCount} row${rowCount === 1 ? '' : 's'} · ${columnCount} column${columnCount === 1 ? '' : 's'}`;

    header.appendChild(title);
    header.appendChild(meta);

    if (sheet?.truncated && sheet.truncated_reasons?.length) {
        const notice = document.createElement('div');
        notice.className = 'spreadsheet-truncation';
        notice.textContent = sheet.truncated_reasons.join(' ');
        header.appendChild(notice);
    }

    return header;
}

function renderSheetTabs(workbook) {
    const tabs = document.createElement('div');
    tabs.className = 'sheet-tabs';

    for (const sheet of workbook.sheets || []) {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'sheet-tab';
        if (sheet.index === workbook.active_sheet_index) {
            button.classList.add('active');
        }
        if (!sheet.visible) {
            button.classList.add('hidden-sheet');
        }
        button.dataset.sheetIndex = String(sheet.index);
        button.textContent = sheet.name;
        button.addEventListener('click', () => {
            if (sheet.index !== currentWorkbook?.active_sheet_index) {
                void loadWorkbookSheet(sheet.index);
            }
        });
        tabs.appendChild(button);
    }

    return tabs;
}

async function loadWorkbookSheet(sheetIndex) {
    if (!currentWorkbook) return;

    // Serve cached sheets instantly — no backend round-trip or re-parse.
    const cache = currentWorkbook._sheetCache || (currentWorkbook._sheetCache = new Map());
    if (cache.has(sheetIndex)) {
        currentWorkbook.active_sheet_index = sheetIndex;
        currentWorkbook.active_sheet = cache.get(sheetIndex);
        renderWorkbook(currentWorkbook);
        return;
    }

    if (!invoke || !currentFilePath) return;

    showStatus('Loading sheet...', 'loading', 1500);

    try {
        const sheet = await invoke('open_xlsx_sheet', { path: currentFilePath, sheetIndex });
        cache.set(sheetIndex, sheet);
        currentWorkbook.active_sheet_index = sheet.index;
        currentWorkbook.active_sheet = sheet;
        renderWorkbook(currentWorkbook);
        showStatus(sheet.name + ' loaded', 'success', 1200);
    } catch (err) {
        showError(err);
    }
}

const ROW_HEADER_WIDTH = 54;

function renderSheetGrid(sheet, limits = {}) {
    if (!sheet.rows || sheet.rows.length === 0) {
        return renderEmptySheetState();
    }

    const maxColumns = limits.max_columns || 200;
    const maxRows = limits.max_rows || 2000;
    const images = sheet.images || [];

    // Find the data extent with plain loops; spreading a large array into
    // Math.max can overflow the call stack on big sheets.
    let observedColumns = 1;
    let observedRows = 1;
    for (const row of sheet.rows) {
        if (row.index > observedRows) observedRows = row.index;
        const cells = row.cells;
        if (cells && cells.length) {
            // cells are in ascending column order, so the last one is the widest.
            const lastColumn = cells[cells.length - 1].column || 1;
            if (lastColumn > observedColumns) observedColumns = lastColumn;
        }
    }

    // Extend the visible range so picture anchors always land on real cells.
    const imageColumns = images.length
        ? Math.max(0, ...images.map(image => (image.to_col != null ? image.to_col : image.from_col) + 1))
        : 0;
    const displayColumnCount = Math.max(1, Math.min(Math.max(observedColumns, imageColumns), maxColumns));

    const imageRows = images.length
        ? Math.max(0, ...images.map(image => (image.to_row != null ? image.to_row : image.from_row) + 1))
        : 0;
    const lastRow = Math.max(observedRows, imageRows, 1);
    const maxGridCells = 60000;
    const maxRowsByCells = Math.max(1, Math.floor(maxGridCells / displayColumnCount));
    const displayRowCount = Math.min(lastRow, maxRowsByCells, maxRows);

    // Translate the workbook's own geometry into pixels so the grid mirrors
    // Excel — this is what makes anchored images line up with empty cells.
    const defaultColChars = sheet.default_col_width || 8.43;
    const defaultRowPts = sheet.default_row_height || 15;
    const columnDefs = sheet.columns || [];
    const columnWidthPx = (column) => {
        const definition = columnDefs.find(col => column >= col.min && column <= col.max);
        if (definition) {
            return definition.hidden ? 0 : charsToPixels(definition.width);
        }
        return charsToPixels(defaultColChars);
    };

    const rowsByIndex = new Map(sheet.rows.map(row => [row.index, row]));
    const rowHeightPx = (row) => pointsToPixels(row && row.height != null ? row.height : defaultRowPts);

    const wrapper = document.createElement('div');
    wrapper.className = 'spreadsheet-grid-wrap';

    if (displayRowCount < lastRow) {
        const notice = document.createElement('div');
        notice.className = 'spreadsheet-truncation';
        notice.textContent = `Showing ${displayRowCount} of ${lastRow} rows to keep rendering fast.`;
        wrapper.appendChild(notice);
    }

    const gridLayer = document.createElement('div');
    gridLayer.className = 'spreadsheet-grid-layer';

    const table = document.createElement('table');
    table.className = 'spreadsheet-grid spreadsheet-grid--sized';

    // Fixed column widths matching the workbook (table-layout: fixed honours these).
    const colgroup = document.createElement('colgroup');
    const headerCol = document.createElement('col');
    headerCol.style.width = `${ROW_HEADER_WIDTH}px`;
    colgroup.appendChild(headerCol);
    const colElements = [];
    for (let column = 1; column <= displayColumnCount; column++) {
        const col = document.createElement('col');
        col.style.width = `${columnWidthPx(column)}px`;
        colgroup.appendChild(col);
        colElements[column] = col;
    }
    table.appendChild(colgroup);

    const recomputeTableWidth = () => {
        let total = ROW_HEADER_WIDTH;
        for (let column = 1; column <= displayColumnCount; column++) {
            total += parseFloat(colElements[column].style.width) || 0;
        }
        table.style.width = `${total}px`;
    };
    recomputeTableWidth();

    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const corner = document.createElement('th');
    corner.className = 'sheet-corner';
    headerRow.appendChild(corner);
    for (let column = 1; column <= displayColumnCount; column++) {
        const th = document.createElement('th');
        th.scope = 'col';
        th.textContent = columnName(column);
        th.appendChild(makeColumnResizer(colElements[column], recomputeTableWidth, () => repositionImages()));
        headerRow.appendChild(th);
    }
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const defaultRowPx = pointsToPixels(defaultRowPts);
    // A default-height row we can measure once: the browser renders rows at
    // max(specified height, content height), and cell text usually forces a
    // taller min height than the nominal row height. Measuring one default row
    // captures that floor so image offsets don't drift down the sheet.
    let measureRowEl = null;

    const tbody = document.createElement('tbody');
    for (let rowIndex = 1; rowIndex <= displayRowCount; rowIndex++) {
        const row = rowsByIndex.get(rowIndex);
        const tr = document.createElement('tr');
        tr.style.height = `${rowHeightPx(row)}px`;
        if (!measureRowEl && !(row && row.height != null)) {
            measureRowEl = tr;
        }

        const rowHeader = document.createElement('th');
        rowHeader.scope = 'row';
        rowHeader.textContent = rowIndex;
        tr.appendChild(rowHeader);

        const cells = (row && row.cells) || [];
        let cellPointer = 0;
        for (let column = 1; column <= displayColumnCount; column++) {
            const td = document.createElement('td');
            // cells are already in column order, so advance a pointer instead of
            // building a per-row map.
            const cell = cellPointer < cells.length && cells[cellPointer].column === column
                ? cells[cellPointer++]
                : null;
            if (cell) {
                td.textContent = cell.value || '';
                if (cell.value_type !== 'string') {
                    td.className = 'xlsx-cell-' + cell.value_type;
                }
                if (cell.formula) {
                    td.title = '=' + cell.formula;
                    td.classList.add('xlsx-cell-formula');
                }
            }
            tr.appendChild(td);
        }

        tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    gridLayer.appendChild(table);

    let imagesLayer = null;
    if (images.length) {
        imagesLayer = document.createElement('div');
        imagesLayer.className = 'spreadsheet-grid-images';
        gridLayer.appendChild(imagesLayer);
    }

    // The header height and the rendered row-height floor are the only values
    // that need measuring (they depend on font metrics); everything else is
    // computed arithmetically. Both are cached after the first measurement.
    let headerHeight = null;
    let rowTops = null; // prefix sums of row tops within the body
    let bodyHeight = 0;
    let extrapolationRowPx = defaultRowPx;

    const measureGeometry = () => {
        headerHeight = table.tHead ? table.tHead.offsetHeight : defaultRowPx;
        // Actual rendered height of a default row = max(nominal, content floor).
        const contentMin = measureRowEl ? measureRowEl.offsetHeight : defaultRowPx;
        extrapolationRowPx = Math.max(defaultRowPx, contentMin);
        rowTops = new Array(displayRowCount + 2);
        let acc = 0;
        for (let rowIndex = 1; rowIndex <= displayRowCount; rowIndex++) {
            rowTops[rowIndex] = acc;
            const row = rowsByIndex.get(rowIndex);
            const nominal = row && row.height != null ? pointsToPixels(row.height) : defaultRowPx;
            acc += Math.max(nominal, contentMin);
        }
        rowTops[displayRowCount + 1] = acc;
        bodyHeight = acc;
    };

    const placeImages = () => {
        if (!imagesLayer) return;
        if (rowTops == null) {
            measureGeometry();
        }

        // Prefix sums of column left edges (recomputed since resizing changes them).
        const colLefts = new Array(displayColumnCount + 2);
        colLefts[1] = ROW_HEADER_WIDTH;
        for (let column = 1; column <= displayColumnCount; column++) {
            colLefts[column + 1] = colLefts[column] + (parseFloat(colElements[column].style.width) || 0);
        }
        const lastColWidth = parseFloat(colElements[displayColumnCount].style.width) || 0;
        const colLeftAt = (column) => column <= displayColumnCount + 1
            ? colLefts[column]
            : colLefts[displayColumnCount + 1] + lastColWidth * (column - displayColumnCount - 1);
        const rowTopAt = (rowIndex) => rowIndex <= displayRowCount + 1
            ? rowTops[rowIndex]
            : bodyHeight + extrapolationRowPx * (rowIndex - displayRowCount - 1);

        const fragment = document.createDocumentFragment();
        for (const image of images) {
            if (!image.data_uri) continue;

            const left = colLeftAt(image.from_col + 1) + (image.from_col_off || 0) / EMU_PER_PIXEL;
            const top = headerHeight + rowTopAt(image.from_row + 1) + (image.from_row_off || 0) / EMU_PER_PIXEL;

            let width;
            let height;
            if (image.anchor_type === 'two' && image.to_col != null && image.to_row != null) {
                const right = colLeftAt(image.to_col + 1) + (image.to_col_off || 0) / EMU_PER_PIXEL;
                const bottom = headerHeight + rowTopAt(image.to_row + 1) + (image.to_row_off || 0) / EMU_PER_PIXEL;
                width = Math.max(0, right - left);
                height = Math.max(0, bottom - top);
            } else if (image.ext_cx && image.ext_cy) {
                width = image.ext_cx / EMU_PER_PIXEL;
                height = image.ext_cy / EMU_PER_PIXEL;
            }

            const img = document.createElement('img');
            img.className = 'spreadsheet-image';
            img.src = image.data_uri;
            img.alt = '';
            img.loading = 'lazy';
            img.decoding = 'async';
            img.style.left = `${left}px`;
            img.style.top = `${top}px`;
            if (width) img.style.width = `${width}px`;
            if (height) img.style.height = `${height}px`;
            fragment.appendChild(img);
        }
        imagesLayer.replaceChildren(fragment);
    };

    // Defer once so the header has been laid out, then place.
    if (imagesLayer) {
        requestAnimationFrame(placeImages);
    }
    // Re-run image placement after geometry changes (e.g. column resize).
    repositionImages = placeImages;

    wrapper.appendChild(gridLayer);

    return wrapper;
}

// Set per-render by renderSheetGrid; lets the resize handler re-place images.
let repositionImages = () => {};

// Builds a drag handle on a column header that resizes the matching <col>.
function makeColumnResizer(col, onResize, onResizeEnd) {
    const MIN_WIDTH = 24;
    const handle = document.createElement('div');
    handle.className = 'col-resizer';

    let startX = 0;
    let startWidth = 0;

    const onMove = (event) => {
        const next = Math.max(MIN_WIDTH, startWidth + (event.clientX - startX));
        col.style.width = `${next}px`;
        onResize();
    };
    const onUp = (event) => {
        handle.releasePointerCapture(event.pointerId);
        handle.removeEventListener('pointermove', onMove);
        handle.removeEventListener('pointerup', onUp);
        document.body.classList.remove('col-resizing');
        onResizeEnd();
    };

    handle.addEventListener('pointerdown', (event) => {
        event.preventDefault();
        event.stopPropagation();
        startX = event.clientX;
        startWidth = parseFloat(col.style.width) || 0;
        handle.setPointerCapture(event.pointerId);
        handle.addEventListener('pointermove', onMove);
        handle.addEventListener('pointerup', onUp);
        document.body.classList.add('col-resizing');
    });

    return handle;
}

// Excel stores column widths in "characters of the maximum digit width"; the
// Calibri 11 digit is ~7px and cells carry ~5px of padding.
function charsToPixels(chars) {
    return Math.round(chars * 7) + 5;
}

// Row heights are stored in points; convert to CSS pixels at 96 DPI.
function pointsToPixels(points) {
    return Math.round(points * 96 / 72);
}

const EMU_PER_PIXEL = 9525;

function renderEmptySheetState() {
    const empty = document.createElement('div');
    empty.className = 'spreadsheet-empty';
    empty.innerHTML = `
        <h2>This sheet is empty</h2>
        <p>Papyr opened the workbook, but this sheet has no visible preview cells.</p>
    `;
    return empty;
}

// --- Document rendering ---

function renderDocument(doc) {
    currentDocument = doc;
    currentWorkbook = null;
    welcomeScreen.style.display = 'none';
    documentView.style.display = 'flex';
    deskContent.replaceChildren();
    deskContent.classList.remove('spreadsheet-content');
    desk.classList.remove('desk--spreadsheet');
    setDocumentZoomWheel(true);
    setCommentsAvailability(true);
    setCommentsVisibility(false);
    closeFindBar();

    // Cache resolved default style once for entire render pass
    doc._defaultStyle = computeDefaultStyle(doc);

    const pages = splitIntoPages(doc.body);
    const fragment = document.createDocumentFragment();

    if (!doc.body || doc.body.length === 0 || pages.every(page => page.length === 0)) {
        fragment.appendChild(renderEmptyDocumentPage());
    } else {
        for (const pageBlocks of pages) {
            fragment.appendChild(renderPage(pageBlocks, doc));
        }
    }

    deskContent.appendChild(fragment);
    updateDocumentImages();
    renderComments(doc.comments);

    if (doc.comments && doc.comments.length > 0) {
        setCommentsVisibility(true);
    }

    const filename = getFileName(currentFilePath) || 'Papyr';
    document.title = filename + ' - Papyr';

    if (findInput.value.trim()) {
        scheduleFind();
    } else {
        clearFindHighlights();
        findMatches = [];
        findIndex = -1;
        findCount.textContent = '';
    }
}

function renderPage(pageBlocks, doc) {
    const page = document.createElement('div');
    page.className = 'page';

    if (doc.headers && doc.headers.length > 0) {
        const header = document.createElement('div');
        header.className = 'page-header';
        renderBlocks(doc.headers[0].content, header, doc);
        page.appendChild(header);
    }

    const content = document.createElement('div');
    content.className = 'page-content';
    renderBlocks(pageBlocks, content, doc);
    page.appendChild(content);

    const pageFootnotes = collectFootnotes(pageBlocks, doc.footnotes);
    if (pageFootnotes.length > 0) {
        const fnSection = document.createElement('div');
        fnSection.className = 'page-footnotes';
        fnSection.appendChild(document.createElement('hr'));
        const fnFragment = document.createDocumentFragment();
        pageFootnotes.forEach(fn => {
            const fnEl = document.createElement('div');
            fnEl.className = 'footnote';
            fnEl.id = 'footnote-' + fn.id;
            const marker = document.createElement('sup');
            marker.textContent = fn.id;
            fnEl.appendChild(marker);
            const fnContent = document.createElement('span');
            renderBlocks(fn.content, fnContent, doc);
            fnEl.appendChild(fnContent);
            fnFragment.appendChild(fnEl);
        });
        fnSection.appendChild(fnFragment);
        page.appendChild(fnSection);
    }

    if (doc.footers && doc.footers.length > 0) {
        const footer = document.createElement('div');
        footer.className = 'page-footer';
        renderBlocks(doc.footers[0].content, footer, doc);
        page.appendChild(footer);
    }

    return page;
}

function renderEmptyDocumentPage() {
    const page = document.createElement('div');
    page.className = 'page';

    const content = document.createElement('div');
    content.className = 'page-content empty-document';
    content.innerHTML = `
        <div class="doc-empty-state">
            <h2>This document is empty</h2>
            <p>Papyr opened the file, but there is no visible body content to render yet.</p>
        </div>
    `;
    page.appendChild(content);
    return page;
}

function splitIntoPages(blocks) {
    const pages = [];
    let current = [];

    for (const block of blocks) {
        if (block.type === 'page_break') {
            pages.push(current);
            current = [];
        } else {
            current.push(block);
        }
    }
    if (current.length > 0) {
        pages.push(current);
    }
    if (pages.length === 0) {
        pages.push([]);
    }
    return pages;
}

function renderBlocks(blocks, container, doc) {
    if (!blocks) return;

    const fragment = document.createDocumentFragment();
    for (const block of blocks) {
        switch (block.type) {
            case 'paragraph':
                fragment.appendChild(renderParagraph(block, doc));
                break;
            case 'table':
                fragment.appendChild(renderTable(block, doc));
                break;
            case 'page_break':
                break;
        }
    }
    container.appendChild(fragment);
}

function renderParagraph(para, doc) {
    const style = resolveParagraphStyle(para, doc);
    const headingLevel = style?.heading_level || detectHeadingLevel(para.style);

    let el;
    if (headingLevel && headingLevel >= 1 && headingLevel <= 6) {
        el = document.createElement('h' + headingLevel);
    } else {
        el = document.createElement('p');
    }

    applyParagraphStyle(el, para, style);

    // List item handling
    if (para.list_level != null) {
        el.classList.add('doc-list-item');
        el.classList.add('doc-list-level-' + para.list_level);
        el.style.paddingLeft = ((para.list_level + 1) * 1.5) + 'em';
        if (para.list_format === 'bullet') {
            el.classList.add('doc-list-bullet');
        } else if (para.list_format) {
            el.classList.add('doc-list-ordered');
        }
    }

    const fragment = document.createDocumentFragment();
    if (para.runs) {
        for (const run of para.runs) {
            renderRun(run, fragment, doc);
        }
    }

    if (fragment.childNodes.length === 0 && para.list_level == null) {
        el.innerHTML = '&nbsp;';
    } else {
        el.appendChild(fragment);
    }

    return el;
}

function applyParagraphStyle(el, para, style) {
    const align = para.alignment || style?.alignment;
    if (align) {
        el.style.textAlign = align;
    }
    if (style?.font_size) {
        el.style.fontSize = `calc(${style.font_size}pt * var(--doc-zoom, 1))`;
    }
    if (style?.bold === true) {
        el.style.fontWeight = 'bold';
    } else if (style?.bold === false) {
        el.style.fontWeight = 'normal';
    }
    if (style?.italic === true) {
        el.style.fontStyle = 'italic';
    } else if (style?.italic === false) {
        el.style.fontStyle = 'normal';
    }
    if (style?.color && style.color !== 'auto') {
        el.style.color = '#' + style.color;
    }
}

function resolveParagraphStyle(para, doc) {
    // Styles are already inheritance-resolved in Rust, so direct lookup is sufficient
    if (para?.style && doc?.styles) {
        return doc.styles[para.style] || doc._defaultStyle || null;
    }
    return doc._defaultStyle || null;
}

function computeDefaultStyle(doc) {
    if (!doc?.styles) return null;
    if (doc.styles['Normal']) return doc.styles['Normal'];
    if (doc.styles['normal']) return doc.styles['normal'];
    const key = Object.keys(doc.styles).find(k => k.toLowerCase() === 'normal');
    return key ? doc.styles[key] : null;
}

function renderRun(run, container, doc) {
    // Image run
    if (run.image_id && doc.images && doc.images[run.image_id]) {
        const img = document.createElement('img');
        img.src = doc.images[run.image_id];
        img.className = 'doc-image';
        img.alt = 'Document image';
        img.addEventListener('load', () => updateDocumentImageSize(img), { once: true });
        container.appendChild(img);
        if (img.complete) {
            updateDocumentImageSize(img);
        }
        return;
    }

    // Footnote reference
    if (run.footnote_ref) {
        const sup = document.createElement('sup');
        const link = document.createElement('a');
        link.href = '#footnote-' + run.footnote_ref;
        link.className = 'footnote-ref';
        link.textContent = run.footnote_ref;
        sup.appendChild(link);
        container.appendChild(sup);
    }

    if (!run.text || run.text.length === 0) return;

    const hasFormatting = run.bold || run.italic || run.underline || run.strikethrough ||
        run.font_size || run.color || run.highlight || run.comment_ref != null;

    // Plain text node shortcut
    if (!hasFormatting && !run.text.includes('\t') && !run.link_url) {
        container.appendChild(document.createTextNode(run.text));
        return;
    }

    // Determine output target (link wrapper or direct container)
    let target = container;
    if (run.link_url) {
        const a = document.createElement('a');
        a.href = run.link_url;
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
        a.className = 'doc-link';
        target = a;
    }

    // Render text, splitting on tabs into explicit tab-stop elements
    if (run.text.includes('\t')) {
        const parts = run.text.split('\t');
        for (let i = 0; i < parts.length; i++) {
            if (parts[i]) {
                target.appendChild(hasFormatting ? styledSpan(parts[i], run) : document.createTextNode(parts[i]));
            }
            if (i < parts.length - 1) {
                const tab = document.createElement('span');
                tab.className = 'doc-tab';
                target.appendChild(tab);
            }
        }
    } else {
        target.appendChild(styledSpan(run.text, run));
    }

    if (target !== container) {
        container.appendChild(target);
    }
}

function styledSpan(text, run) {
    const span = document.createElement('span');
    span.textContent = text;
    if (run.bold) span.style.fontWeight = 'bold';
    if (run.italic) span.style.fontStyle = 'italic';
    if (run.underline) span.style.textDecoration = 'underline';
    if (run.strikethrough) {
        span.style.textDecoration = (span.style.textDecoration || '') +
            (span.style.textDecoration ? ' line-through' : 'line-through');
    }
    if (run.font_size) span.style.fontSize = `calc(${run.font_size}pt * var(--doc-zoom, 1))`;
    if (run.color && run.color !== 'auto') span.style.color = '#' + run.color;
    if (run.highlight) {
        span.style.backgroundColor = highlightColorMap[run.highlight] || run.highlight;
    }
    if (run.comment_ref != null) {
        span.classList.add('commented-text');
        span.dataset.commentId = run.comment_ref;
    }
    return span;
}

function updateDocumentImages() {
    if (!deskContent) return;

    const images = deskContent.querySelectorAll('.doc-image');
    images.forEach((img) => updateDocumentImageSize(img));
}

function updateDocumentImageSize(img) {
    const intrinsicWidth = Number.parseFloat(img.dataset.intrinsicWidth) || img.naturalWidth;
    if (!Number.isFinite(intrinsicWidth) || intrinsicWidth <= 0) return;

    img.dataset.intrinsicWidth = String(intrinsicWidth);
    img.style.width = (Math.round(intrinsicWidth * currentDocumentZoom * 100) / 100) + 'px';
}

function renderTable(table, doc) {
    const tableEl = document.createElement('table');
    tableEl.className = 'doc-table';

    if (table.rows) {
        const fragment = document.createDocumentFragment();
        for (const row of table.rows) {
            const tr = document.createElement('tr');
            if (row.cells) {
                for (const cell of row.cells) {
                    const td = document.createElement('td');
                    if (cell.col_span > 1) td.colSpan = cell.col_span;
                    if (cell.row_span > 1) td.rowSpan = cell.row_span;
                    if (cell.shading) td.style.backgroundColor = cell.shading;
                    renderBlocks(cell.content, td, doc);
                    tr.appendChild(td);
                }
            }
            fragment.appendChild(tr);
        }
        tableEl.appendChild(fragment);
    }

    return tableEl;
}

function collectFootnotes(blocks, footnotes) {
    if (!footnotes || !blocks) return [];
    const refs = new Set();
    const collectRefs = (blocks) => {
        for (const b of blocks) {
            if (b.type === 'paragraph' && b.runs) {
                for (const r of b.runs) {
                    if (r.footnote_ref) refs.add(r.footnote_ref);
                }
            }
            if (b.type === 'table' && b.rows) {
                for (const row of b.rows) {
                    for (const cell of row.cells || []) {
                        collectRefs(cell.content || []);
                    }
                }
            }
        }
    };
    collectRefs(blocks);
    return footnotes.filter(fn => refs.has(fn.id));
}

function detectHeadingLevel(styleName) {
    if (!styleName) return 0;
    const lower = styleName.toLowerCase();
    if (lower === 'title') return 1;
    if (lower === 'subtitle') return 2;
    if (!lower.startsWith('heading')) return 0;
    const ch = lower.charAt(lower.length - 1);
    return (ch >= '1' && ch <= '6') ? (ch.charCodeAt(0) - 48) : 0;
}

const highlightColorMap = {
    yellow: '#ffff00',
    green: '#00ff00',
    cyan: '#00ffff',
    magenta: '#ff00ff',
    blue: '#0000ff',
    red: '#ff0000',
    darkBlue: '#00008b',
    darkCyan: '#008b8b',
    darkGreen: '#006400',
    darkMagenta: '#8b008b',
    darkRed: '#8b0000',
    darkYellow: '#9b870c',
    darkGray: '#a9a9a9',
    lightGray: '#d3d3d3',
    black: '#000000',
};

// --- Comments ---

function renderComments(comments) {
    commentsList.replaceChildren();
    if (!comments || comments.length === 0) {
        commentsList.innerHTML = '<p class="no-comments">No comments in this document.</p>';
        return;
    }

    const fragment = document.createDocumentFragment();
    for (const comment of comments) {
        const el = document.createElement('div');
        el.className = 'comment-card';
        el.id = 'comment-' + comment.id;
        el.dataset.commentId = comment.id;
        el.innerHTML = `
            <div class="comment-meta">
                <strong>${escapeHtml(comment.author)}</strong>
                ${comment.date ? '<span class="comment-date">' + formatDate(comment.date) + '</span>' : ''}
            </div>
            <div class="comment-text">${escapeHtml(comment.text)}</div>
        `;
        fragment.appendChild(el);
    }
    commentsList.appendChild(fragment);
}

function toggleComments() {
    if (commentsButton.disabled) return;
    setCommentsVisibility(!commentsVisible);
}

function setCommentsAvailability(isAvailable) {
    commentsButton.disabled = !isAvailable;
}

function setCommentsVisibility(isVisible) {
    commentsVisible = Boolean(isVisible);
    commentsPanel.style.display = commentsVisible ? 'flex' : 'none';
    commentsButton.classList.toggle('is-active', commentsVisible);
    commentsButton.setAttribute('aria-pressed', String(commentsVisible));
}

function scrollToComment(commentId) {
    const el = document.getElementById('comment-' + commentId);
    if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.classList.add('highlight');
        setTimeout(() => el.classList.remove('highlight'), 2000);
    }
    if (!commentsVisible) toggleComments();
}

function scrollToCommentedText(commentId) {
    const el = document.querySelector(`[data-comment-id="${commentId}"]`);
    if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.classList.add('flash');
        setTimeout(() => el.classList.remove('flash'), 2000);
    }
}

function handleDeskClick(event) {
    const commentedText = event.target.closest('.commented-text');
    if (!commentedText) return;
    const commentId = commentedText.dataset.commentId;
    if (commentId != null) {
        scrollToComment(commentId);
    }
}

function handleCommentListClick(event) {
    const card = event.target.closest('.comment-card');
    if (!card) return;
    const commentId = card.id?.replace('comment-', '');
    if (commentId) {
        scrollToCommentedText(commentId);
    }
}

// --- Find ---

function openFindBar() {
    findBar.style.display = 'flex';
    findInput.focus();
    findInput.select();
}

function closeFindBar(resetHighlights = true) {
    findBar.style.display = 'none';
    if (resetHighlights) {
        clearFindHighlights();
        findMatches = [];
        findIndex = -1;
        findCount.textContent = '';
        findInput.value = '';
    }
    if (findDebounceTimer) {
        clearTimeout(findDebounceTimer);
        findDebounceTimer = null;
    }
}

function scheduleFind() {
    if (findDebounceTimer) {
        clearTimeout(findDebounceTimer);
    }
    findDebounceTimer = setTimeout(performFind, 120);
}

function performFind() {
    clearFindHighlights();
    const query = findInput.value.trim().toLowerCase();
    if (!query) {
        findCount.textContent = '';
        findMatches = [];
        findIndex = -1;
        return;
    }

    findMatches = [];
    const textNodes = [];
    const walker = document.createTreeWalker(deskContent, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
        textNodes.push(walker.currentNode);
    }

    for (const node of textNodes) {
        const text = node.textContent;
        const lowerText = text.toLowerCase();
        let start = 0;
        let matchIndex = lowerText.indexOf(query, start);
        if (matchIndex === -1) continue;

        const fragment = document.createDocumentFragment();
        while (matchIndex !== -1) {
            if (matchIndex > start) {
                fragment.appendChild(document.createTextNode(text.slice(start, matchIndex)));
            }
            const mark = document.createElement('mark');
            mark.className = 'find-highlight';
            mark.textContent = text.slice(matchIndex, matchIndex + query.length);
            fragment.appendChild(mark);
            findMatches.push(mark);
            start = matchIndex + query.length;
            matchIndex = lowerText.indexOf(query, start);
        }
        if (start < text.length) {
            fragment.appendChild(document.createTextNode(text.slice(start)));
        }
        node.parentNode.replaceChild(fragment, node);
    }

    findCount.textContent = findMatches.length + ' found';
    findIndex = findMatches.length > 0 ? 0 : -1;
    if (findIndex >= 0) highlightCurrentMatch();
}

function navigateFind(dir) {
    if (findMatches.length === 0) return;
    findIndex = (findIndex + dir + findMatches.length) % findMatches.length;
    highlightCurrentMatch();
}

function highlightCurrentMatch() {
    findMatches.forEach(m => m.classList.remove('current'));
    if (findIndex >= 0 && findIndex < findMatches.length) {
        findMatches[findIndex].classList.add('current');
        findMatches[findIndex].scrollIntoView({ behavior: 'smooth', block: 'center' });
        findCount.textContent = (findIndex + 1) + '/' + findMatches.length;
    }
}

function clearFindHighlights() {
    if (findDebounceTimer) {
        clearTimeout(findDebounceTimer);
        findDebounceTimer = null;
    }
    const marks = findMatches.length > 0 ? findMatches : Array.from(document.querySelectorAll('mark.find-highlight'));
    marks.forEach(mark => {
        const parent = mark.parentNode;
        if (!parent) return;
        parent.replaceChild(document.createTextNode(mark.textContent), mark);
        parent.normalize();
    });
    findMatches = [];
}

// --- Theme ---

async function initializeTheme() {
    const localTheme = normalizeTheme(localStorage.getItem('papyr-theme')) || 'light';
    applyTheme(localTheme);
    localStorage.setItem('papyr-theme', localTheme);

    if (!invoke) return;

    try {
        const savedTheme = normalizeTheme(await invoke('get_theme_preference'));
        if (!savedTheme) return;

        applyTheme(savedTheme);
        localStorage.setItem('papyr-theme', savedTheme);
    } catch (err) {
        console.log('Could not load theme preference:', err);
    }
}

function hasTauriApi() {
    return Boolean(invoke && open && listen);
}

function reportMissingTauriApi() {
    const message = 'Papyr failed to load its Tauri desktop APIs. Rebuild the app after enabling withGlobalTauri in tauri.conf.json.';
    if (fileInfo) {
        fileInfo.textContent = message;
        fileInfo.style.color = '#e74c3c';
    }
    showStatus(message, 'error', null);
}

function toggleTheme() {
    const current = document.body.classList.contains('dark-theme') ? 'dark' : 'light';
    void setTheme(current === 'light' ? 'dark' : 'light');
}

async function setTheme(theme) {
    const normalizedTheme = normalizeTheme(theme) || 'light';
    applyTheme(normalizedTheme);
    localStorage.setItem('papyr-theme', normalizedTheme);

    if (!invoke) return;

    try {
        await invoke('set_theme_preference', { theme: normalizedTheme });
    } catch (err) {
        console.log('Could not save theme preference:', err);
    }
}

function applyTheme(theme) {
    if (theme === 'dark') {
        document.documentElement.classList.add('dark-theme');
        document.body.classList.add('dark-theme');
        themeIcon.innerHTML = getThemeIconSvg('dark');
    } else {
        document.documentElement.classList.remove('dark-theme');
        document.body.classList.remove('dark-theme');
        themeIcon.innerHTML = getThemeIconSvg('light');
    }
}

function normalizeTheme(theme) {
    return theme === 'dark' || theme === 'light' ? theme : null;
}

function getThemeIconSvg(theme) {
    if (theme === 'dark') {
        return `
            <svg viewBox="0 0 24 24">
                <path class="icon-fill" d="M14.75 3.25a8.25 8.25 0 1 0 5.9 12.75 9 9 0 0 1-5.9-12.75Z"></path>
                <path d="M14.75 3.25a8.25 8.25 0 1 0 5.9 12.75 9 9 0 0 1-5.9-12.75Z"></path>
                <path d="m18.6 4 .38 1.02L20 5.4l-1.02.38-.38 1.02-.38-1.02-1.02-.38 1.02-.38L18.6 4Z"></path>
            </svg>
        `;
    }

    return `
        <svg viewBox="0 0 24 24">
            <circle class="icon-fill" cx="12" cy="12" r="4.25"></circle>
            <circle cx="12" cy="12" r="4.25"></circle>
            <path d="M12 2.5v2.25M12 19.25v2.25M21.5 12h-2.25M4.75 12H2.5"></path>
            <path d="m18.72 5.28-1.6 1.6M6.88 17.12l-1.6 1.6M18.72 18.72l-1.6-1.6M6.88 6.88l-1.6-1.6"></path>
        </svg>
    `;
}

// --- Utilities ---

function escapeHtml(text) {
    if (!text) return '';
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function formatDate(dateStr) {
    try {
        return new Date(dateStr).toLocaleDateString(undefined, {
            year: 'numeric', month: 'short', day: 'numeric'
        });
    } catch { return dateStr; }
}

function getFileName(path) {
    return path ? path.split(/[\\/]/).pop() : '';
}

function isSupportedFilePath(path) {
    if (typeof path !== 'string') return false;
    const lower = path.toLowerCase();
    return lower.endsWith('.docx') || SUPPORTED_SPREADSHEET_EXTENSIONS.some((ext) => lower.endsWith('.' + ext));
}

function columnName(index) {
    let name = '';
    let current = index;
    while (current > 0) {
        current -= 1;
        name = String.fromCharCode(65 + (current % 26)) + name;
        current = Math.floor(current / 26);
    }
    return name;
}
