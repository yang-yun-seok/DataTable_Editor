const STORAGE_KEY = "dt-editor.project.v2";
const HISTORY_LIMIT = 60;
const TYPE_META = {
    INT: { label: "INT", editor: "text", placeholder: "0" },
    FLOAT: { label: "FLOAT", editor: "text", placeholder: "0.0" },
    STRING: { label: "STRING", editor: "text", placeholder: "text" },
    BOOLEAN: { label: "BOOLEAN", editor: "boolean", placeholder: "true / false" },
    ENUM: { label: "ENUM", editor: "enum", placeholder: "ACTIVE" },
    DATE: { label: "DATE", editor: "date", placeholder: "2026-05-21" },
    JSON: { label: "JSON", editor: "json", placeholder: "{\"key\":\"value\"}" },
};
const LEGACY_TYPE_MAP = {
    ARRAY: "JSON",
    VECTOR3: "STRING",
};
const TYPE_KEYS = Object.keys(TYPE_META);
const CARDINALITY_OPTIONS = ["N:1", "1:1", "1:N", "N:N"];
const DEFAULT_HELPERS = [
    { value: "", label: "Custom" },
    { value: "{{today}}", label: "Today" },
    { value: "{{now}}", label: "Timestamp" },
    { value: "{{uuid}}", label: "UUID" },
    { value: "{{autoincrement}}", label: "Auto increment" },
    { value: "{{true}}", label: "TRUE" },
    { value: "{{false}}", label: "FALSE" },
];

const dom = {
    sidebar: document.getElementById("sidebar"),
    topbar: document.getElementById("topbar"),
    workspace: document.getElementById("workspace"),
    jsonInput: document.getElementById("import-json-file"),
    csvInput: document.getElementById("import-csv-file"),
    excelInput: document.getElementById("import-excel-file"),
    toastHost: document.getElementById("toast-host"),
};

const state = {
    project: null,
    validation: null,
    history: {
        undo: [],
        redo: [],
    },
    ui: {
        view: { kind: "schema", tableId: null },
        dataSearch: {},
        dataSort: {},
        bulkPasteDraft: {},
        bulkPasteOpen: {},
        selectedRows: {},
        bulkEditColumn: {},
        bulkEditValue: {},
        sqlImportOpen: false,
        sqlImportDraft: "",
        scrollPositions: {},
        focus: null,
        dragColumn: null,
        erd: { zoom: 0.95, panX: 140, panY: 90 },
        erdDrag: null,
        erdPan: null,
    },
    saveStatus: {
        source: "demo",
        lastSavedAt: null,
    },
};

function uid(prefix) {
    if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
        return `${prefix}_${crypto.randomUUID().slice(0, 8)}`;
    }
    return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
}

function clone(value) {
    if (typeof structuredClone === "function") {
        return structuredClone(value);
    }
    return JSON.parse(JSON.stringify(value));
}

function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, (char) => {
        switch (char) {
            case "&":
                return "&amp;";
            case "<":
                return "&lt;";
            case ">":
                return "&gt;";
            case '"':
                return "&quot;";
            case "'":
                return "&#39;";
            default:
                return char;
        }
    });
}

function escapeAttr(value) {
    return escapeHtml(value);
}

function toCellString(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "string") return value;
    if (typeof value === "number" || typeof value === "boolean") return String(value);
    try {
        return JSON.stringify(value);
    } catch (_error) {
        return String(value);
    }
}

function normalizeEnumValues(value) {
    const source = Array.isArray(value)
        ? value
        : String(value ?? "")
              .split(/[\n,]/)
              .map((item) => item.trim());
    return Array.from(new Set(source.map((item) => String(item).trim()).filter(Boolean)));
}

function getDefaultHelperSelection(defaultValue) {
    const raw = String(defaultValue ?? "").trim();
    return DEFAULT_HELPERS.some((helper) => helper.value === raw) ? raw : "";
}

function normalizeType(type) {
    const key = String(type ?? "STRING").toUpperCase();
    const mapped = LEGACY_TYPE_MAP[key] || key;
    return TYPE_META[mapped] ? mapped : "STRING";
}

function normalizeRelationCardinality(value, isForeignKey = false) {
    const raw = String(value ?? "").trim().toUpperCase();
    if (!isForeignKey) return "";
    return CARDINALITY_OPTIONS.includes(raw) ? raw : "N:1";
}

function normalizePosition(position, fallback = { x: 120, y: 120 }) {
    const x = Number(position?.x);
    const y = Number(position?.y);
    return {
        x: Number.isFinite(x) ? x : fallback.x,
        y: Number.isFinite(y) ? y : fallback.y,
    };
}

function createColumn(overrides = {}) {
    const type = normalizeType(overrides.type);
    const foreignKey = Boolean(overrides.fk || overrides.refTableId || overrides.refColumnId);
    return {
        id: overrides.id || uid("col"),
        name: String(overrides.name || "column").trim() || "column",
        type,
        pk: Boolean(overrides.pk),
        fk: foreignKey,
        nn: Boolean(overrides.nn || overrides.pk),
        uq: Boolean(overrides.uq || overrides.pk),
        defaultValue: toCellString(overrides.defaultValue ?? overrides.default ?? ""),
        description: String(overrides.description ?? "").trim(),
        refTableId: overrides.refTableId ?? overrides.ref?.tableId ?? null,
        refColumnId: overrides.refColumnId ?? overrides.ref?.columnId ?? null,
        relationName: String(overrides.relationName ?? overrides.relationshipName ?? "").trim(),
        relationCardinality: normalizeRelationCardinality(
            overrides.relationCardinality ?? overrides.cardinality ?? "",
            foreignKey,
        ),
        enumValues: normalizeEnumValues(overrides.enumValues ?? overrides.options ?? ""),
    };
}

function normalizeFilterPresets(presets) {
    if (!Array.isArray(presets)) return [];
    return presets
        .map((preset) => ({
            id: preset?.id || uid("preset"),
            name: String(preset?.name || "Saved view").trim() || "Saved view",
            query: String(preset?.query || "").trim(),
            sort: {
                columnId: preset?.sort?.columnId || null,
                direction: preset?.sort?.direction === "desc" ? "desc" : preset?.sort?.direction === "asc" ? "asc" : null,
            },
        }))
        .filter((preset) => preset.name);
}

function makeRowFromNamedValues(columns, values) {
    const cells = {};
    columns.forEach((column) => {
        cells[column.id] = toCellString(values[column.name] ?? column.defaultValue ?? "");
    });
    return {
        id: uid("row"),
        cells,
    };
}

function createSeedProject() {
    const playersTableId = uid("table");
    const charactersTableId = uid("table");
    const itemsTableId = uid("table");

    const playerIdCol = uid("col");
    const playerNicknameCol = uid("col");
    const playerCreatedCol = uid("col");

    const characterIdCol = uid("col");
    const characterPlayerIdCol = uid("col");
    const characterClassCol = uid("col");
    const characterLevelCol = uid("col");

    const itemIdCol = uid("col");
    const itemOwnerCol = uid("col");
    const itemNameCol = uid("col");
    const itemPowerCol = uid("col");

    const playersColumns = [
        createColumn({
            id: playerIdCol,
            name: "player_id",
            type: "INT",
            pk: true,
            nn: true,
            uq: true,
            description: "플레이어 고유 식별자",
        }),
        createColumn({
            id: playerNicknameCol,
            name: "nickname",
            type: "STRING",
            nn: true,
            uq: true,
            description: "노출용 닉네임",
        }),
        createColumn({
            id: playerCreatedCol,
            name: "created_at",
            type: "DATE",
            nn: true,
            defaultValue: "2026-01-01",
            description: "가입 일자",
        }),
    ];

    const charactersColumns = [
        createColumn({
            id: characterIdCol,
            name: "character_id",
            type: "INT",
            pk: true,
            nn: true,
            uq: true,
            description: "캐릭터 고유 식별자",
        }),
        createColumn({
            id: characterPlayerIdCol,
            name: "player_id",
            type: "INT",
            fk: true,
            nn: true,
            refTableId: playersTableId,
            refColumnId: playerIdCol,
            description: "소유 플레이어 ID",
        }),
        createColumn({
            id: characterClassCol,
            name: "class_type",
            type: "ENUM",
            nn: true,
            defaultValue: "WARRIOR",
            enumValues: ["WARRIOR", "ARCHER", "MAGE"],
            description: "전투 클래스",
        }),
        createColumn({
            id: characterLevelCol,
            name: "level",
            type: "INT",
            nn: true,
            defaultValue: "1",
            description: "현재 레벨",
        }),
    ];

    const itemsColumns = [
        createColumn({
            id: itemIdCol,
            name: "item_id",
            type: "INT",
            pk: true,
            nn: true,
            uq: true,
            description: "아이템 인스턴스 ID",
        }),
        createColumn({
            id: itemOwnerCol,
            name: "character_id",
            type: "INT",
            fk: true,
            nn: true,
            refTableId: charactersTableId,
            refColumnId: characterIdCol,
            description: "장착 캐릭터 ID",
        }),
        createColumn({
            id: itemNameCol,
            name: "item_name",
            type: "STRING",
            nn: true,
            description: "아이템 이름",
        }),
        createColumn({
            id: itemPowerCol,
            name: "attack_power",
            type: "INT",
            defaultValue: "0",
            description: "공격력 보정값",
        }),
    ];

    return normalizeProject({
        version: 2,
        updatedAt: new Date().toISOString(),
        tables: [
            {
                id: playersTableId,
                name: "Players",
                note: "계정과 플레이어 기본 프로필",
                position: { x: 120, y: 120 },
                columns: playersColumns,
                rows: [
                    makeRowFromNamedValues(playersColumns, {
                        player_id: "1",
                        nickname: "DragonSlayer",
                        created_at: "2026-05-01",
                    }),
                    makeRowFromNamedValues(playersColumns, {
                        player_id: "2",
                        nickname: "MagicMaster",
                        created_at: "2026-05-03",
                    }),
                ],
            },
            {
                id: charactersTableId,
                name: "Characters",
                note: "플레이어가 보유한 전투 캐릭터",
                position: { x: 560, y: 140 },
                columns: charactersColumns,
                rows: [
                    makeRowFromNamedValues(charactersColumns, {
                        character_id: "101",
                        player_id: "1",
                        class_type: "WARRIOR",
                        level: "48",
                    }),
                    makeRowFromNamedValues(charactersColumns, {
                        character_id: "102",
                        player_id: "1",
                        class_type: "ARCHER",
                        level: "31",
                    }),
                    makeRowFromNamedValues(charactersColumns, {
                        character_id: "103",
                        player_id: "2",
                        class_type: "MAGE",
                        level: "57",
                    }),
                ],
            },
            {
                id: itemsTableId,
                name: "Items",
                note: "캐릭터가 착용 중인 장비 인스턴스",
                position: { x: 1020, y: 180 },
                columns: itemsColumns,
                rows: [
                    makeRowFromNamedValues(itemsColumns, {
                        item_id: "5001",
                        character_id: "101",
                        item_name: "Excalibur",
                        attack_power: "120",
                    }),
                    makeRowFromNamedValues(itemsColumns, {
                        item_id: "5002",
                        character_id: "103",
                        item_name: "Elder Wand",
                        attack_power: "85",
                    }),
                ],
            },
        ],
    });
}

function normalizeProject(raw) {
    const source = Array.isArray(raw) ? { tables: raw } : raw && typeof raw === "object" ? raw : { tables: [] };
    const tables = Array.isArray(source.tables) ? source.tables.map(normalizeTable) : [];
    const project = {
        version: 2,
        updatedAt: source.updatedAt || new Date().toISOString(),
        tables,
    };
    repairLegacyReferences(project);
    project.tables.forEach(syncTableRows);
    project.tables.forEach(syncTableArtifacts);
    return project;
}

function normalizeTable(table) {
    const columns = Array.isArray(table?.columns) ? table.columns.map(createColumn) : [];
    if (columns.length === 0) {
        columns.push(
            createColumn({
                name: "id",
                type: "INT",
                pk: true,
                nn: true,
                uq: true,
                description: "기본 키",
            }),
        );
    }

    const rowsSource = Array.isArray(table?.rows) ? table.rows : Array.isArray(table?.data) ? table.data : [];
    const rows = rowsSource.map((row) => normalizeRow(row, columns));

    return {
        id: table?.id || uid("table"),
        name: String(table?.name || "Untitled Table").trim() || "Untitled Table",
        note: String(table?.note ?? "").trim(),
        position: normalizePosition(table?.position, nextTablePosition()),
        columns,
        rows,
        filterPresets: normalizeFilterPresets(table?.filterPresets),
    };
}

function normalizeRow(row, columns) {
    const source = row && typeof row === "object" ? row : {};
    const rawCells =
        source.cells && typeof source.cells === "object" && !Array.isArray(source.cells) ? source.cells : source;
    const cells = {};

    columns.forEach((column) => {
        const value =
            rawCells[column.id] ??
            rawCells[column.name] ??
            rawCells[column.name?.trim?.()] ??
            column.defaultValue ??
            "";
        cells[column.id] = toCellString(value);
    });

    return {
        id: source.id || uid("row"),
        cells,
    };
}

function syncTableRows(table) {
    table.rows = Array.isArray(table.rows) ? table.rows : [];
    table.rows = table.rows.map((row) => {
        const cells = {};
        table.columns.forEach((column) => {
            cells[column.id] = toCellString(row?.cells?.[column.id] ?? row?.cells?.[column.name] ?? column.defaultValue ?? "");
        });
        return {
            id: row?.id || uid("row"),
            cells,
        };
    });
}

function syncTableArtifacts(table) {
    table.filterPresets = normalizeFilterPresets(table.filterPresets).map((preset) => {
        const sortColumnExists = preset.sort.columnId && table.columns.some((column) => column.id === preset.sort.columnId);
        return {
            ...preset,
            sort: sortColumnExists ? preset.sort : { columnId: null, direction: null },
        };
    });
}

function repairLegacyReferences(project) {
    project.tables.forEach((table) => {
        table.columns.forEach((column) => {
            if (!column.fk) {
                column.refTableId = null;
                column.refColumnId = null;
                return;
            }

            if (column.refTableId && column.refColumnId) return;
            const guess = guessReference(project, table.id, column);
            if (guess) {
                column.refTableId = guess.tableId;
                column.refColumnId = guess.columnId;
            }
        });
    });
}

function guessReference(project, sourceTableId, column) {
    const baseName = normalizeKey(column.name).replace(/id$|code$|type$/g, "");
    if (!baseName) return null;

    const candidates = [];
    project.tables.forEach((table) => {
        if (table.id === sourceTableId) return;
        const tableKey = normalizeKey(table.name);
        let score = 0;
        if (tableKey === baseName || tableKey === `${baseName}s`) score += 6;
        if (tableKey.includes(baseName) || baseName.includes(tableKey)) score += 3;

        const preferredColumn =
            table.columns.find((item) => item.pk) ||
            table.columns.find((item) => normalizeKey(item.name) === `${baseName}id`) ||
            table.columns[0];

        if (!preferredColumn) return;
        if (preferredColumn.pk) score += 2;
        if (score > 0) {
            candidates.push({
                score,
                tableId: table.id,
                columnId: preferredColumn.id,
            });
        }
    });

    candidates.sort((left, right) => right.score - left.score);
    return candidates[0] || null;
}

function normalizeKey(value) {
    return String(value ?? "")
        .toLowerCase()
        .replace(/[_\s-]+/g, "")
        .replace(/[^a-z0-9]/g, "");
}

function getTableById(project, tableId) {
    return project.tables.find((table) => table.id === tableId) || null;
}

function getColumnById(table, columnId) {
    return table?.columns.find((column) => column.id === columnId) || null;
}

function getActiveTable() {
    if (state.ui.view.kind !== "data" || !state.ui.view.tableId) return null;
    return getTableById(state.project, state.ui.view.tableId);
}

function nextTablePosition(index = state.project?.tables?.length ?? 0) {
    return {
        x: 120 + index * 40,
        y: 120 + index * 30,
    };
}

function nextTableName(project, base = "New Table") {
    const used = new Set(project.tables.map((table) => table.name));
    if (!used.has(base)) return base;
    let count = 2;
    while (used.has(`${base} ${count}`)) {
        count += 1;
    }
    return `${base} ${count}`;
}

function nextColumnName(table, base = "column") {
    const used = new Set(table.columns.map((column) => column.name));
    if (!used.has(base)) return base;
    let count = 2;
    while (used.has(`${base}_${count}`)) {
        count += 1;
    }
    return `${base}_${count}`;
}

function createBlankTable(project) {
    const idColumn = createColumn({
        name: "id",
        type: "INT",
        pk: true,
        nn: true,
        uq: true,
        description: "기본 키",
    });

    return {
        id: uid("table"),
        name: nextTableName(project),
        note: "",
        position: nextTablePosition(project.tables.length),
        columns: [idColumn],
        rows: [],
        filterPresets: [],
    };
}

function createPresetName(table, query, sort) {
    const sortColumn = sort?.columnId ? getColumnById(table, sort.columnId) : null;
    const base = query ? `Search: ${query}` : sortColumn ? `Sort: ${sortColumn.name}` : "Current view";
    const used = new Set((table.filterPresets || []).map((preset) => preset.name));
    if (!used.has(base)) return base;
    let count = 2;
    while (used.has(`${base} ${count}`)) {
        count += 1;
    }
    return `${base} ${count}`;
}

function buildNewRow(table, sourceRow = null) {
    const cells = {};
    table.columns.forEach((column) => {
        let value = sourceRow?.cells?.[column.id] ?? evaluateColumnDefault(table, column);
        if (column.pk) {
            if (column.type === "INT") {
                value = String(nextNumericPrimaryKey(table, column.id));
            } else if (!sourceRow) {
                value = evaluateColumnDefault(table, column);
            }
        }
        cells[column.id] = toCellString(value);
    });
    return {
        id: uid("row"),
        cells,
    };
}

function evaluateColumnDefault(table, column) {
    const raw = String(column.defaultValue ?? "").trim();
    if (!raw) return "";

    switch (raw) {
        case "{{today}}":
            return new Date().toISOString().slice(0, 10);
        case "{{now}}":
            return new Date().toISOString();
        case "{{uuid}}":
            return typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
                ? crypto.randomUUID()
                : uid("uuid");
        case "{{autoincrement}}":
            return column.type === "INT" ? String(nextNumericPrimaryKey(table, column.id)) : "";
        case "{{true}}":
            return "true";
        case "{{false}}":
            return "false";
        default:
            return raw;
    }
}

function nextNumericPrimaryKey(table, columnId) {
    const values = table.rows
        .map((row) => Number.parseInt(String(row.cells[columnId] ?? "").trim(), 10))
        .filter((value) => Number.isFinite(value));
    const maxValue = values.length ? Math.max(...values) : 0;
    return maxValue + 1;
}

function pushUndo(snapshot) {
    state.history.undo.push(clone(snapshot));
    if (state.history.undo.length > HISTORY_LIMIT) {
        state.history.undo.shift();
    }
}

function replaceProject(nextProject, options = {}) {
    const { recordHistory = false, pushRedo = false, toast = null } = options;
    const currentSnapshot = state.project ? clone(state.project) : null;
    if (recordHistory && currentSnapshot) {
        pushUndo(currentSnapshot);
        state.history.redo = [];
    }
    if (pushRedo && currentSnapshot) {
        state.history.redo.push(currentSnapshot);
    }

    state.project = normalizeProject(nextProject);
    state.project.updatedAt = new Date().toISOString();
    refreshDerivedState();
    persistProject();
    renderApp();

    if (toast) {
        showToast(toast);
    }
}

function updateProject(mutator, toast = null) {
    const draft = clone(state.project);
    mutator(draft);
    replaceProject(draft, { recordHistory: true, toast });
}

function persistProject() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.project));
    state.saveStatus.lastSavedAt = new Date();
}

function loadProject() {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
        try {
            state.project = normalizeProject(JSON.parse(stored));
            state.saveStatus.source = "local";
            state.saveStatus.lastSavedAt = state.project.updatedAt ? new Date(state.project.updatedAt) : new Date();
            refreshDerivedState();
            return;
        } catch (_error) {
            showToast("저장된 프로젝트를 읽지 못해 샘플로 초기화합니다.", "warning");
        }
    }

    state.project = createSeedProject();
    state.saveStatus.source = "demo";
    refreshDerivedState();
    persistProject();
}

function refreshDerivedState() {
    state.validation = validateProject(state.project);
    ensureView();
    syncUiState();
}

function syncUiState() {
    const tableIds = new Set(state.project.tables.map((table) => table.id));
    const keyedMaps = [
        "dataSearch",
        "dataSort",
        "bulkPasteDraft",
        "bulkPasteOpen",
        "selectedRows",
        "bulkEditColumn",
        "bulkEditValue",
    ];

    keyedMaps.forEach((key) => {
        Object.keys(state.ui[key]).forEach((tableId) => {
            if (!tableIds.has(tableId)) {
                delete state.ui[key][tableId];
            }
        });
    });

    state.project.tables.forEach((table) => {
        const rowIds = new Set(table.rows.map((row) => row.id));
        const selected = state.ui.selectedRows[table.id] || {};
        Object.keys(selected).forEach((rowId) => {
            if (!rowIds.has(rowId)) {
                delete selected[rowId];
            }
        });
        state.ui.selectedRows[table.id] = selected;

        const bulkColumnId = state.ui.bulkEditColumn[table.id];
        if (bulkColumnId && !table.columns.some((column) => column.id === bulkColumnId)) {
            delete state.ui.bulkEditColumn[table.id];
        }
    });
}

function ensureView() {
    if (!state.project.tables.length) {
        state.ui.view = { kind: "schema", tableId: null };
        return;
    }

    if (state.ui.view.kind === "data") {
        const table = getTableById(state.project, state.ui.view.tableId);
        if (!table) {
            state.ui.view = { kind: "schema", tableId: null };
        }
    }
}

function setView(kind, tableId = null) {
    state.ui.view = { kind, tableId };
    renderApp();
}

function validateProject(project) {
    const issues = [];
    const tableIssueCounts = {};
    const columnIssueMap = {};
    const cellIssueMap = {};

    function markTable(tableId) {
        tableIssueCounts[tableId] = (tableIssueCounts[tableId] || 0) + 1;
    }

    function markColumn(tableId, columnId) {
        if (!tableId || !columnId) return;
        if (!columnIssueMap[tableId]) columnIssueMap[tableId] = {};
        columnIssueMap[tableId][columnId] = true;
    }

    function markCell(tableId, rowId, columnId) {
        if (!tableId || !rowId || !columnId) return;
        if (!cellIssueMap[tableId]) cellIssueMap[tableId] = {};
        if (!cellIssueMap[tableId][rowId]) cellIssueMap[tableId][rowId] = {};
        cellIssueMap[tableId][rowId][columnId] = true;
    }

    function pushIssue(issue) {
        issues.push(issue);
        if (issue.tableId) markTable(issue.tableId);
        if (issue.columnId) markColumn(issue.tableId, issue.columnId);
        if (issue.rowId && issue.columnId) markCell(issue.tableId, issue.rowId, issue.columnId);
    }

    const duplicateTableNames = new Map();
    project.tables.forEach((table) => {
        const key = table.name.trim().toLowerCase();
        if (!duplicateTableNames.has(key)) duplicateTableNames.set(key, []);
        duplicateTableNames.get(key).push(table.id);
    });

    duplicateTableNames.forEach((tableIds, key) => {
        if (tableIds.length > 1) {
            tableIds.forEach((tableId) => {
                pushIssue({
                    level: "warning",
                    tableId,
                    title: "테이블 이름이 중복됩니다.",
                    detail: `"${key}" 이름을 공유하는 테이블이 여러 개 있습니다. 내보내기와 관계 파악이 불안정해질 수 있습니다.`,
                });
            });
        }
    });

    project.tables.forEach((table) => {
        if (table.columns.length === 0) {
            pushIssue({
                level: "error",
                tableId: table.id,
                title: "컬럼이 없습니다.",
                detail: "테이블에는 최소 1개의 컬럼이 필요합니다.",
            });
        }

        const pkColumns = table.columns.filter((column) => column.pk);
        if (pkColumns.length === 0) {
            pushIssue({
                level: "warning",
                tableId: table.id,
                title: "기본 키가 없습니다.",
                detail: "행 식별과 FK 연결을 위해 PK를 최소 1개 두는 편이 좋습니다.",
            });
        }
        if (pkColumns.length > 1) {
            pushIssue({
                level: "warning",
                tableId: table.id,
                title: "복합 PK가 감지되었습니다.",
                detail: "편집기는 복합 키를 저장은 하지만, 일부 기능은 단일 PK 기준으로 동작합니다.",
            });
        }

        const duplicateColumnNames = new Map();
        table.columns.forEach((column) => {
            const key = column.name.trim().toLowerCase();
            if (!duplicateColumnNames.has(key)) duplicateColumnNames.set(key, []);
            duplicateColumnNames.get(key).push(column.id);

            if (!column.name.trim()) {
                pushIssue({
                    level: "error",
                    tableId: table.id,
                    columnId: column.id,
                    title: "빈 컬럼명이 있습니다.",
                    detail: "모든 컬럼에는 이름이 필요합니다.",
                });
            }

            if (column.type === "ENUM" && column.enumValues.length === 0) {
                pushIssue({
                    level: "warning",
                    tableId: table.id,
                    columnId: column.id,
                    title: "ENUM 옵션이 비어 있습니다.",
                    detail: "ENUM 타입은 허용 값을 쉼표로 입력해 두는 편이 좋습니다.",
                });
            }

            if (column.defaultValue.trim()) {
                const defaultCheck = validateScalar(column, column.defaultValue);
                if (!defaultCheck.valid) {
                    pushIssue({
                        level: "warning",
                        tableId: table.id,
                        columnId: column.id,
                        title: "기본값이 타입 규칙과 맞지 않습니다.",
                        detail: `${column.name}: ${defaultCheck.reason}`,
                    });
                }
            }

            if (column.fk) {
                if (!column.refTableId || !column.refColumnId) {
                    pushIssue({
                        level: "error",
                        tableId: table.id,
                        columnId: column.id,
                        title: "FK 대상이 비어 있습니다.",
                        detail: "FK 체크만 켜져 있고 참조 테이블/컬럼이 선택되지 않았습니다.",
                    });
                } else {
                    const targetTable = getTableById(project, column.refTableId);
                    const targetColumn = getColumnById(targetTable, column.refColumnId);
                    if (!targetTable || !targetColumn) {
                        pushIssue({
                            level: "error",
                            tableId: table.id,
                            columnId: column.id,
                            title: "FK 대상이 유효하지 않습니다.",
                            detail: "삭제되었거나 존재하지 않는 컬럼을 참조하고 있습니다.",
                        });
                    }
                }
            }
        });

        duplicateColumnNames.forEach((columnIds, key) => {
            if (columnIds.length > 1) {
                columnIds.forEach((columnId) => {
                    pushIssue({
                        level: "error",
                        tableId: table.id,
                        columnId,
                        title: "컬럼 이름이 중복됩니다.",
                        detail: `"${key}" 이름을 공유하는 컬럼이 있습니다.`,
                    });
                });
            }
        });

        const uniqueColumns = table.columns.filter((column) => column.pk || column.uq);
        const uniqueTrackers = Object.fromEntries(uniqueColumns.map((column) => [column.id, new Map()]));
        const foreignIndex = buildForeignIndex(project, table);

        table.rows.forEach((row, rowIndex) => {
            table.columns.forEach((column) => {
                const rawValue = toCellString(row.cells[column.id] ?? "");
                const trimmed = rawValue.trim();

                if (column.nn && !trimmed) {
                    pushIssue({
                        level: "error",
                        tableId: table.id,
                        rowId: row.id,
                        columnId: column.id,
                        title: "필수 값이 비어 있습니다.",
                        detail: `${rowIndex + 1}행 / ${column.name} 컬럼은 비워둘 수 없습니다.`,
                    });
                }

                if (trimmed) {
                    const check = validateScalar(column, rawValue);
                    if (!check.valid) {
                        pushIssue({
                            level: "error",
                            tableId: table.id,
                            rowId: row.id,
                            columnId: column.id,
                            title: "값이 타입 규칙과 맞지 않습니다.",
                            detail: `${rowIndex + 1}행 / ${column.name}: ${check.reason}`,
                        });
                    }
                }

                if ((column.pk || column.uq) && trimmed) {
                    const tracker = uniqueTrackers[column.id];
                    const uniqueKey = trimmed.toLowerCase();
                    const existing = tracker.get(uniqueKey);
                    if (existing) {
                        markCell(table.id, existing.rowId, column.id);
                        pushIssue({
                            level: "error",
                            tableId: table.id,
                            rowId: row.id,
                            columnId: column.id,
                            title: "고유값이 중복됩니다.",
                            detail: `${rowIndex + 1}행 / ${column.name} 값이 ${existing.index + 1}행과 중복됩니다.`,
                        });
                    } else {
                        tracker.set(uniqueKey, { rowId: row.id, index: rowIndex });
                    }
                }

                if (column.fk && column.refTableId && column.refColumnId && trimmed) {
                    const referenceSet = foreignIndex[column.id];
                    if (referenceSet && !referenceSet.has(trimmed)) {
                        pushIssue({
                            level: "error",
                            tableId: table.id,
                            rowId: row.id,
                            columnId: column.id,
                            title: "참조 대상이 존재하지 않습니다.",
                            detail: `${rowIndex + 1}행 / ${column.name} 값 "${trimmed}" 를 참조할 대상이 없습니다.`,
                        });
                    }
                }
            });
        });
    });

    const errorCount = issues.filter((issue) => issue.level === "error").length;
    const warningCount = issues.length - errorCount;

    return {
        issues,
        errorCount,
        warningCount,
        tableIssueCounts,
        columnIssueMap,
        cellIssueMap,
    };
}

function buildForeignIndex(project, table) {
    const index = {};
    table.columns.forEach((column) => {
        if (!column.fk || !column.refTableId || !column.refColumnId) return;
        const targetTable = getTableById(project, column.refTableId);
        if (!targetTable) return;
        index[column.id] = new Set(
            targetTable.rows
                .map((row) => toCellString(row.cells[column.refColumnId] ?? "").trim())
                .filter(Boolean),
        );
    });
    return index;
}

function validateScalar(column, rawValue) {
    const value = String(rawValue ?? "").trim();
    if (!value) return { valid: true, reason: "" };

    switch (column.type) {
        case "INT":
            return /^-?\d+$/.test(value)
                ? { valid: true, reason: "" }
                : { valid: false, reason: "정수 형식이어야 합니다." };
        case "FLOAT":
            return /^-?\d+(\.\d+)?$/.test(value)
                ? { valid: true, reason: "" }
                : { valid: false, reason: "실수 형식이어야 합니다." };
        case "BOOLEAN":
            return /^(true|false|1|0|yes|no|y|n)$/i.test(value)
                ? { valid: true, reason: "" }
                : { valid: false, reason: "true / false 계열 값이어야 합니다." };
        case "DATE":
            return /^\d{4}-\d{2}-\d{2}$/.test(value) && !Number.isNaN(new Date(value).getTime())
                ? { valid: true, reason: "" }
                : { valid: false, reason: "YYYY-MM-DD 형식이어야 합니다." };
        case "JSON":
            try {
                JSON.parse(value);
                return { valid: true, reason: "" };
            } catch (_error) {
                return { valid: false, reason: "유효한 JSON 문자열이어야 합니다." };
            }
        case "ENUM":
            return column.enumValues.length === 0 || column.enumValues.includes(value)
                ? { valid: true, reason: "" }
                : { valid: false, reason: "ENUM 허용 값 목록에 없는 값입니다." };
        default:
            return { valid: true, reason: "" };
    }
}

function getProjectStats() {
    return {
        tables: state.project.tables.length,
        columns: state.project.tables.reduce((sum, table) => sum + table.columns.length, 0),
        rows: state.project.tables.reduce((sum, table) => sum + table.rows.length, 0),
        relations: state.project.tables.reduce(
            (sum, table) => sum + table.columns.filter((column) => column.fk && column.refTableId && column.refColumnId).length,
            0,
        ),
    };
}

function renderApp() {
    renderSidebar();
    renderTopbar();
    renderWorkspace();
    if (window.lucide) {
        window.lucide.createIcons();
    }
}

function renderSidebar() {
    const stats = getProjectStats();
    const validationBadge = state.validation.errorCount + state.validation.warningCount;

    dom.sidebar.innerHTML = `
        <div class="brand-shell">
            <div class="brand-mark"><i data-lucide="table-properties" class="w-5 h-5"></i></div>
            <div>
                <h1 class="brand-title">Data Table Editor</h1>
                <p class="brand-copy">스키마, 관계, 샘플 데이터를 한 화면에서 같이 다루는 정적 데이터 설계 워크벤치</p>
            </div>
        </div>

        <section class="side-section">
            <div class="metrics">
                <div class="metric">
                    <span class="metric__value">${stats.tables}</span>
                    <span class="metric__label">Tables</span>
                </div>
                <div class="metric">
                    <span class="metric__value">${stats.columns}</span>
                    <span class="metric__label">Columns</span>
                </div>
                <div class="metric">
                    <span class="metric__value">${stats.rows}</span>
                    <span class="metric__label">Rows</span>
                </div>
                <div class="metric">
                    <span class="metric__value">${state.validation.errorCount}</span>
                    <span class="metric__label">Errors</span>
                </div>
            </div>
        </section>

        <section class="side-section">
            <div class="side-title">Workspace</div>
            <div class="side-nav">
                ${renderSidebarViewButton("schema", "Schema", "columns-3", "컬럼, 제약 조건, 관계 설정")}
                ${renderSidebarViewButton("erd", "ERD", "orbit", "명시적 FK 관계 시각화")}
                ${renderSidebarViewButton("validation", "Validation", "shield-alert", `${validationBadge}개 이슈 확인`)}
            </div>
        </section>

        <section class="side-section">
            <div class="side-title">
                <span>Tables</span>
                <button class="mini-button" data-action="add-table">추가</button>
            </div>
            <div class="side-list">
                ${
                    state.project.tables.length
                        ? state.project.tables.map((table) => renderTableLink(table)).join("")
                        : `<div class="sidebar-footnote">아직 테이블이 없습니다. 새 테이블을 만들고 컬럼과 관계를 추가해 보세요.</div>`
                }
            </div>
        </section>

        <section class="side-section">
            <div class="side-title">Project I/O</div>
            <div class="action-grid">
                <button data-action="import-json">
                    <span class="action-grid__title">Import JSON</span>
                    <span class="action-grid__meta">전체 프로젝트를 불러옵니다.</span>
                </button>
                <button data-action="export-json">
                    <span class="action-grid__title">Export JSON</span>
                    <span class="action-grid__meta">현재 프로젝트 전체를 저장합니다.</span>
                </button>
                <button data-action="import-csv">
                    <span class="action-grid__title">Import CSV</span>
                    <span class="action-grid__meta">단일 시트 데이터를 테이블로 변환합니다.</span>
                </button>
                <button data-action="import-excel">
                    <span class="action-grid__title">Import Excel</span>
                    <span class="action-grid__meta">시트별로 테이블을 생성합니다.</span>
                </button>
                <button data-action="export-excel">
                    <span class="action-grid__title">Export Excel</span>
                    <span class="action-grid__meta">전체 테이블을 워크북으로 내보냅니다.</span>
                </button>
                <button data-action="reset-project">
                    <span class="action-grid__title">Reset Demo</span>
                    <span class="action-grid__meta">샘플 프로젝트로 다시 시작합니다.</span>
                </button>
            </div>
        </section>

        <div class="sidebar-footnote">
            <strong>Autosave:</strong> ${state.saveStatus.source === "local" ? "로컬 저장 사용 중" : "샘플 데이터에서 시작"}<br>
            마지막 저장 ${formatTimestamp(state.saveStatus.lastSavedAt)}
        </div>
    `;
}

function renderSidebarViewButton(kind, label, icon, meta) {
    const active = state.ui.view.kind === kind;
    return `
        <button class="side-button ${active ? "is-active" : ""}" data-action="set-view" data-view="${kind}">
            <span class="side-button__lead">
                <i data-lucide="${icon}" class="w-4 h-4"></i>
                <span class="side-button__copy">
                    <span class="side-button__label">${label}</span>
                    <span class="side-button__meta">${meta}</span>
                </span>
            </span>
        </button>
    `;
}

function renderTableLink(table) {
    const active = state.ui.view.kind === "data" && state.ui.view.tableId === table.id;
    const issues = state.validation.tableIssueCounts[table.id] || 0;
    return `
        <button class="table-link ${active ? "is-active" : ""}" data-action="open-table" data-table-id="${table.id}">
            <span class="table-link__lead">
                <i data-lucide="table" class="w-4 h-4"></i>
                <span class="table-link__copy">
                    <span class="table-link__label">${escapeHtml(table.name)}</span>
                    <span class="table-link__meta">${table.rows.length} rows · ${table.columns.length} cols</span>
                </span>
            </span>
            ${issues ? `<span class="issue-pill">${issues}</span>` : `<span class="table-link__count">Open</span>`}
        </button>
    `;
}

function renderTopbar() {
    const activeTable = getActiveTable();
    let title = "Schema";
    let subtitle = "컬럼, 기본값, 제약 조건, 관계를 한 번에 설계합니다.";
    let actions = "";

    if (state.ui.view.kind === "schema") {
        actions = `
            <button class="ghost-button" data-action="toggle-sql-import">${state.ui.sqlImportOpen ? "Hide SQL import" : "Import SQL"}</button>
            <button class="ghost-button" data-action="copy-sql">Copy SQL</button>
            <button class="ghost-button" data-action="export-sql">Export SQL</button>
        `;
    }

    if (state.ui.view.kind === "erd") {
        title = "ERD";
        subtitle = "실제 FK 대상 기준으로 관계선을 그리고 자동 레이아웃과 SVG 복사를 지원합니다.";
        actions = `
            <button class="ghost-button" data-action="auto-layout">Auto layout</button>
            <button class="ghost-button" data-action="fit-erd">Fit view</button>
            <button class="ghost-button" data-action="copy-svg">Copy SVG</button>
        `;
    } else if (state.ui.view.kind === "validation") {
        title = "Validation";
        subtitle = "스키마와 데이터의 구조적 문제를 즉시 찾아냅니다.";
        actions = `
            <span class="tag tag--error">${state.validation.errorCount} errors</span>
            <span class="tag tag--warning">${state.validation.warningCount} warnings</span>
        `;
    } else if (activeTable) {
        title = activeTable.name;
        subtitle = `${activeTable.rows.length} rows · ${activeTable.columns.length} columns · 검색/정렬 지원`;
        actions = `
            <button class="ghost-button" data-action="toggle-bulk-paste" data-table-id="${activeTable.id}">Bulk paste</button>
            <button class="ghost-button" data-action="export-csv" data-table-id="${activeTable.id}">Export CSV</button>
            <button class="ghost-button" data-action="duplicate-row" data-table-id="${activeTable.id}">Duplicate last</button>
            <button class="solid-button" data-action="add-row" data-table-id="${activeTable.id}">Add row</button>
        `;
    } else if (state.project.tables.length === 0) {
        subtitle = "새 테이블을 추가해서 작업을 시작하세요.";
    }

    dom.topbar.innerHTML = `
        <div class="topbar-copy">
            <h2 class="topbar-title">${escapeHtml(title)}</h2>
            <p class="topbar-subtitle">${escapeHtml(subtitle)}</p>
        </div>
        <div class="topbar-actions">
            <span class="status-chip">Autosaved ${formatTimestamp(state.saveStatus.lastSavedAt)}</span>
            <div class="toolbar-group">
                <button class="icon-button" data-action="undo" ${state.history.undo.length ? "" : "disabled"} title="Undo">
                    <i data-lucide="undo-2" class="w-4 h-4"></i>
                </button>
                <button class="icon-button" data-action="redo" ${state.history.redo.length ? "" : "disabled"} title="Redo">
                    <i data-lucide="redo-2" class="w-4 h-4"></i>
                </button>
                <button class="ghost-button" data-action="add-table">Add table</button>
                ${actions}
            </div>
        </div>
    `;
}

function renderWorkspace() {
    const view = state.ui.view.kind;
    const scrollKey = getScrollKey();
    const previousScroll = state.ui.scrollPositions[scrollKey];

    if (view === "schema") {
        dom.workspace.innerHTML = renderSchemaViewWide();
    } else if (view === "erd") {
        dom.workspace.innerHTML = renderErdView();
    } else if (view === "validation") {
        dom.workspace.innerHTML = renderValidationViewWide();
    } else {
        dom.workspace.innerHTML = renderDataViewWide();
    }

    const scrollRoot = dom.workspace.querySelector("[data-scroll-root]");
    if (scrollRoot && previousScroll) {
        scrollRoot.scrollTop = previousScroll.top;
        scrollRoot.scrollLeft = previousScroll.left;
    }

    requestAnimationFrame(() => {
        applyFocusAfterRender();
        if (state.ui.view.kind === "erd") {
            drawErdLines();
        }
    });
}

function renderSchemaView() {
    const stats = getProjectStats();
    return `
        <div class="workspace-scroll" data-scroll-root="${getScrollKey()}">
            <div class="workspace-stack">
                <section class="panel">
                    <div class="panel-header">
                        <div class="panel-header__copy">
                            <div class="panel-eyebrow">Project pulse</div>
                            <h3 class="panel-title">핵심 구조를 먼저 고정했습니다.</h3>
                            <p class="panel-subtitle">행 데이터는 컬럼 이름이 아니라 컬럼 ID로 저장됩니다. 컬럼명 변경, FK 지정, import 후 정규화가 이제 같은 모델 위에서 돌아갑니다.</p>
                        </div>
                        <div class="toolbar-strip">
                            <span class="tag">Stable cells</span>
                            <span class="tag">Explicit FK</span>
                            <span class="tag">Undo / Redo</span>
                            <span class="tag">Autosave</span>
                        </div>
                    </div>
                    <div class="panel-body">
                        <div class="summary-grid">
                            <div class="summary-card">
                                <div class="summary-card__label">Tables</div>
                                <div class="summary-card__value">${stats.tables}</div>
                            </div>
                            <div class="summary-card">
                                <div class="summary-card__label">Columns</div>
                                <div class="summary-card__value">${stats.columns}</div>
                            </div>
                            <div class="summary-card">
                                <div class="summary-card__label">Rows</div>
                                <div class="summary-card__value">${stats.rows}</div>
                            </div>
                            <div class="summary-card">
                                <div class="summary-card__label">Relations</div>
                                <div class="summary-card__value">${stats.relations}</div>
                            </div>
                        </div>
                    </div>
                </section>

                ${
                    state.ui.sqlImportOpen
                        ? `
                            <section class="panel">
                                <div class="panel-header">
                                    <div class="panel-header__copy">
                                        <div class="panel-eyebrow">SQL import</div>
                                        <h3 class="panel-title">CREATE TABLE DDL에서 스키마를 가져옵니다.</h3>
                                        <p class="panel-subtitle">컬럼, PK, FK, UNIQUE, NOT NULL, DEFAULT, 기본 enum check 구문까지 읽어 테이블로 변환합니다.</p>
                                    </div>
                                </div>
                                <div class="panel-body">
                                    <textarea class="textarea" data-field="sql-import-text" placeholder="CREATE TABLE players (...);">${escapeHtml(state.ui.sqlImportDraft)}</textarea>
                                    <div class="toolbar-strip" style="margin-top: 12px;">
                                        <button class="solid-button" data-action="apply-sql-import" data-mode="append">Append tables</button>
                                        <button class="ghost-button" data-action="apply-sql-import" data-mode="replace">Replace project</button>
                                        <button class="ghost-button" data-action="toggle-sql-import">Close</button>
                                    </div>
                                </div>
                            </section>
                        `
                        : ""
                }

                ${
                    state.project.tables.length
                        ? `<div class="schema-table-list">${state.project.tables.map((table) => renderSchemaTablePanel(table)).join("")}</div>`
                        : renderEmptyState(
                              "테이블이 없습니다",
                              "새 테이블을 추가하고 컬럼, 관계, 샘플 데이터를 설계해 보세요.",
                              "add-table",
                              "첫 테이블 만들기",
                          )
                }
            </div>
        </div>
    `;
}

function renderSchemaTablePanel(table) {
    const pkCount = table.columns.filter((column) => column.pk).length;
    const relationCount = table.columns.filter((column) => column.fk).length;
    const issueCount = state.validation.tableIssueCounts[table.id] || 0;

    return `
        <section class="panel" data-table-panel="${table.id}">
            <div class="panel-header">
                <div class="panel-header__copy">
                    <div class="panel-eyebrow">Table</div>
                    <input
                        class="input input--title"
                        type="text"
                        data-field="table-name"
                        data-table-id="${table.id}"
                        value="${escapeAttr(table.name)}"
                    >
                    <input
                        class="input input--subtle"
                        type="text"
                        data-field="table-note"
                        data-table-id="${table.id}"
                        value="${escapeAttr(table.note)}"
                        placeholder="테이블 설명"
                    >
                </div>
                <div class="toolbar-strip">
                    <span class="tag tag--neutral">${table.rows.length} rows</span>
                    <span class="tag tag--neutral">${pkCount} PK</span>
                    <span class="tag tag--neutral">${relationCount} FK</span>
                    ${issueCount ? `<span class="tag tag--warning">${issueCount} issues</span>` : `<span class="tag">Clean</span>`}
                    <button class="ghost-button" data-action="open-table" data-table-id="${table.id}">Data</button>
                    <button class="danger-button" data-action="delete-table" data-table-id="${table.id}">Delete</button>
                </div>
            </div>
            <div class="panel-body">
                <div class="schema-table-wrap">
                    <table class="schema-table">
                        <thead>
                            <tr>
                                <th style="width: 48px;"></th>
                                <th style="min-width: 180px;">Column</th>
                                <th style="min-width: 120px;">Type</th>
                                <th>PK</th>
                                <th>FK</th>
                                <th>NN</th>
                                <th>UQ</th>
                                <th style="min-width: 200px;">Reference</th>
                                <th style="min-width: 110px;">Cardinality</th>
                                <th style="min-width: 180px;">Relation</th>
                                <th style="min-width: 140px;">Default</th>
                                <th style="min-width: 130px;">Helper</th>
                                <th style="min-width: 180px;">Enum Options</th>
                                <th style="min-width: 200px;">Description</th>
                                <th style="width: 72px;"></th>
                            </tr>
                        </thead>
                        <tbody>
                            ${table.columns.map((column) => renderSchemaColumnRow(table, column)).join("")}
                        </tbody>
                    </table>
                </div>
                <div class="toolbar-strip" style="margin-top: 16px;">
                    <button class="solid-button" data-action="add-column" data-table-id="${table.id}">Add column</button>
                    <button class="ghost-button" data-action="add-row" data-table-id="${table.id}">Add seed row</button>
                    <span class="muted-note">컬럼은 드래그로 순서를 바꿀 수 있고, helper 기본값은 새 행 생성과 bulk paste 빈칸에 적용됩니다.</span>
                </div>
            </div>
        </section>
    `;
}

function renderSchemaColumnRow(table, column) {
    const hasIssue = Boolean(state.validation.columnIssueMap[table.id]?.[column.id]);
    const referenceOptions = getReferenceOptions(table.id, column.refTableId, column.refColumnId);
    return `
        <tr
            class="${hasIssue ? "schema-row--issue" : ""}"
            draggable="true"
            data-drag="column"
            data-drop="column"
            data-table-id="${table.id}"
            data-column-id="${column.id}"
            data-focus-column="${column.id}"
        >
            <td>
                <span class="drag-handle"><i data-lucide="grip-vertical" class="w-4 h-4"></i></span>
            </td>
            <td>
                <input
                    class="input ${hasIssue ? "input--invalid" : ""}"
                    type="text"
                    data-field="column-name"
                    data-table-id="${table.id}"
                    data-column-id="${column.id}"
                    value="${escapeAttr(column.name)}"
                >
            </td>
            <td>
                <select
                    class="select ${hasIssue ? "select--invalid" : ""}"
                    data-field="column-type"
                    data-table-id="${table.id}"
                    data-column-id="${column.id}"
                >
                    ${TYPE_KEYS.map((type) => `<option value="${type}" ${column.type === type ? "selected" : ""}>${TYPE_META[type].label}</option>`).join("")}
                </select>
            </td>
            <td><input class="checkbox" type="checkbox" data-field="column-pk" data-table-id="${table.id}" data-column-id="${column.id}" ${column.pk ? "checked" : ""}></td>
            <td><input class="checkbox" type="checkbox" data-field="column-fk" data-table-id="${table.id}" data-column-id="${column.id}" ${column.fk ? "checked" : ""}></td>
            <td><input class="checkbox" type="checkbox" data-field="column-nn" data-table-id="${table.id}" data-column-id="${column.id}" ${column.nn ? "checked" : ""}></td>
            <td><input class="checkbox" type="checkbox" data-field="column-uq" data-table-id="${table.id}" data-column-id="${column.id}" ${column.uq ? "checked" : ""}></td>
            <td>
                <select
                    class="select ${hasIssue ? "select--invalid" : ""}"
                    data-field="column-ref"
                    data-table-id="${table.id}"
                    data-column-id="${column.id}"
                    ${column.fk ? "" : "disabled"}
                >
                    ${referenceOptions}
                </select>
            </td>
            <td>
                <select
                    class="select"
                    data-field="column-cardinality"
                    data-table-id="${table.id}"
                    data-column-id="${column.id}"
                    ${column.fk ? "" : "disabled"}
                >
                    ${CARDINALITY_OPTIONS.map((option) => `<option value="${option}" ${column.relationCardinality === option ? "selected" : ""}>${option}</option>`).join("")}
                </select>
            </td>
            <td>
                <input
                    class="input"
                    type="text"
                    data-field="column-relation-name"
                    data-table-id="${table.id}"
                    data-column-id="${column.id}"
                    value="${escapeAttr(column.relationName)}"
                    placeholder="owns / belongs_to"
                    ${column.fk ? "" : "disabled"}
                >
            </td>
            <td>
                <input
                    class="input"
                    type="text"
                    data-field="column-default"
                    data-table-id="${table.id}"
                    data-column-id="${column.id}"
                    value="${escapeAttr(column.defaultValue)}"
                >
            </td>
            <td>
                <select
                    class="select"
                    data-field="column-default-helper"
                    data-table-id="${table.id}"
                    data-column-id="${column.id}"
                >
                    ${DEFAULT_HELPERS.map((helper) => `<option value="${helper.value}" ${getDefaultHelperSelection(column.defaultValue) === helper.value ? "selected" : ""}>${helper.label}</option>`).join("")}
                </select>
            </td>
            <td>
                <input
                    class="input"
                    type="text"
                    data-field="column-enum"
                    data-table-id="${table.id}"
                    data-column-id="${column.id}"
                    value="${escapeAttr(column.enumValues.join(", "))}"
                    placeholder="ACTIVE, PAUSED"
                    ${column.type === "ENUM" ? "" : "disabled"}
                >
            </td>
            <td>
                <input
                    class="input"
                    type="text"
                    data-field="column-description"
                    data-table-id="${table.id}"
                    data-column-id="${column.id}"
                    value="${escapeAttr(column.description)}"
                >
            </td>
            <td>
                <button class="icon-button" data-action="delete-column" data-table-id="${table.id}" data-column-id="${column.id}">
                    <i data-lucide="trash-2" class="w-4 h-4"></i>
                </button>
            </td>
        </tr>
    `;
}

function getReferenceOptions(sourceTableId, selectedTableId, selectedColumnId) {
    const options = ['<option value="">Select target</option>'];
    state.project.tables.forEach((table) => {
        if (table.id === sourceTableId) return;
        table.columns.forEach((column) => {
            const value = `${table.id}:${column.id}`;
            const label = `${table.name}.${column.name}${column.pk ? " (PK)" : ""}`;
            options.push(`<option value="${value}" ${selectedTableId === table.id && selectedColumnId === column.id ? "selected" : ""}>${escapeHtml(label)}</option>`);
        });
    });
    return options.join("");
}

function renderDataView() {
    const table = getActiveTable();
    if (!table) {
        return `
            <div class="workspace-scroll" data-scroll-root="${getScrollKey()}">
                ${renderEmptyState("열린 테이블이 없습니다", "왼쪽 사이드바에서 테이블을 선택해 데이터 편집으로 이동하세요.", "set-view", "Schema", "schema")}
            </div>
        `;
    }

    const query = state.ui.dataSearch[table.id] || "";
    const sort = state.ui.dataSort[table.id] || { columnId: null, direction: null };
    const rows = getVisibleRows(table, query, sort);
    const presets = table.filterPresets || [];
    const bulkPasteOpen = Boolean(state.ui.bulkPasteOpen[table.id]);
    const bulkPasteDraft = state.ui.bulkPasteDraft[table.id] || "";
    const selectedMap = state.ui.selectedRows[table.id] || {};
    const selectedVisibleCount = rows.filter((row) => selectedMap[row.id]).length;
    const totalSelectedCount = Object.keys(selectedMap).length;
    const bulkEditColumnId = state.ui.bulkEditColumn[table.id] || table.columns[0]?.id || "";
    const bulkEditValue = state.ui.bulkEditValue[table.id] || "";

    return `
        <div class="workspace-scroll" data-scroll-root="${getScrollKey()}">
            <div class="data-shell">
                <section class="panel">
                    <div class="panel-header">
                        <div class="panel-header__copy">
                            <div class="panel-eyebrow">Data Editor</div>
                            <h3 class="panel-title">${escapeHtml(table.name)}</h3>
                            <p class="panel-subtitle">${escapeHtml(table.note || "샘플 데이터를 입력하고 정렬/검색/검증 상태를 확인합니다.")}</p>
                        </div>
                        <div class="toolbar-strip">
                            <span class="tag tag--neutral">${table.rows.length} rows</span>
                            <span class="tag tag--neutral">${table.columns.length} columns</span>
                            ${state.validation.tableIssueCounts[table.id] ? `<span class="tag tag--warning">${state.validation.tableIssueCounts[table.id]} issues</span>` : `<span class="tag">Valid</span>`}
                            <button class="ghost-button" data-action="set-view" data-view="schema">Open schema</button>
                            <button class="danger-button" data-action="clear-rows" data-table-id="${table.id}">Clear rows</button>
                        </div>
                    </div>
                    <div class="panel-body">
                        <div class="data-toolbar">
                            <input
                                class="input"
                                style="max-width: 360px;"
                                type="search"
                                data-field="data-search"
                                data-table-id="${table.id}"
                                value="${escapeAttr(query)}"
                                placeholder="행 전체 검색"
                            >
                            <button class="solid-button" data-action="add-row" data-table-id="${table.id}">Add row</button>
                            <button class="ghost-button" data-action="duplicate-row" data-table-id="${table.id}">Duplicate last row</button>
                            <button class="ghost-button" data-action="save-filter-preset" data-table-id="${table.id}">Save view</button>
                            <button class="ghost-button" data-action="reset-data-view" data-table-id="${table.id}">Reset view</button>
                            <button class="ghost-button" data-action="toggle-bulk-paste" data-table-id="${table.id}">${bulkPasteOpen ? "Hide paste" : "Bulk paste"}</button>
                            <button class="ghost-button" data-action="export-csv" data-table-id="${table.id}">Export CSV</button>
                        </div>
                        <div class="bulk-edit-bar" style="margin-top: 16px;">
                            <span class="tag tag--neutral">${totalSelectedCount} selected</span>
                            <select class="select" style="max-width: 220px;" data-field="bulk-edit-column" data-table-id="${table.id}">
                                ${table.columns.map((column) => `<option value="${column.id}" ${bulkEditColumnId === column.id ? "selected" : ""}>${escapeHtml(column.name)}</option>`).join("")}
                            </select>
                            <input
                                class="input"
                                style="max-width: 240px;"
                                type="text"
                                data-field="bulk-edit-value"
                                data-table-id="${table.id}"
                                value="${escapeAttr(bulkEditValue)}"
                                placeholder="Apply value to selected rows"
                            >
                            <button class="ghost-button" data-action="apply-bulk-edit" data-table-id="${table.id}">Apply to selected</button>
                            <button class="ghost-button" data-action="clear-selected-rows" data-table-id="${table.id}">Clear selection</button>
                            <button class="danger-button" data-action="delete-selected-rows" data-table-id="${table.id}" ${totalSelectedCount ? "" : "disabled"}>Delete selected</button>
                        </div>
                        ${
                            presets.length
                                ? `<div class="preset-bar" style="margin-top: 16px;">${presets.map((preset) => renderFilterPresetChip(table, preset)).join("")}</div>`
                                : ""
                        }
                        ${
                            bulkPasteOpen
                                ? `
                                    <div class="paste-panel" style="margin-top: 16px;">
                                        <div class="panel-eyebrow">Bulk paste</div>
                                        <p class="muted-note">Paste TSV or CSV rows. If the first row matches column names, it will be treated as a header.</p>
                                        <textarea
                                            class="textarea"
                                            data-field="bulk-paste-text"
                                            data-table-id="${table.id}"
                                            placeholder="player_id\tnickname\tcreated_at&#10;3\tNewUser\t2026-05-21"
                                        >${escapeHtml(bulkPasteDraft)}</textarea>
                                        <div class="toolbar-strip" style="margin-top: 12px;">
                                            <button class="solid-button" data-action="apply-bulk-paste" data-table-id="${table.id}">Append rows</button>
                                            <button class="ghost-button" data-action="toggle-bulk-paste" data-table-id="${table.id}">Close</button>
                                        </div>
                                    </div>
                                `
                                : ""
                        }

                        <div class="data-table-wrap" style="margin-top: 16px;">
                            <table class="data-table">
                                <thead>
                                    <tr>
                                        <th style="width: 54px;">
                                            <input
                                                class="checkbox"
                                                type="checkbox"
                                                data-field="select-visible-rows"
                                                data-table-id="${table.id}"
                                                ${rows.length > 0 && selectedVisibleCount === rows.length ? "checked" : ""}
                                            >
                                        </th>
                                        <th style="width: 66px;">#</th>
                                        ${table.columns.map((column) => renderDataHeaderCell(table, column, sort)).join("")}
                                        <th style="width: 88px;"></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${
                                        rows.length
                                            ? rows.map((row, index) => renderDataRow(table, row, index)).join("")
                                            : `<tr><td colspan="${table.columns.length + 3}">${renderTableEmptyRow("조건에 맞는 데이터가 없습니다.")}</td></tr>`
                                    }
                                </tbody>
                            </table>
                        </div>
                    </div>
                </section>
            </div>
        </div>
    `;
}

function renderFilterPresetChip(table, preset) {
    const sortColumn = preset.sort.columnId ? getColumnById(table, preset.sort.columnId) : null;
    const detail = [preset.query || null, sortColumn ? `${sortColumn.name} ${preset.sort.direction || ""}`.trim() : null]
        .filter(Boolean)
        .join(" · ");

    return `
        <div class="preset-chip">
            <button class="preset-chip__main" data-action="apply-filter-preset" data-table-id="${table.id}" data-preset-id="${preset.id}">
                <span class="preset-chip__name">${escapeHtml(preset.name)}</span>
                ${detail ? `<span class="preset-chip__meta">${escapeHtml(detail)}</span>` : ""}
            </button>
            <button class="preset-chip__delete" data-action="delete-filter-preset" data-table-id="${table.id}" data-preset-id="${preset.id}" title="Delete preset">
                <i data-lucide="x" class="w-3.5 h-3.5"></i>
            </button>
        </div>
    `;
}

function renderDataHeaderCell(table, column, sort) {
    const active = sort.columnId === column.id;
    const direction = active ? sort.direction : null;
    const mark = direction === "asc" ? "↑" : direction === "desc" ? "↓" : "";
    return `
        <th>
            <button
                class="ghost-button"
                style="width: 100%; justify-content: space-between;"
                data-action="sort-column"
                data-table-id="${table.id}"
                data-column-id="${column.id}"
            >
                <span>${escapeHtml(column.name)}</span>
                <span>${mark}</span>
            </button>
        </th>
    `;
}

function renderDataRow(table, row, index) {
    const focus = state.ui.focus?.rowId === row.id;
    const selected = Boolean(state.ui.selectedRows[table.id]?.[row.id]);
    const rowClassName = [focus ? "data-row--focus" : "", selected ? "data-row--selected" : ""].filter(Boolean).join(" ");
    return `
        <tr class="${rowClassName}" data-focus-row="${row.id}">
            <td>
                <input
                    class="checkbox"
                    type="checkbox"
                    data-field="row-selected"
                    data-table-id="${table.id}"
                    data-row-id="${row.id}"
                    ${selected ? "checked" : ""}
                >
            </td>
            <td class="row-index">${index + 1}</td>
            ${table.columns.map((column) => renderDataCell(table, row, column)).join("")}
            <td>
                <button class="icon-button" data-action="delete-row" data-table-id="${table.id}" data-row-id="${row.id}">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </td>
        </tr>
    `;
}

function renderDataCell(table, row, column) {
    const rawValue = toCellString(row.cells[column.id] ?? "");
    const invalid = Boolean(state.validation.cellIssueMap[table.id]?.[row.id]?.[column.id]);
    const common = `data-field="cell-value" data-table-id="${table.id}" data-row-id="${row.id}" data-column-id="${column.id}"`;
    const cssClass = `input ${invalid ? "input--invalid" : ""}`;
    const selectClass = `select ${invalid ? "select--invalid" : ""}`;

    let control = "";
    if (column.type === "BOOLEAN") {
        control = `
            <select class="${selectClass}" ${common}>
                <option value="" ${rawValue === "" ? "selected" : ""}></option>
                <option value="true" ${rawValue === "true" ? "selected" : ""}>true</option>
                <option value="false" ${rawValue === "false" ? "selected" : ""}>false</option>
            </select>
        `;
    } else if (column.type === "ENUM") {
        const options = Array.from(new Set([...column.enumValues, rawValue].filter(Boolean)));
        control = `
            <select class="${selectClass}" ${common}>
                <option value="" ${rawValue === "" ? "selected" : ""}></option>
                ${options.map((option) => `<option value="${escapeAttr(option)}" ${option === rawValue ? "selected" : ""}>${escapeHtml(option)}</option>`).join("")}
            </select>
        `;
    } else {
        const inputType = column.type === "DATE" ? "date" : "text";
        control = `
            <input
                class="${cssClass}"
                type="${inputType}"
                ${common}
                value="${escapeAttr(rawValue)}"
                placeholder="${escapeAttr(TYPE_META[column.type].placeholder)}"
            >
        `;
    }

    return `
        <td>
            <div style="display: flex; align-items: center; gap: 8px;">
                ${control}
                ${invalid ? '<span class="cell-issue-dot"></span>' : ""}
            </div>
        </td>
    `;
}

function renderValidationView() {
    const issues = [...state.validation.issues].sort((left, right) => {
        if (left.level === right.level) return 0;
        return left.level === "error" ? -1 : 1;
    });

    return `
        <div class="workspace-scroll" data-scroll-root="${getScrollKey()}">
            <div class="workspace-stack">
                <section class="panel">
                    <div class="panel-header">
                        <div class="panel-header__copy">
                            <div class="panel-eyebrow">Validation</div>
                            <h3 class="panel-title">스키마와 데이터 상태를 한 번에 점검합니다.</h3>
                            <p class="panel-subtitle">컬럼 중복, 타입 불일치, 고유값 충돌, 잘못된 FK 참조를 즉시 확인할 수 있습니다.</p>
                        </div>
                        <div class="toolbar-strip">
                            <span class="tag tag--error">${state.validation.errorCount} errors</span>
                            <span class="tag tag--warning">${state.validation.warningCount} warnings</span>
                        </div>
                    </div>
                    <div class="panel-body">
                        <div class="summary-grid">
                            <div class="summary-card">
                                <div class="summary-card__label">Errors</div>
                                <div class="summary-card__value">${state.validation.errorCount}</div>
                            </div>
                            <div class="summary-card">
                                <div class="summary-card__label">Warnings</div>
                                <div class="summary-card__value">${state.validation.warningCount}</div>
                            </div>
                            <div class="summary-card">
                                <div class="summary-card__label">Affected Tables</div>
                                <div class="summary-card__value">${Object.keys(state.validation.tableIssueCounts).length}</div>
                            </div>
                            <div class="summary-card">
                                <div class="summary-card__label">Clean Tables</div>
                                <div class="summary-card__value">${Math.max(0, state.project.tables.length - Object.keys(state.validation.tableIssueCounts).length)}</div>
                            </div>
                        </div>
                    </div>
                </section>

                <section class="panel">
                    <div class="panel-header">
                        <div class="panel-header__copy">
                            <div class="panel-eyebrow">Issue list</div>
                            <h3 class="panel-title">전체 이슈</h3>
                            <p class="panel-subtitle">오류는 먼저, 경고는 그 다음으로 정렬했습니다.</p>
                        </div>
                    </div>
                    <div class="panel-body">
                        ${
                            issues.length
                                ? `<div class="validation-list">${issues.map((issue) => renderIssueRow(issue)).join("")}</div>`
                                : renderEmptyState("이슈가 없습니다", "현재 프로젝트는 검증 기준을 통과했습니다.")
                        }
                    </div>
                </section>
            </div>
        </div>
    `;
}

function renderIssueRow(issue) {
    const table = issue.tableId ? getTableById(state.project, issue.tableId) : null;
    const location = [];
    if (table) location.push(table.name);
    if (issue.rowId) location.push("row");
    return `
        <div class="issue-row">
            <div class="issue-row__copy">
                <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
                    <span class="issue-level issue-level--${issue.level}">${issue.level.toUpperCase()}</span>
                    ${table ? `<span class="tag tag--neutral">${escapeHtml(table.name)}</span>` : ""}
                </div>
                <div class="issue-row__title">${escapeHtml(issue.title)}</div>
                <div class="issue-row__detail">${escapeHtml(issue.detail)}</div>
                ${location.length ? `<div class="issue-row__meta">${escapeHtml(location.join(" · "))}</div>` : ""}
            </div>
            <div class="issue-row__actions">
                ${issue.tableId ? `<button class="ghost-button" data-action="jump-schema" data-table-id="${issue.tableId}" data-column-id="${issue.columnId || ""}">Open schema</button>` : ""}
                ${issue.tableId && issue.rowId ? `<button class="solid-button" data-action="jump-data" data-table-id="${issue.tableId}" data-row-id="${issue.rowId}" data-column-id="${issue.columnId || ""}">Open data</button>` : ""}
            </div>
        </div>
    `;
}

function renderErdView() {
    return `
        <div class="erd-shell">
            <div class="erd-toolbar">
                <button class="icon-button" data-action="auto-layout" title="Auto layout"><i data-lucide="wand-sparkles" class="w-4 h-4"></i></button>
                <button class="icon-button" data-action="fit-erd" title="Fit"><i data-lucide="maximize" class="w-4 h-4"></i></button>
                <button class="icon-button" data-action="copy-svg" title="Copy SVG"><i data-lucide="copy" class="w-4 h-4"></i></button>
                <span class="tag tag--neutral">Zoom ${(state.ui.erd.zoom * 100).toFixed(0)}%</span>
            </div>
            <div id="erd-viewport" class="erd-viewport" data-erd-viewport>
                <div
                    id="erd-canvas"
                    class="erd-canvas"
                    style="transform: translate(${state.ui.erd.panX}px, ${state.ui.erd.panY}px) scale(${state.ui.erd.zoom});"
                >
                    <svg id="erd-svg" class="erd-svg"></svg>
                    ${state.project.tables.map((table) => renderErdNode(table)).join("")}
                </div>
            </div>
        </div>
    `;
}

function renderErdNode(table) {
    return `
        <section
            class="erd-node"
            data-erd-node="${table.id}"
            style="left: ${table.position.x}px; top: ${table.position.y}px;"
        >
            <div class="erd-node__head" data-erd-drag-handle data-table-id="${table.id}">
                <div>
                    <div class="erd-node__title">${escapeHtml(table.name)}</div>
                    <div class="erd-node__meta">${table.rows.length} rows · ${table.columns.length} cols</div>
                </div>
                <button class="icon-button" data-action="open-table" data-table-id="${table.id}">
                    <i data-lucide="arrow-up-right" class="w-4 h-4"></i>
                </button>
            </div>
            <div class="erd-list">
                ${table.columns.map((column) => renderErdColumn(column)).join("")}
            </div>
        </section>
    `;
}

function renderErdColumn(column) {
    const relationMeta =
        column.fk && (column.relationCardinality || column.relationName)
            ? [column.relationCardinality || null, column.relationName || null].filter(Boolean).join(" · ")
            : "";
    return `
        <div class="erd-row" data-erd-column="${column.id}">
            <div class="erd-row__name">
                ${column.pk ? '<span class="relation-badge relation-badge--pk">PK</span>' : column.fk ? '<span class="relation-badge relation-badge--fk">FK</span>' : ""}
                <div class="erd-row__copy">
                    <span class="erd-row__text">${escapeHtml(column.name)}</span>
                    ${relationMeta ? `<span class="erd-row__meta">${escapeHtml(relationMeta)}</span>` : ""}
                </div>
            </div>
            <span class="erd-row__type">${escapeHtml(column.type)}</span>
        </div>
    `;
}

function renderEmptyState(title, copy, action = null, actionLabel = "", actionValue = "", variant = "default") {
    const actionMarkup = action
        ? `<button class="solid-button" style="margin-top: 14px;" data-action="${action}" ${actionValue ? `data-view="${escapeAttr(actionValue)}"` : ""}>${escapeHtml(actionLabel)}</button>`
        : "";
    const variantClass = variant === "compact" ? " empty-state--compact" : "";
    return `
        <div class="empty-state${variantClass}">
            <div>
                <i data-lucide="database-zap" class="w-10 h-10"></i>
                <h3>${escapeHtml(title)}</h3>
                <p>${escapeHtml(copy)}</p>
                ${actionMarkup}
            </div>
        </div>
    `;
}

function renderTableEmptyRow(message) {
    return `
        <div class="empty-state" style="padding: 32px 12px;">
            <div>
                <h3 style="margin: 0 0 6px;">${escapeHtml(message)}</h3>
                <p style="margin: 0;">검색 조건을 비우거나 새 행을 추가해 보세요.</p>
            </div>
        </div>
    `;
}

function formatTimestamp(date) {
    if (!date) return "방금 전";
    try {
        return new Intl.DateTimeFormat("ko-KR", {
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit",
        }).format(date);
    } catch (_error) {
        return "";
    }
}

function getScrollKey() {
    return `${state.ui.view.kind}:${state.ui.view.tableId || "root"}`;
}

function showToast(message, tone = "info") {
    const toast = document.createElement("div");
    toast.className = `toast ${tone === "error" ? "toast--error" : tone === "warning" ? "toast--warning" : ""}`;
    toast.innerHTML = `
        <i data-lucide="${tone === "error" ? "octagon-alert" : tone === "warning" ? "triangle-alert" : "check-circle-2"}" class="w-5 h-5"></i>
        <span>${escapeHtml(message)}</span>
    `;
    dom.toastHost.appendChild(toast);
    if (window.lucide) {
        window.lucide.createIcons({ root: toast });
    }
    window.setTimeout(() => {
        toast.remove();
    }, 2600);
}

function getVisibleRows(table, query, sort) {
    const lowered = query.trim().toLowerCase();
    let rows = [...table.rows];

    if (lowered) {
        rows = rows.filter((row) =>
            table.columns.some((column) => toCellString(row.cells[column.id] ?? "").toLowerCase().includes(lowered)),
        );
    }

    if (sort.columnId && sort.direction) {
        rows.sort((left, right) => compareRowsForSort(table, left, right, sort.columnId, sort.direction));
    }

    return rows;
}

function getCurrentTableViewState(tableId) {
    const table = getTableById(state.project, tableId);
    const query = state.ui.dataSearch[tableId] || "";
    const sort = state.ui.dataSort[tableId] || { columnId: null, direction: null };
    const rows = table ? getVisibleRows(table, query, sort) : [];
    const selectedMap = state.ui.selectedRows[tableId] || {};
    return {
        table,
        query,
        sort,
        rows,
        selectedMap,
    };
}

function compareRowsForSort(table, leftRow, rightRow, columnId, direction) {
    const column = getColumnById(table, columnId);
    const left = toCellString(leftRow.cells[columnId] ?? "").trim();
    const right = toCellString(rightRow.cells[columnId] ?? "").trim();
    let result = 0;

    if (column?.type === "INT" || column?.type === "FLOAT") {
        result = Number(left || 0) - Number(right || 0);
    } else if (column?.type === "DATE") {
        result = new Date(left || "1970-01-01").getTime() - new Date(right || "1970-01-01").getTime();
    } else if (column?.type === "BOOLEAN") {
        result = String(left).localeCompare(String(right));
    } else {
        result = left.localeCompare(right, "ko");
    }

    return direction === "asc" ? result : -result;
}

function applyFocusAfterRender() {
    if (!state.ui.focus) return;
    const { rowId, columnId, tableId } = state.ui.focus;

    if (state.ui.view.kind === "data" && state.ui.view.tableId === tableId && rowId) {
        const row = dom.workspace.querySelector(`[data-focus-row="${rowId}"]`);
        row?.scrollIntoView({ block: "center", behavior: "smooth" });
    }

    if (state.ui.view.kind === "schema" && columnId) {
        const columnRow = dom.workspace.querySelector(`[data-focus-column="${columnId}"]`);
        columnRow?.scrollIntoView({ block: "center", behavior: "smooth" });
    }

    state.ui.focus = null;
}

function rememberScroll(target) {
    if (!target?.dataset?.scrollRoot) return;
    state.ui.scrollPositions[target.dataset.scrollRoot] = {
        top: target.scrollTop,
        left: target.scrollLeft,
    };
}

function bindEvents() {
    document.addEventListener("click", handleClick);
    document.addEventListener("change", handleChange);
    document.addEventListener("input", handleInput);
    document.addEventListener("keydown", handleKeydown);
    document.addEventListener("mousedown", handleMouseDown);
    document.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseup", handleMouseUp);
    document.addEventListener("dragstart", handleDragStart);
    document.addEventListener("dragover", handleDragOver);
    document.addEventListener("drop", handleDrop);
    document.addEventListener("dragend", handleDragEnd);
    document.addEventListener(
        "scroll",
        (event) => {
            rememberScroll(event.target);
        },
        true,
    );

    dom.workspace.addEventListener("wheel", handleErdWheel, { passive: false });
    dom.jsonInput.addEventListener("change", handleJsonFileChange);
    dom.csvInput.addEventListener("change", handleCsvFileChange);
    dom.excelInput.addEventListener("change", handleExcelFileChange);
}

function handleClick(event) {
    const button = event.target.closest("[data-action]");
    if (!button) return;
    const action = button.dataset.action;

    switch (action) {
        case "set-view":
            setView(button.dataset.view);
            return;
        case "open-table":
            setView("data", button.dataset.tableId);
            return;
        case "add-table":
            updateProject((project) => {
                project.tables.push(createBlankTable(project));
            }, "새 테이블을 추가했습니다.");
            return;
        case "delete-table":
            if (!window.confirm("이 테이블을 삭제할까요?")) return;
            updateProject((project) => {
                project.tables = project.tables.filter((table) => table.id !== button.dataset.tableId);
                clearBrokenReferences(project);
            }, "테이블을 삭제했습니다.");
            return;
        case "add-column":
            updateProject((project) => {
                const table = getTableById(project, button.dataset.tableId);
                if (!table) return;
                const column = createColumn({
                    name: nextColumnName(table),
                    type: "STRING",
                    description: "",
                });
                table.columns.push(column);
                syncTableRows(table);
                state.ui.focus = { tableId: table.id, columnId: column.id };
            }, "컬럼을 추가했습니다.");
            return;
        case "delete-column":
            if (!window.confirm("이 컬럼을 삭제할까요?")) return;
            updateProject((project) => {
                const table = getTableById(project, button.dataset.tableId);
                if (!table) return;
                table.columns = table.columns.filter((column) => column.id !== button.dataset.columnId);
                table.rows.forEach((row) => {
                    delete row.cells[button.dataset.columnId];
                });
                clearBrokenReferences(project);
            }, "컬럼을 삭제했습니다.");
            return;
        case "add-row":
            updateProject((project) => {
                const table = getTableById(project, button.dataset.tableId);
                if (!table) return;
                const row = buildNewRow(table);
                table.rows.push(row);
                state.ui.focus = { tableId: table.id, rowId: row.id };
            }, "행을 추가했습니다.");
            return;
        case "duplicate-row":
            updateProject((project) => {
                const table = getTableById(project, button.dataset.tableId);
                if (!table || table.rows.length === 0) return;
                const source = button.dataset.rowId
                    ? table.rows.find((row) => row.id === button.dataset.rowId)
                    : table.rows[table.rows.length - 1];
                if (!source) return;
                const row = buildNewRow(table, source);
                table.rows.push(row);
                state.ui.focus = { tableId: table.id, rowId: row.id };
            }, "행을 복제했습니다.");
            return;
        case "delete-row":
            updateProject((project) => {
                const table = getTableById(project, button.dataset.tableId);
                if (!table) return;
                table.rows = table.rows.filter((row) => row.id !== button.dataset.rowId);
            }, "행을 삭제했습니다.");
            return;
        case "clear-rows":
            if (!window.confirm("이 테이블의 모든 행을 비울까요?")) return;
            updateProject((project) => {
                const table = getTableById(project, button.dataset.tableId);
                if (!table) return;
                table.rows = [];
            }, "모든 행을 삭제했습니다.");
            return;
        case "undo":
            performUndo();
            return;
        case "redo":
            performRedo();
            return;
        case "sort-column":
            toggleSort(button.dataset.tableId, button.dataset.columnId);
            return;
        case "jump-schema":
            state.ui.focus = {
                tableId: button.dataset.tableId,
                columnId: button.dataset.columnId || null,
            };
            setView("schema");
            return;
        case "jump-data":
            state.ui.focus = {
                tableId: button.dataset.tableId,
                rowId: button.dataset.rowId || null,
                columnId: button.dataset.columnId || null,
            };
            setView("data", button.dataset.tableId);
            return;
        case "import-json":
            dom.jsonInput.click();
            return;
        case "import-csv":
            dom.csvInput.click();
            return;
        case "import-excel":
            dom.excelInput.click();
            return;
        case "export-json":
            exportJson();
            return;
        case "copy-sql":
            copyProjectSql();
            return;
        case "export-sql":
            exportSql();
            return;
        case "toggle-sql-import":
            state.ui.sqlImportOpen = !state.ui.sqlImportOpen;
            renderWorkspace();
            return;
        case "apply-sql-import": {
            const mode = button.dataset.mode === "replace" ? "replace" : "append";
            const sqlText = state.ui.sqlImportDraft || "";
            if (!sqlText.trim()) {
                showToast("SQL text is empty.", "warning");
                return;
            }

            try {
                const importedTables = parseSqlTables(sqlText, mode === "append" ? state.project : { tables: [] });
                if (!importedTables.length) {
                    showToast("No CREATE TABLE blocks were found.", "warning");
                    return;
                }

                if (mode === "replace") {
                    const nextProject = {
                        version: 2,
                        updatedAt: new Date().toISOString(),
                        tables: importedTables,
                    };
                    state.ui.sqlImportDraft = "";
                    state.ui.sqlImportOpen = false;
                    replaceProject(nextProject, { recordHistory: true, toast: `Imported ${importedTables.length} tables from SQL.` });
                } else {
                    updateProject((project) => {
                        importedTables.forEach((table) => project.tables.push(table));
                        state.ui.sqlImportDraft = "";
                        state.ui.sqlImportOpen = false;
                    }, `Imported ${importedTables.length} tables from SQL.`);
                }

                setView("data", importedTables[0].id);
            } catch (error) {
                showToast(error?.message || "Failed to parse SQL.", "error");
            }
            return;
        }
        case "export-csv":
            exportCsv(button.dataset.tableId || state.ui.view.tableId);
            return;
        case "export-excel":
            exportExcel();
            return;
        case "copy-svg":
            copyErdSvg();
            return;
        case "auto-layout":
            autoLayoutErd();
            return;
        case "fit-erd":
            fitErdView();
            return;
        case "edit-relation": {
            const sourceTable = getTableById(state.project, button.dataset.tableId);
            const sourceColumn = getColumnById(sourceTable, button.dataset.columnId);
            if (!sourceTable || !sourceColumn || !sourceColumn.fk) return;

            const nextName = window.prompt("Relation label", sourceColumn.relationName || "");
            if (nextName === null) return;
            const nextCardinality = window.prompt("Cardinality (N:1, 1:1, 1:N, N:N)", sourceColumn.relationCardinality || "N:1");
            if (nextCardinality === null) return;

            updateProject((project) => {
                const table = getTableById(project, button.dataset.tableId);
                const column = getColumnById(table, button.dataset.columnId);
                if (!column) return;
                column.relationName = nextName.trim();
                column.relationCardinality = normalizeRelationCardinality(nextCardinality, column.fk);
                state.ui.focus = { tableId: table.id, columnId: column.id };
            }, "Updated relation metadata.");
            return;
        }
        case "toggle-bulk-paste":
            state.ui.bulkPasteOpen[button.dataset.tableId] = !state.ui.bulkPasteOpen[button.dataset.tableId];
            renderWorkspace();
            return;
        case "apply-bulk-paste": {
            const tableId = button.dataset.tableId;
            const draft = state.ui.bulkPasteDraft[tableId] || "";
            const matrix = parseDelimitedMatrix(draft);
            if (!matrix.length) {
                showToast("Paste data is empty.", "warning");
                return;
            }
            let insertedCount = 0;
            let focusRowId = null;
            updateProject((project) => {
                const table = getTableById(project, tableId);
                if (!table) return;
                const result = appendMatrixToTable(table, matrix);
                insertedCount = result.insertedCount;
                focusRowId = result.lastRowId;
                state.ui.bulkPasteDraft[tableId] = "";
                state.ui.bulkPasteOpen[tableId] = false;
                if (focusRowId) {
                    state.ui.focus = { tableId, rowId: focusRowId };
                }
            }, insertedCount ? `Pasted ${insertedCount} rows.` : "Nothing was pasted.");
            return;
        }
        case "save-filter-preset": {
            const table = getTableById(state.project, button.dataset.tableId);
            if (!table) return;
            const query = state.ui.dataSearch[table.id] || "";
            const sort = state.ui.dataSort[table.id] || { columnId: null, direction: null };
            const suggested = createPresetName(table, query, sort);
            const name = window.prompt("Preset name", suggested);
            if (!name || !name.trim()) return;
            updateProject((project) => {
                const targetTable = getTableById(project, table.id);
                if (!targetTable) return;
                targetTable.filterPresets.push({
                    id: uid("preset"),
                    name: name.trim(),
                    query,
                    sort: clone(sort),
                });
            }, "Saved the current view.");
            return;
        }
        case "apply-filter-preset": {
            const table = getTableById(state.project, button.dataset.tableId);
            const preset = table?.filterPresets?.find((item) => item.id === button.dataset.presetId);
            if (!table || !preset) return;
            state.ui.dataSearch[table.id] = preset.query || "";
            state.ui.dataSort[table.id] = clone(preset.sort || { columnId: null, direction: null });
            renderWorkspace();
            return;
        }
        case "delete-filter-preset":
            if (!window.confirm("Delete this saved view?")) return;
            updateProject((project) => {
                const table = getTableById(project, button.dataset.tableId);
                if (!table) return;
                table.filterPresets = (table.filterPresets || []).filter((preset) => preset.id !== button.dataset.presetId);
            }, "Deleted the saved view.");
            return;
        case "apply-bulk-edit": {
            const tableId = button.dataset.tableId;
            const { table, selectedMap } = getCurrentTableViewState(tableId);
            const columnId = state.ui.bulkEditColumn[tableId] || table?.columns[0]?.id || "";
            const nextValue = toCellString(state.ui.bulkEditValue[tableId] ?? "");
            const selectedIds = Object.keys(selectedMap).filter((rowId) => selectedMap[rowId]);

            if (!table || !columnId) return;
            if (!selectedIds.length) {
                showToast("Select rows first.", "warning");
                return;
            }

            updateProject((project) => {
                const targetTable = getTableById(project, tableId);
                if (!targetTable) return;
                targetTable.rows.forEach((row) => {
                    if (selectedMap[row.id]) {
                        row.cells[columnId] = nextValue;
                    }
                });
            }, `Updated ${selectedIds.length} selected rows.`);
            return;
        }
        case "clear-selected-rows":
            state.ui.selectedRows[button.dataset.tableId] = {};
            renderWorkspace();
            return;
        case "delete-selected-rows": {
            const tableId = button.dataset.tableId;
            const { table, selectedMap } = getCurrentTableViewState(tableId);
            const selectedIds = Object.keys(selectedMap).filter((rowId) => selectedMap[rowId]);
            if (!table || !selectedIds.length) {
                showToast("No selected rows to delete.", "warning");
                return;
            }
            if (!window.confirm(`Delete ${selectedIds.length} selected rows?`)) return;

            updateProject((project) => {
                const targetTable = getTableById(project, tableId);
                if (!targetTable) return;
                targetTable.rows = targetTable.rows.filter((row) => !selectedMap[row.id]);
                state.ui.selectedRows[tableId] = {};
            }, `Deleted ${selectedIds.length} selected rows.`);
            return;
        }
        case "reset-data-view":
            state.ui.dataSearch[button.dataset.tableId] = "";
            state.ui.dataSort[button.dataset.tableId] = { columnId: null, direction: null };
            renderWorkspace();
            return;
        case "reset-project":
            if (!window.confirm("현재 작업을 지우고 샘플 프로젝트로 되돌릴까요?")) return;
            replaceProject(createSeedProject(), { recordHistory: true, toast: "샘플 프로젝트로 초기화했습니다." });
            return;
        default:
            return;
    }
}

function handleChange(event) {
    const field = event.target.dataset.field;
    if (!field) return;

    if (field === "table-name") {
        updateProject((project) => {
            const table = getTableById(project, event.target.dataset.tableId);
            if (table) table.name = event.target.value.trim() || "Untitled Table";
        });
        return;
    }

    if (field === "table-note") {
        updateProject((project) => {
            const table = getTableById(project, event.target.dataset.tableId);
            if (table) table.note = event.target.value.trim();
        });
        return;
    }

    if (field.startsWith("column-")) {
        updateProject((project) => {
            const table = getTableById(project, event.target.dataset.tableId);
            const column = getColumnById(table, event.target.dataset.columnId);
            if (!column) return;

            switch (field) {
                case "column-name":
                    column.name = event.target.value.trim() || "column";
                    break;
                case "column-type":
                    column.type = normalizeType(event.target.value);
                    if (column.type !== "ENUM") {
                        column.enumValues = column.enumValues || [];
                    }
                    break;
                case "column-pk":
                    column.pk = event.target.checked;
                    if (column.pk) {
                        column.nn = true;
                        column.uq = true;
                    }
                    break;
                case "column-fk":
                    column.fk = event.target.checked;
                    if (!column.fk) {
                        column.refTableId = null;
                        column.refColumnId = null;
                        column.relationName = "";
                        column.relationCardinality = "";
                    } else if (!column.refTableId || !column.refColumnId) {
                        const guess = guessReference(project, table.id, column);
                        if (guess) {
                            column.refTableId = guess.tableId;
                            column.refColumnId = guess.columnId;
                        }
                        column.relationCardinality = normalizeRelationCardinality(column.relationCardinality, true);
                    } else if (!column.relationCardinality) {
                        column.relationCardinality = normalizeRelationCardinality(column.relationCardinality, true);
                    }
                    break;
                case "column-nn":
                    column.nn = event.target.checked;
                    break;
                case "column-uq":
                    column.uq = event.target.checked;
                    break;
                case "column-ref": {
                    const [tableId, columnId] = String(event.target.value || "").split(":");
                    column.refTableId = tableId || null;
                    column.refColumnId = columnId || null;
                    column.fk = Boolean(column.refTableId && column.refColumnId);
                    column.relationCardinality = normalizeRelationCardinality(column.relationCardinality, column.fk);
                    if (!column.fk) {
                        column.relationName = "";
                    }
                    break;
                }
                case "column-cardinality":
                    column.relationCardinality = normalizeRelationCardinality(event.target.value, column.fk);
                    break;
                case "column-relation-name":
                    column.relationName = event.target.value.trim();
                    break;
                case "column-enum":
                    column.enumValues = normalizeEnumValues(event.target.value);
                    break;
                case "column-default":
                    column.defaultValue = toCellString(event.target.value);
                    break;
                case "column-default-helper":
                    column.defaultValue =
                        event.target.value === ""
                            ? DEFAULT_HELPERS.some((helper) => helper.value === column.defaultValue)
                                ? ""
                                : column.defaultValue
                            : event.target.value;
                    break;
                case "column-description":
                    column.description = event.target.value.trim();
                    break;
                default:
                    break;
            }
        });
        return;
    }

    if (field === "row-selected") {
        const tableId = event.target.dataset.tableId;
        const rowId = event.target.dataset.rowId;
        const selected = state.ui.selectedRows[tableId] || {};
        if (event.target.checked) {
            selected[rowId] = true;
        } else {
            delete selected[rowId];
        }
        state.ui.selectedRows[tableId] = selected;
        renderWorkspace();
        return;
    }

    if (field === "select-visible-rows") {
        const tableId = event.target.dataset.tableId;
        const { rows } = getCurrentTableViewState(tableId);
        const selected = state.ui.selectedRows[tableId] || {};
        rows.forEach((row) => {
            if (event.target.checked) {
                selected[row.id] = true;
            } else {
                delete selected[row.id];
            }
        });
        state.ui.selectedRows[tableId] = selected;
        renderWorkspace();
        return;
    }

    if (field === "bulk-edit-column") {
        state.ui.bulkEditColumn[event.target.dataset.tableId] = event.target.value;
        return;
    }

    if (field === "cell-value") {
        updateProject((project) => {
            const table = getTableById(project, event.target.dataset.tableId);
            const row = table?.rows.find((item) => item.id === event.target.dataset.rowId);
            if (!row) return;
            row.cells[event.target.dataset.columnId] = toCellString(event.target.value);
        });
    }
}

function handleInput(event) {
    const field = event.target.dataset.field;
    if (field === "data-search") {
        state.ui.dataSearch[event.target.dataset.tableId] = event.target.value;
        renderWorkspace();
        return;
    }

    if (field === "bulk-paste-text") {
        state.ui.bulkPasteDraft[event.target.dataset.tableId] = event.target.value;
        return;
    }

    if (field === "bulk-edit-value") {
        state.ui.bulkEditValue[event.target.dataset.tableId] = event.target.value;
        return;
    }

    if (field === "sql-import-text") {
        state.ui.sqlImportDraft = event.target.value;
    }
}

function handleKeydown(event) {
    const mod = event.ctrlKey || event.metaKey;
    if (!mod) return;

    const key = event.key.toLowerCase();
    if (key === "z" && !event.shiftKey) {
        event.preventDefault();
        performUndo();
    } else if (key === "y" || (key === "z" && event.shiftKey)) {
        event.preventDefault();
        performRedo();
    }
}

function performUndo() {
    if (!state.history.undo.length) return;
    const previous = state.history.undo.pop();
    state.history.redo.push(clone(state.project));
    state.project = normalizeProject(previous);
    refreshDerivedState();
    persistProject();
    renderApp();
    showToast("이전 상태로 되돌렸습니다.");
}

function performRedo() {
    if (!state.history.redo.length) return;
    const next = state.history.redo.pop();
    pushUndo(state.project);
    state.project = normalizeProject(next);
    refreshDerivedState();
    persistProject();
    renderApp();
    showToast("다시 적용했습니다.");
}

function toggleSort(tableId, columnId) {
    const current = state.ui.dataSort[tableId] || { columnId: null, direction: null };
    if (current.columnId !== columnId) {
        state.ui.dataSort[tableId] = { columnId, direction: "asc" };
    } else if (current.direction === "asc") {
        state.ui.dataSort[tableId] = { columnId, direction: "desc" };
    } else {
        state.ui.dataSort[tableId] = { columnId: null, direction: null };
    }
    renderWorkspace();
}

function clearBrokenReferences(project) {
    project.tables.forEach((table) => {
        table.columns.forEach((column) => {
            if (!column.fk) return;
            const targetTable = getTableById(project, column.refTableId);
            const targetColumn = getColumnById(targetTable, column.refColumnId);
            if (!targetTable || !targetColumn) {
                column.fk = false;
                column.refTableId = null;
                column.refColumnId = null;
                column.relationName = "";
                column.relationCardinality = "";
            }
        });
        syncTableArtifacts(table);
    });
}

function handleDragStart(event) {
    const row = event.target.closest('[data-drag="column"]');
    if (!row) return;
    state.ui.dragColumn = {
        tableId: row.dataset.tableId,
        columnId: row.dataset.columnId,
    };
    row.classList.add("schema-row--dragging");
    if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = "move";
    }
}

function handleDragOver(event) {
    if (!state.ui.dragColumn) return;
    if (event.target.closest('[data-drop="column"]')) {
        event.preventDefault();
    }
}

function handleDrop(event) {
    if (!state.ui.dragColumn) return;
    const row = event.target.closest('[data-drop="column"]');
    if (!row) return;
    event.preventDefault();

    const source = state.ui.dragColumn;
    const targetTableId = row.dataset.tableId;
    const targetColumnId = row.dataset.columnId;
    if (source.tableId !== targetTableId || source.columnId === targetColumnId) return;

    updateProject((project) => {
        const table = getTableById(project, source.tableId);
        if (!table) return;
        const sourceIndex = table.columns.findIndex((column) => column.id === source.columnId);
        const targetIndex = table.columns.findIndex((column) => column.id === targetColumnId);
        if (sourceIndex < 0 || targetIndex < 0) return;
        const [moved] = table.columns.splice(sourceIndex, 1);
        table.columns.splice(targetIndex, 0, moved);
        state.ui.focus = { tableId: table.id, columnId: moved.id };
    }, "컬럼 순서를 바꿨습니다.");
}

function handleDragEnd(event) {
    const row = event.target.closest('[data-drag="column"]');
    if (row) {
        row.classList.remove("schema-row--dragging");
    }
    state.ui.dragColumn = null;
}

function handleMouseDown(event) {
    if (state.ui.view.kind !== "erd") return;
    const handle = event.target.closest("[data-erd-drag-handle]");
    const viewport = event.target.closest("[data-erd-viewport]");

    if (handle) {
        const node = handle.closest("[data-erd-node]");
        if (!node) return;
        state.ui.erdDrag = {
            tableId: handle.dataset.tableId,
            startX: event.clientX,
            startY: event.clientY,
            originX: Number.parseFloat(node.style.left),
            originY: Number.parseFloat(node.style.top),
            moved: false,
            snapshot: clone(state.project),
        };
        return;
    }

    if (viewport && !event.target.closest("[data-erd-node]")) {
        viewport.classList.add("is-panning");
        state.ui.erdPan = {
            startX: event.clientX,
            startY: event.clientY,
            panX: state.ui.erd.panX,
            panY: state.ui.erd.panY,
        };
    }
}

function handleMouseMove(event) {
    if (state.ui.erdDrag) {
        const deltaX = (event.clientX - state.ui.erdDrag.startX) / state.ui.erd.zoom;
        const deltaY = (event.clientY - state.ui.erdDrag.startY) / state.ui.erd.zoom;
        const nextX = state.ui.erdDrag.originX + deltaX;
        const nextY = state.ui.erdDrag.originY + deltaY;

        const node = dom.workspace.querySelector(`[data-erd-node="${state.ui.erdDrag.tableId}"]`);
        if (node) {
            node.style.left = `${nextX}px`;
            node.style.top = `${nextY}px`;
        }

        const table = getTableById(state.project, state.ui.erdDrag.tableId);
        if (table) {
            table.position.x = nextX;
            table.position.y = nextY;
        }
        state.ui.erdDrag.moved = true;
        drawErdLines();
        return;
    }

    if (state.ui.erdPan) {
        state.ui.erd.panX = state.ui.erdPan.panX + (event.clientX - state.ui.erdPan.startX);
        state.ui.erd.panY = state.ui.erdPan.panY + (event.clientY - state.ui.erdPan.startY);
        const canvas = document.getElementById("erd-canvas");
        if (canvas) {
            canvas.style.transform = `translate(${state.ui.erd.panX}px, ${state.ui.erd.panY}px) scale(${state.ui.erd.zoom})`;
        }
    }
}

function handleMouseUp() {
    if (state.ui.erdDrag) {
        if (state.ui.erdDrag.moved && state.ui.erdDrag.snapshot) {
            pushUndo(state.ui.erdDrag.snapshot);
            state.history.redo = [];
            refreshDerivedState();
            renderTopbar();
            if (window.lucide) {
                window.lucide.createIcons({ root: dom.topbar });
            }
        }
        persistProject();
    }
    state.ui.erdDrag = null;
    state.ui.erdPan = null;
    dom.workspace.querySelector("[data-erd-viewport]")?.classList.remove("is-panning");
}

function handleErdWheel(event) {
    if (state.ui.view.kind !== "erd") return;
    const viewport = event.target.closest("[data-erd-viewport]");
    if (!viewport) return;
    event.preventDefault();

    const rect = viewport.getBoundingClientRect();
    const worldX = (event.clientX - rect.left - state.ui.erd.panX) / state.ui.erd.zoom;
    const worldY = (event.clientY - rect.top - state.ui.erd.panY) / state.ui.erd.zoom;
    const zoomFactor = event.deltaY > 0 ? 0.92 : 1.08;
    const nextZoom = clamp(state.ui.erd.zoom * zoomFactor, 0.35, 1.9);

    state.ui.erd.panX = event.clientX - rect.left - worldX * nextZoom;
    state.ui.erd.panY = event.clientY - rect.top - worldY * nextZoom;
    state.ui.erd.zoom = nextZoom;

    const canvas = document.getElementById("erd-canvas");
    if (canvas) {
        canvas.style.transform = `translate(${state.ui.erd.panX}px, ${state.ui.erd.panY}px) scale(${state.ui.erd.zoom})`;
    }
}

function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
}

function drawErdLines() {
    const svg = document.getElementById("erd-svg");
    if (!svg) return;
    svg.innerHTML = "";

    const nodes = {};
    dom.workspace.querySelectorAll("[data-erd-node]").forEach((element) => {
        nodes[element.dataset.erdNode] = {
            element,
            x: Number.parseFloat(element.style.left),
            y: Number.parseFloat(element.style.top),
            w: element.offsetWidth,
            h: element.offsetHeight,
        };
    });

    let edgeIndex = 0;
    state.project.tables.forEach((table) => {
        table.columns.forEach((column) => {
            if (!column.fk || !column.refTableId || !column.refColumnId) return;
            const source = nodes[table.id];
            const target = nodes[column.refTableId];
            if (!source || !target) return;

            const sourceColumnElement = source.element.querySelector(`[data-erd-column="${column.id}"]`);
            const targetColumnElement = target.element.querySelector(`[data-erd-column="${column.refColumnId}"]`);
            const sourceY = source.y + (sourceColumnElement?.offsetTop || source.h / 2) + (sourceColumnElement?.offsetHeight || 0) / 2;
            const targetY = target.y + (targetColumnElement?.offsetTop || target.h / 2) + (targetColumnElement?.offsetHeight || 0) / 2;

            const pathData = buildConnectionPath(source, target, sourceY, targetY, edgeIndex);
            edgeIndex += 1;

            const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
            path.setAttribute("class", "erd-connection");
            path.setAttribute("d", pathData.curve);
            path.setAttribute("fill", "none");
            path.setAttribute("stroke", "#0891b2");
            path.setAttribute("stroke-width", "2.6");
            path.setAttribute("stroke-linecap", "round");
            path.setAttribute("stroke-opacity", "0.88");

            const arrow = document.createElementNS("http://www.w3.org/2000/svg", "path");
            arrow.setAttribute("class", "erd-connection-arrow");
            arrow.setAttribute("d", pathData.arrow);
            arrow.setAttribute("fill", "none");
            arrow.setAttribute("stroke", "#0891b2");
            arrow.setAttribute("stroke-width", "2.6");
            arrow.setAttribute("stroke-linecap", "round");
            arrow.setAttribute("stroke-linejoin", "round");

            const dot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
            dot.setAttribute("class", "erd-connection-dot");
            dot.setAttribute("cx", String(pathData.startX));
            dot.setAttribute("cy", String(pathData.startY));
            dot.setAttribute("r", "4.4");
            dot.setAttribute("fill", "#0f172a");

            svg.appendChild(path);
            svg.appendChild(arrow);
            svg.appendChild(dot);

            const relationLabel = [column.relationCardinality || null, column.relationName || null].filter(Boolean).join(" · ");
            if (relationLabel) {
                const labelX = (pathData.startX + pathData.endX + pathData.cp1x + pathData.cp2x) / 4;
                const labelY = (pathData.startY + pathData.endY) / 2 - 10;
                const labelWidth = Math.max(64, relationLabel.length * 7 + 18);
                const labelGroup = document.createElementNS("http://www.w3.org/2000/svg", "g");
                labelGroup.setAttribute("class", "erd-relation-hit");
                labelGroup.setAttribute("data-action", "edit-relation");
                labelGroup.setAttribute("data-table-id", table.id);
                labelGroup.setAttribute("data-column-id", column.id);

                const bg = document.createElementNS("http://www.w3.org/2000/svg", "rect");
                bg.setAttribute("class", "erd-relation-label");
                bg.setAttribute("data-action", "edit-relation");
                bg.setAttribute("data-table-id", table.id);
                bg.setAttribute("data-column-id", column.id);
                bg.setAttribute("x", String(labelX - labelWidth / 2));
                bg.setAttribute("y", String(labelY - 12));
                bg.setAttribute("width", String(labelWidth));
                bg.setAttribute("height", "22");
                bg.setAttribute("rx", "11");
                bg.setAttribute("fill", "rgba(255,255,255,0.94)");
                bg.setAttribute("stroke", "rgba(8,145,178,0.18)");

                const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
                text.setAttribute("class", "erd-relation-text");
                text.setAttribute("data-action", "edit-relation");
                text.setAttribute("data-table-id", table.id);
                text.setAttribute("data-column-id", column.id);
                text.setAttribute("x", String(labelX));
                text.setAttribute("y", String(labelY + 3));
                text.setAttribute("text-anchor", "middle");
                text.setAttribute("fill", "#0f172a");
                text.setAttribute("font-size", "11");
                text.setAttribute("font-family", "Noto Sans KR, sans-serif");
                text.textContent = relationLabel;

                labelGroup.appendChild(bg);
                labelGroup.appendChild(text);
                svg.appendChild(labelGroup);
            }
        });
    });
}

function buildConnectionPath(source, target, sourceY, targetY, edgeIndex = 0) {
    const lane = (edgeIndex % 5) * 24;
    const sourceRight = source.x + source.w;
    const targetLeft = target.x;
    const sourceLeft = source.x;
    const targetRight = target.x + target.w;

    let startX;
    let endX;
    let cp1x;
    let cp2x;
    let arrow;

    if (targetLeft > sourceRight + 30) {
        startX = sourceRight;
        endX = targetLeft;
        cp1x = startX + Math.max(80, (endX - startX) / 2) + lane;
        cp2x = endX - Math.max(80, (endX - startX) / 2);
        arrow = `M ${endX - 10} ${targetY - 6} L ${endX} ${targetY} L ${endX - 10} ${targetY + 6}`;
    } else if (sourceLeft > targetRight + 30) {
        startX = sourceLeft;
        endX = targetRight;
        cp1x = startX - Math.max(80, (startX - endX) / 2) - lane;
        cp2x = endX + Math.max(80, (startX - endX) / 2);
        arrow = `M ${endX + 10} ${targetY - 6} L ${endX} ${targetY} L ${endX + 10} ${targetY + 6}`;
    } else {
        const routeRight = source.x + source.w / 2 < target.x + target.w / 2;
        if (routeRight) {
            startX = sourceRight;
            endX = targetRight;
            const bow = Math.max(startX, endX) + 90 + lane;
            cp1x = bow;
            cp2x = bow;
            arrow = `M ${endX + 10} ${targetY - 6} L ${endX} ${targetY} L ${endX + 10} ${targetY + 6}`;
        } else {
            startX = sourceLeft;
            endX = targetLeft;
            const bow = Math.min(startX, endX) - 90 - lane;
            cp1x = bow;
            cp2x = bow;
            arrow = `M ${endX - 10} ${targetY - 6} L ${endX} ${targetY} L ${endX - 10} ${targetY + 6}`;
        }
    }

    return {
        curve: `M ${startX} ${sourceY} C ${cp1x} ${sourceY}, ${cp2x} ${targetY}, ${endX} ${targetY}`,
        arrow,
        startX,
        startY: sourceY,
        endX,
        endY: targetY,
        cp1x,
        cp2x,
    };
}

function autoLayoutErd() {
    updateProject((project) => {
        const levels = Object.fromEntries(project.tables.map((table) => [table.id, 0]));
        const edges = [];

        project.tables.forEach((table) => {
            table.columns.forEach((column) => {
                if (column.fk && column.refTableId && column.refColumnId) {
                    edges.push({ from: column.refTableId, to: table.id });
                }
            });
        });

        for (let step = 0; step < project.tables.length; step += 1) {
            let changed = false;
            edges.forEach((edge) => {
                if (levels[edge.to] <= levels[edge.from]) {
                    levels[edge.to] = levels[edge.from] + 1;
                    changed = true;
                }
            });
            if (!changed) break;
        }

        const grouped = {};
        Object.entries(levels).forEach(([tableId, level]) => {
            if (!grouped[level]) grouped[level] = [];
            grouped[level].push(tableId);
        });

        Object.entries(grouped)
            .sort((left, right) => Number(left[0]) - Number(right[0]))
            .forEach(([level, tableIds], columnIndex) => {
                let currentY = 120;
                tableIds.forEach((tableId) => {
                    const table = getTableById(project, tableId);
                    if (!table) return;
                    table.position.x = 120 + columnIndex * 420;
                    table.position.y = currentY;
                    currentY += 140 + table.columns.length * 42;
                });
            });
    }, "ERD를 자동 정렬했습니다.");
    fitErdView();
}

function fitErdView() {
    const viewport = dom.workspace.querySelector("[data-erd-viewport]");
    if (!viewport || state.project.tables.length === 0) return;

    const minX = Math.min(...state.project.tables.map((table) => table.position.x));
    const minY = Math.min(...state.project.tables.map((table) => table.position.y));
    const maxX = Math.max(...state.project.tables.map((table) => table.position.x + 320));
    const maxY = Math.max(...state.project.tables.map((table) => table.position.y + 180 + table.columns.length * 34));
    const width = maxX - minX;
    const height = maxY - minY;

    const padding = 80;
    const zoom = clamp(Math.min((viewport.clientWidth - padding * 2) / width, (viewport.clientHeight - padding * 2) / height), 0.4, 1.2);

    state.ui.erd.zoom = zoom;
    state.ui.erd.panX = padding - minX * zoom;
    state.ui.erd.panY = padding - minY * zoom;
    const canvas = document.getElementById("erd-canvas");
    if (canvas) {
        canvas.style.transform = `translate(${state.ui.erd.panX}px, ${state.ui.erd.panY}px) scale(${state.ui.erd.zoom})`;
    }
}

function copyErdSvg() {
    if (state.ui.view.kind !== "erd") {
        showToast("ERD 화면에서만 SVG를 복사할 수 있습니다.", "warning");
        return;
    }

    const svgText = buildSvgMarkup();
    navigator.clipboard
        .writeText(svgText)
        .then(() => showToast("ERD SVG를 클립보드에 복사했습니다."))
        .catch(() => showToast("SVG 복사에 실패했습니다.", "error"));
}

function copyProjectSql() {
    const sqlText = buildProjectSql();
    copyText(sqlText, "Copied SQL DDL.", "Failed to copy SQL.");
}

function exportSql() {
    const blob = new Blob([buildProjectSql()], { type: "text/sql;charset=utf-8" });
    triggerDownload(blob, "data-table-project.sql");
    showToast("Exported SQL DDL.");
}

function copyText(text, successMessage, errorMessage) {
    if (navigator.clipboard?.writeText) {
        navigator.clipboard
            .writeText(text)
            .then(() => showToast(successMessage))
            .catch(() => fallbackCopyText(text, successMessage, errorMessage));
        return;
    }
    fallbackCopyText(text, successMessage, errorMessage);
}

function fallbackCopyText(text, successMessage, errorMessage) {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    try {
        document.execCommand("copy");
        showToast(successMessage);
    } catch (_error) {
        showToast(errorMessage, "error");
    } finally {
        textarea.remove();
    }
}

function parseSqlTables(sqlText, baseProject = { tables: [] }) {
    const source = String(sqlText ?? "");
    const blocks = extractCreateTableBlocks(source);
    if (!blocks.length) {
        return [];
    }

    const metadata = extractSqlImportMetadata(source);
    const existingTables = Array.isArray(baseProject?.tables) ? baseProject.tables : [];
    const importedTables = [];
    const pendingReferences = [];

    blocks.forEach((block, index) => {
        const sourceTableKey = normalizeKey(block.tableName) || `importedtable${index + 1}`;
        const parsedDefinitions = splitSqlDefinitions(block.body)
            .map((definition) => parseSqlDefinition(definition))
            .filter(Boolean);

        const columns = [];
        const tableConstraints = [];

        parsedDefinitions.forEach((parsed, definitionIndex) => {
            if (parsed.kind !== "column") {
                tableConstraints.push(parsed);
                return;
            }

            const normalizedType = normalizeImportedSqlType(parsed.type, parsed.enumValues);
            const baseColumnName = parsed.name || `column_${definitionIndex + 1}`;
            const columnName = nextColumnName({ columns }, baseColumnName);
            const column = createColumn({
                name: columnName,
                type: normalizedType.type,
                pk: parsed.pk,
                nn: parsed.nn,
                uq: parsed.uq,
                fk: Boolean(parsed.reference),
                defaultValue: parsed.defaultValue || normalizedType.defaultValue || "",
                enumValues: normalizedType.enumValues,
            });
            column._sqlSourceKey = normalizeKey(parsed.name || columnName);

            if (parsed.reference) {
                pendingReferences.push({
                    sourceTableKey,
                    sourceColumnId: column.id,
                    refTableName: parsed.reference.tableName,
                    refColumnName: parsed.reference.columnNames[0] || "id",
                });
            }

            columns.push(column);
        });

        if (!columns.length) {
            const idColumn = createColumn({
                name: "id",
                type: "INT",
                pk: true,
                nn: true,
                uq: true,
            });
            idColumn._sqlSourceKey = "id";
            columns.push(idColumn);
        }

        const table = {
            id: uid("table"),
            name: nextTableName({ tables: [...existingTables, ...importedTables] }, block.tableName || `Imported Table ${index + 1}`),
            note: metadata.notes.get(sourceTableKey) || "Imported from SQL",
            position: nextTablePosition(existingTables.length + importedTables.length),
            columns,
            rows: [],
            filterPresets: [],
            _sqlSourceKey: sourceTableKey,
        };

        tableConstraints.forEach((constraint) => {
            applyParsedTableConstraint(table, constraint, pendingReferences);
        });

        table.columns.forEach((column) => {
            const sourceColumnKey = column._sqlSourceKey || normalizeKey(column.name);
            const metaKey = `${sourceTableKey}.${sourceColumnKey}`;
            const helperDefault = metadata.helpers.get(metaKey);
            const relationMeta = metadata.relations.get(metaKey);

            if (helperDefault) {
                column.defaultValue = helperDefault;
            }

            if (relationMeta) {
                column.fk = true;
                column.relationCardinality = normalizeRelationCardinality(relationMeta.cardinality, true);
                column.relationName = relationMeta.name;
            }
        });

        importedTables.push(table);
    });

    resolvePendingSqlReferences(importedTables, existingTables, pendingReferences);

    return normalizeProject({
        version: 2,
        updatedAt: new Date().toISOString(),
        tables: importedTables,
    }).tables;
}

function extractSqlImportMetadata(sqlText) {
    const notes = new Map();
    const helpers = new Map();
    const relations = new Map();

    String(sqlText ?? "")
        .split(/\r?\n/)
        .forEach((line) => {
            const trimmed = line.trim();

            let match = trimmed.match(/^--\s*note\s+(.+?)\s*:\s*(.+)$/i);
            if (match) {
                notes.set(normalizeKey(stripSqlIdentifier(match[1])), match[2].trim());
                return;
            }

            match = trimmed.match(/^--\s*helper\s+(.+?)\.(.+?)\s*:\s*(.+)$/i);
            if (match) {
                helpers.set(
                    `${normalizeKey(stripSqlIdentifier(match[1]))}.${normalizeKey(stripSqlIdentifier(match[2]))}`,
                    match[3].trim(),
                );
                return;
            }

            match = trimmed.match(/^--\s*relation\s+(.+?)\.(.+?)\s*:\s*(.+)$/i);
            if (match) {
                const detail = match[3].trim();
                const cardinality = CARDINALITY_OPTIONS.find((option) => detail.toUpperCase().includes(option)) || "";
                const name = detail.replace(cardinality, "").replace(/^[\s·\-•|]+|[\s·\-•|]+$/g, "").trim();
                relations.set(`${normalizeKey(stripSqlIdentifier(match[1]))}.${normalizeKey(stripSqlIdentifier(match[2]))}`, {
                    cardinality,
                    name,
                });
            }
        });

    return { notes, helpers, relations };
}

function extractCreateTableBlocks(sqlText) {
    const source = String(sqlText ?? "");
    const blocks = [];
    const createTablePattern = /CREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?/gi;

    let match;
    while ((match = createTablePattern.exec(source))) {
        let cursor = match.index + match[0].length;
        while (cursor < source.length && /\s/.test(source[cursor])) {
            cursor += 1;
        }

        const nameStart = cursor;
        let inDoubleQuotes = false;
        let inBackticks = false;
        let inBrackets = false;

        while (cursor < source.length) {
            const char = source[cursor];
            if (char === '"' && !inBackticks && !inBrackets) {
                if (inDoubleQuotes && source[cursor + 1] === '"') {
                    cursor += 2;
                    continue;
                }
                inDoubleQuotes = !inDoubleQuotes;
                cursor += 1;
                continue;
            }
            if (char === "`" && !inDoubleQuotes && !inBrackets) {
                inBackticks = !inBackticks;
                cursor += 1;
                continue;
            }
            if (char === "[" && !inDoubleQuotes && !inBackticks) {
                inBrackets = true;
                cursor += 1;
                continue;
            }
            if (char === "]" && inBrackets) {
                inBrackets = false;
                cursor += 1;
                continue;
            }
            if (!inDoubleQuotes && !inBackticks && !inBrackets && char === "(") {
                break;
            }
            cursor += 1;
        }

        if (source[cursor] !== "(") {
            continue;
        }

        const tableName = stripSqlIdentifier(source.slice(nameStart, cursor).trim());
        const openIndex = cursor;
        let depth = 0;
        inDoubleQuotes = false;
        inBackticks = false;
        inBrackets = false;
        let closeIndex = -1;

        for (let index = openIndex; index < source.length; index += 1) {
            const char = source[index];
            if (char === '"' && !inBackticks && !inBrackets) {
                if (inDoubleQuotes && source[index + 1] === '"') {
                    index += 1;
                    continue;
                }
                inDoubleQuotes = !inDoubleQuotes;
                continue;
            }
            if (char === "`" && !inDoubleQuotes && !inBrackets) {
                inBackticks = !inBackticks;
                continue;
            }
            if (char === "[" && !inDoubleQuotes && !inBackticks) {
                inBrackets = true;
                continue;
            }
            if (char === "]" && inBrackets) {
                inBrackets = false;
                continue;
            }
            if (inDoubleQuotes || inBackticks || inBrackets) {
                continue;
            }
            if (char === "(") {
                depth += 1;
            } else if (char === ")") {
                depth -= 1;
                if (depth === 0) {
                    closeIndex = index;
                    break;
                }
            }
        }

        if (closeIndex < 0) {
            throw new Error(`Unclosed CREATE TABLE block for ${tableName || "table"}.`);
        }

        blocks.push({
            tableName: tableName || `Imported Table ${blocks.length + 1}`,
            body: source.slice(openIndex + 1, closeIndex),
        });
        createTablePattern.lastIndex = closeIndex + 1;
    }

    return blocks;
}

function splitSqlDefinitions(body) {
    const source = String(body ?? "");
    const definitions = [];
    let chunk = "";
    let depth = 0;
    let inSingleQuotes = false;
    let inDoubleQuotes = false;
    let inBackticks = false;
    let inBrackets = false;

    for (let index = 0; index < source.length; index += 1) {
        const char = source[index];
        const next = source[index + 1];

        if (char === "'" && !inDoubleQuotes && !inBackticks && !inBrackets) {
            chunk += char;
            if (inSingleQuotes && next === "'") {
                chunk += next;
                index += 1;
            } else {
                inSingleQuotes = !inSingleQuotes;
            }
            continue;
        }

        if (!inSingleQuotes && char === '"' && !inBackticks && !inBrackets) {
            chunk += char;
            if (inDoubleQuotes && next === '"') {
                chunk += next;
                index += 1;
            } else {
                inDoubleQuotes = !inDoubleQuotes;
            }
            continue;
        }

        if (!inSingleQuotes && !inDoubleQuotes && char === "`" && !inBrackets) {
            inBackticks = !inBackticks;
            chunk += char;
            continue;
        }

        if (!inSingleQuotes && !inDoubleQuotes && !inBackticks && char === "[") {
            inBrackets = true;
            chunk += char;
            continue;
        }

        if (!inSingleQuotes && !inDoubleQuotes && !inBackticks && char === "]") {
            inBrackets = false;
            chunk += char;
            continue;
        }

        if (!inSingleQuotes && !inDoubleQuotes && !inBackticks && !inBrackets) {
            if (char === "(") {
                depth += 1;
            } else if (char === ")") {
                depth = Math.max(0, depth - 1);
            } else if (char === "," && depth === 0) {
                if (chunk.trim()) {
                    definitions.push(chunk.trim());
                }
                chunk = "";
                continue;
            }
        }

        chunk += char;
    }

    if (chunk.trim()) {
        definitions.push(chunk.trim());
    }

    return definitions;
}

function parseSqlDefinition(definition) {
    const trimmed = String(definition ?? "").trim().replace(/,$/, "");
    if (!trimmed) return null;

    const withoutConstraintName = trimmed.replace(/^CONSTRAINT\s+(?:"(?:[^"]|"")+"|`[^`]+`|\[[^\]]+\]|[^\s]+)\s+/i, "");

    let match = withoutConstraintName.match(/^PRIMARY\s+KEY\s*\(([^)]+)\)/i);
    if (match) {
        return {
            kind: "primaryKey",
            columns: parseSqlIdentifierList(match[1]),
        };
    }

    match = withoutConstraintName.match(/^UNIQUE\s*\(([^)]+)\)/i);
    if (match) {
        return {
            kind: "unique",
            columns: parseSqlIdentifierList(match[1]),
        };
    }

    match = withoutConstraintName.match(/^FOREIGN\s+KEY\s*\(([^)]+)\)\s+REFERENCES\s+(.+?)\s*\(([^)]+)\)/i);
    if (match) {
        return {
            kind: "foreignKey",
            columns: parseSqlIdentifierList(match[1]),
            refTableName: stripSqlIdentifier(match[2]),
            refColumnNames: parseSqlIdentifierList(match[3]),
        };
    }

    return parseSqlColumnDefinition(trimmed);
}

function parseSqlColumnDefinition(definition) {
    const { token: nameToken, rest } = readSqlHeadToken(definition);
    if (!nameToken) return null;

    const name = stripSqlIdentifier(nameToken);
    const constraintIndex = findSqlConstraintIndex(rest);
    const rawType = (constraintIndex >= 0 ? rest.slice(0, constraintIndex) : rest).trim() || "TEXT";
    const constraintTail = (constraintIndex >= 0 ? rest.slice(constraintIndex) : "").trim();

    const enumMatch = rawType.match(/^ENUM\s*\(([\s\S]+)\)$/i);
    const checkEnumMatch = constraintTail.match(/\bCHECK\b\s*\(\s*.+?\s+IN\s*\(([\s\S]+?)\)\s*\)/i);
    const enumValues = enumMatch ? parseSqlLiteralList(enumMatch[1]) : checkEnumMatch ? parseSqlLiteralList(checkEnumMatch[1]) : [];

    const defaultMatch = constraintTail.match(
        /\bDEFAULT\b\s+(.+?)(?=\s+(?:NOT\s+NULL|NULL|PRIMARY\s+KEY|UNIQUE|CHECK|REFERENCES|CONSTRAINT|GENERATED\s+(?:ALWAYS|BY\s+DEFAULT))\b|$)/i,
    );
    const referenceMatch = constraintTail.match(/\bREFERENCES\b\s+(.+?)\s*\(([^)]+)\)/i);

    let defaultValue = defaultMatch ? parseSqlDefaultValue(defaultMatch[1]) : "";
    if (!defaultValue && /\b(?:SERIAL|GENERATED\s+(?:ALWAYS|BY\s+DEFAULT)\s+AS\s+IDENTITY)\b/i.test(`${rawType} ${constraintTail}`)) {
        defaultValue = "{{autoincrement}}";
    }

    return {
        kind: "column",
        name,
        type: rawType,
        pk: /\bPRIMARY\s+KEY\b/i.test(constraintTail),
        nn: /\bNOT\s+NULL\b/i.test(constraintTail),
        uq: /\bUNIQUE\b/i.test(constraintTail),
        defaultValue,
        enumValues,
        reference: referenceMatch
            ? {
                  tableName: stripSqlIdentifier(referenceMatch[1]),
                  columnNames: parseSqlIdentifierList(referenceMatch[2]),
              }
            : null,
    };
}

function readSqlHeadToken(text) {
    const source = String(text ?? "").trim();
    if (!source) {
        return { token: "", rest: "" };
    }

    const first = source[0];
    if (first === '"') {
        let token = '"';
        for (let index = 1; index < source.length; index += 1) {
            token += source[index];
            if (source[index] === '"' && source[index + 1] === '"') {
                token += source[index + 1];
                index += 1;
                continue;
            }
            if (source[index] === '"') {
                return {
                    token,
                    rest: source.slice(index + 1).trim(),
                };
            }
        }
    }

    if (first === "`") {
        const end = source.indexOf("`", 1);
        if (end > 0) {
            return {
                token: source.slice(0, end + 1),
                rest: source.slice(end + 1).trim(),
            };
        }
    }

    if (first === "[") {
        const end = source.indexOf("]", 1);
        if (end > 0) {
            return {
                token: source.slice(0, end + 1),
                rest: source.slice(end + 1).trim(),
            };
        }
    }

    const whitespaceIndex = source.search(/\s/);
    if (whitespaceIndex < 0) {
        return { token: source, rest: "" };
    }

    return {
        token: source.slice(0, whitespaceIndex),
        rest: source.slice(whitespaceIndex).trim(),
    };
}

function findSqlConstraintIndex(text) {
    const source = String(text ?? "");
    const patterns = [
        /\bNOT\s+NULL\b/i,
        /\bNULL\b/i,
        /\bPRIMARY\s+KEY\b/i,
        /\bUNIQUE\b/i,
        /\bDEFAULT\b/i,
        /\bCHECK\b/i,
        /\bREFERENCES\b/i,
        /\bCONSTRAINT\b/i,
        /\bGENERATED\s+(?:ALWAYS|BY\s+DEFAULT)\b/i,
    ];

    let best = -1;
    patterns.forEach((pattern) => {
        const match = pattern.exec(source);
        if (!match) return;
        if (best === -1 || match.index < best) {
            best = match.index;
        }
    });

    return best;
}

function parseSqlIdentifierList(text) {
    return splitSqlDefinitions(text).map((identifier) => stripSqlIdentifier(identifier)).filter(Boolean);
}

function parseSqlLiteralList(text) {
    return splitSqlDefinitions(text).map((value) => parseSqlLiteralValue(value)).filter((value) => value !== "");
}

function parseSqlLiteralValue(value) {
    const trimmed = String(value ?? "").trim();
    if (!trimmed) return "";

    if ((trimmed.startsWith("'") && trimmed.endsWith("'")) || (trimmed.startsWith('"') && trimmed.endsWith('"'))) {
        return trimmed.slice(1, -1).replace(/''/g, "'").replace(/""/g, '"');
    }

    return trimmed;
}

function parseSqlDefaultValue(value) {
    let normalized = String(value ?? "").trim().replace(/,$/, "");
    while (normalized.startsWith("(") && normalized.endsWith(")")) {
        normalized = normalized.slice(1, -1).trim();
    }

    if (!normalized || /^NULL$/i.test(normalized)) return "";
    if (/^CURRENT_DATE$/i.test(normalized)) return "{{today}}";
    if (/^(CURRENT_TIMESTAMP|NOW\(\)|LOCALTIMESTAMP)$/i.test(normalized)) return "{{now}}";
    if (/^(TRUE|FALSE)$/i.test(normalized)) return `{{${normalized.toLowerCase()}}}`;
    if (/^(gen_random_uuid\(\)|uuid_generate_v4\(\)|uuid\(\))$/i.test(normalized)) return "{{uuid}}";

    return parseSqlLiteralValue(normalized);
}

function normalizeImportedSqlType(type, enumValues = []) {
    const raw = String(type ?? "TEXT").trim();
    const upper = raw.toUpperCase();
    const nativeEnumMatch = raw.match(/^ENUM\s*\(([\s\S]+)\)$/i);
    const resolvedEnumValues = enumValues.length ? enumValues : nativeEnumMatch ? parseSqlLiteralList(nativeEnumMatch[1]) : [];

    if (resolvedEnumValues.length) {
        return {
            type: "ENUM",
            enumValues: resolvedEnumValues,
            defaultValue: "",
        };
    }

    if (/\bJSONB?\b/.test(upper)) {
        return { type: "JSON", enumValues: [], defaultValue: "" };
    }
    if (/\bBOOL(?:EAN)?\b/.test(upper)) {
        return { type: "BOOLEAN", enumValues: [], defaultValue: "" };
    }
    if (/\bDATE\b/.test(upper) && !/\bTIME\b/.test(upper)) {
        return { type: "DATE", enumValues: [], defaultValue: "" };
    }
    if (/\b(?:INT|INTEGER|BIGINT|SMALLINT|TINYINT|SERIAL|BIGSERIAL|SMALLSERIAL)\b/.test(upper)) {
        return {
            type: "INT",
            enumValues: [],
            defaultValue: /\bSERIAL\b/.test(upper) ? "{{autoincrement}}" : "",
        };
    }
    if (/\b(?:FLOAT|DOUBLE|DECIMAL|NUMERIC|REAL)\b/.test(upper)) {
        return { type: "FLOAT", enumValues: [], defaultValue: "" };
    }

    return { type: "STRING", enumValues: [], defaultValue: "" };
}

function stripSqlIdentifier(identifier) {
    const source = String(identifier ?? "").trim();
    const finalSegment = source.split(".").pop()?.trim() || source;

    if (finalSegment.startsWith('"') && finalSegment.endsWith('"')) {
        return finalSegment.slice(1, -1).replace(/""/g, '"');
    }
    if (finalSegment.startsWith("`") && finalSegment.endsWith("`")) {
        return finalSegment.slice(1, -1);
    }
    if (finalSegment.startsWith("[") && finalSegment.endsWith("]")) {
        return finalSegment.slice(1, -1);
    }

    return finalSegment;
}

function applyParsedTableConstraint(table, constraint, pendingReferences) {
    const findColumn = (name) =>
        table.columns.find((column) => {
            const sourceKey = column._sqlSourceKey || normalizeKey(column.name);
            return sourceKey === normalizeKey(name) || normalizeKey(column.name) === normalizeKey(name);
        }) || null;

    if (constraint.kind === "primaryKey") {
        constraint.columns.forEach((name) => {
            const column = findColumn(name);
            if (!column) return;
            column.pk = true;
            column.nn = true;
            column.uq = true;
        });
        return;
    }

    if (constraint.kind === "unique" && constraint.columns.length === 1) {
        const column = findColumn(constraint.columns[0]);
        if (column) {
            column.uq = true;
        }
        return;
    }

    if (constraint.kind === "foreignKey" && constraint.columns.length === 1) {
        const column = findColumn(constraint.columns[0]);
        if (!column) return;
        column.fk = true;
        pendingReferences.push({
            sourceTableKey: table._sqlSourceKey || normalizeKey(table.name),
            sourceColumnId: column.id,
            refTableName: constraint.refTableName,
            refColumnName: constraint.refColumnNames[0] || "id",
        });
    }
}

function resolvePendingSqlReferences(importedTables, existingTables, pendingReferences) {
    const sourceTableLookup = new Map(importedTables.map((table) => [table._sqlSourceKey || normalizeKey(table.name), table]));
    const targetTableLookup = new Map();

    [...existingTables, ...importedTables].forEach((table) => {
        const primaryKey = table._sqlSourceKey || normalizeKey(table.name);
        if (!targetTableLookup.has(primaryKey)) {
            targetTableLookup.set(primaryKey, table);
        }
        const fallbackKey = normalizeKey(table.name);
        if (!targetTableLookup.has(fallbackKey)) {
            targetTableLookup.set(fallbackKey, table);
        }
    });

    pendingReferences.forEach((reference) => {
        const sourceTable = sourceTableLookup.get(reference.sourceTableKey);
        const sourceColumn = sourceTable?.columns.find((column) => column.id === reference.sourceColumnId);
        const targetTable =
            targetTableLookup.get(normalizeKey(reference.refTableName)) ||
            [...existingTables, ...importedTables].find((table) => normalizeKey(table.name) === normalizeKey(reference.refTableName));

        const targetColumn =
            targetTable?.columns.find((column) => {
                const sourceKey = column._sqlSourceKey || normalizeKey(column.name);
                return sourceKey === normalizeKey(reference.refColumnName) || normalizeKey(column.name) === normalizeKey(reference.refColumnName);
            }) || null;

        if (!sourceColumn || !targetTable || !targetColumn) return;

        sourceColumn.fk = true;
        sourceColumn.refTableId = targetTable.id;
        sourceColumn.refColumnId = targetColumn.id;
        sourceColumn.relationCardinality = normalizeRelationCardinality(sourceColumn.relationCardinality, true);
    });
}

function buildProjectSql() {
    const lines = [
        "-- Generated by Data Table Editor",
        `-- ${new Date().toISOString()}`,
        "",
    ];

    state.project.tables.forEach((table) => {
        const pkColumns = table.columns.filter((column) => column.pk);
        const fkColumns = table.columns.filter((column) => column.fk && column.refTableId && column.refColumnId);
        const columnLines = table.columns.map((column) => `  ${buildSqlColumnDefinition(column)}`);

        if (pkColumns.length) {
            columnLines.push(`  PRIMARY KEY (${pkColumns.map((column) => quoteSqlIdentifier(column.name)).join(", ")})`);
        }

        fkColumns.forEach((column) => {
            const targetTable = getTableById(state.project, column.refTableId);
            const targetColumn = getColumnById(targetTable, column.refColumnId);
            if (!targetTable || !targetColumn) return;
            columnLines.push(
                `  CONSTRAINT ${quoteSqlIdentifier(`fk_${table.name}_${column.name}`)} FOREIGN KEY (${quoteSqlIdentifier(column.name)}) REFERENCES ${quoteSqlIdentifier(targetTable.name)} (${quoteSqlIdentifier(targetColumn.name)})`,
            );
        });

        lines.push(`CREATE TABLE ${quoteSqlIdentifier(table.name)} (`);
        lines.push(columnLines.join(",\n"));
        lines.push(");");

        fkColumns.forEach((column) => {
            const relationDetail = [column.relationCardinality || null, column.relationName || null].filter(Boolean).join(" · ");
            if (relationDetail) {
                lines.push(`-- relation ${table.name}.${column.name}: ${relationDetail}`);
            }
        });

        table.columns.forEach((column) => {
            const helperComment = buildSqlHelperComment(table, column);
            if (helperComment) {
                lines.push(helperComment);
            }
        });

        if (table.note) {
            lines.push(`-- note ${table.name}: ${table.note}`);
        }

        lines.push("");
    });

    return lines.join("\n");
}

function buildSqlColumnDefinition(column) {
    const parts = [quoteSqlIdentifier(column.name), mapSqlType(column)];
    if (column.nn || column.pk) parts.push("NOT NULL");
    if (column.uq && !column.pk) parts.push("UNIQUE");

    const defaultClause = buildSqlDefaultClause(column);
    if (defaultClause) parts.push(defaultClause);

    if (column.type === "ENUM" && column.enumValues.length) {
        const enumSet = column.enumValues.map((value) => quoteSqlLiteral(value)).join(", ");
        parts.push(`CHECK (${quoteSqlIdentifier(column.name)} IN (${enumSet}))`);
    }

    return parts.join(" ");
}

function mapSqlType(column) {
    switch (column.type) {
        case "INT":
            return "INTEGER";
        case "FLOAT":
            return "REAL";
        case "BOOLEAN":
            return "BOOLEAN";
        case "DATE":
            return "DATE";
        case "JSON":
            return "JSON";
        default:
            return "TEXT";
    }
}

function buildSqlDefaultClause(column) {
    const value = String(column.defaultValue ?? "").trim();
    if (!value) return "";

    if (value === "{{today}}") return "DEFAULT CURRENT_DATE";
    if (value === "{{now}}") return "DEFAULT CURRENT_TIMESTAMP";
    if (value === "{{true}}") return "DEFAULT TRUE";
    if (value === "{{false}}") return "DEFAULT FALSE";
    if (value === "{{uuid}}" || value === "{{autoincrement}}") return "";

    if ((column.type === "INT" || column.type === "FLOAT") && /^-?\d+(\.\d+)?$/.test(value)) {
        return `DEFAULT ${value}`;
    }
    if (column.type === "BOOLEAN" && /^(true|1|yes|y)$/i.test(value)) {
        return "DEFAULT TRUE";
    }
    if (column.type === "BOOLEAN" && /^(false|0|no|n)$/i.test(value)) {
        return "DEFAULT FALSE";
    }
    if (/^(CURRENT_TIMESTAMP|CURRENT_DATE|NOW\(\))$/i.test(value)) {
        return `DEFAULT ${value.toUpperCase()}`;
    }
    return `DEFAULT ${quoteSqlLiteral(value)}`;
}

function buildSqlHelperComment(table, column) {
    const value = String(column.defaultValue ?? "").trim();
    if (!/^\{\{.+\}\}$/.test(value)) return "";
    return `-- helper ${table.name}.${column.name}: ${value}`;
}

function quoteSqlIdentifier(value) {
    return `"${String(value ?? "").replace(/"/g, '""')}"`;
}

function quoteSqlLiteral(value) {
    return `'${String(value ?? "").replace(/'/g, "''")}'`;
}

function buildSvgMarkup() {
    const width = 3200;
    const height = 2200;
    const pieces = [`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" fill="none">`];

    state.project.tables.forEach((table) => {
        table.columns.forEach((column, index) => {
            if (!column.fk || !column.refTableId || !column.refColumnId) return;
            const targetTable = getTableById(state.project, column.refTableId);
            if (!targetTable) return;
            const sourceY = table.position.y + 90 + index * 34;
            const targetIndex = targetTable.columns.findIndex((item) => item.id === column.refColumnId);
            const targetY = targetTable.position.y + 90 + Math.max(targetIndex, 0) * 34;
            const path = buildConnectionPath(
                { x: table.position.x, y: table.position.y, w: 300 },
                { x: targetTable.position.x, y: targetTable.position.y, w: 300 },
                sourceY,
                targetY,
            );
            pieces.push(`<path d="${path.curve}" stroke="#0891b2" stroke-width="2.6" stroke-linecap="round"/>`);
            pieces.push(`<path d="${path.arrow}" stroke="#0891b2" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>`);
            const relationLabel = [column.relationCardinality || null, column.relationName || null].filter(Boolean).join(" · ");
            if (relationLabel) {
                const labelX = (path.startX + path.endX + path.cp1x + path.cp2x) / 4;
                const labelY = (path.startY + path.endY) / 2 - 10;
                const labelWidth = Math.max(64, relationLabel.length * 7 + 18);
                pieces.push(`<rect x="${labelX - labelWidth / 2}" y="${labelY - 12}" width="${labelWidth}" height="22" rx="11" fill="rgba(255,255,255,0.94)" stroke="rgba(8,145,178,0.18)"/>`);
                pieces.push(`<text x="${labelX}" y="${labelY + 3}" text-anchor="middle" fill="#0f172a" font-size="11" font-family="Noto Sans KR, sans-serif">${escapeHtml(relationLabel)}</text>`);
            }
        });
    });

    state.project.tables.forEach((table) => {
        const nodeHeight = 80 + table.columns.length * 34;
        pieces.push(`<g transform="translate(${table.position.x}, ${table.position.y})">`);
        pieces.push(`<rect width="300" height="${nodeHeight}" rx="22" fill="#ffffff" stroke="#c2d3de" stroke-width="1.5"/>`);
        pieces.push(`<rect width="300" height="64" rx="22" fill="#eef6fa"/>`);
        pieces.push(`<text x="18" y="38" fill="#0f172a" font-size="18" font-family="Space Grotesk, sans-serif" font-weight="700">${escapeHtml(table.name)}</text>`);
        table.columns.forEach((column, index) => {
            const y = 92 + index * 34;
            const badge = column.pk ? "PK" : column.fk ? "FK" : "";
            if (badge) {
                const fill = column.pk ? "#fef3c7" : "#dbeafe";
                const text = column.pk ? "#a16207" : "#0369a1";
                pieces.push(`<rect x="18" y="${y - 14}" width="28" height="18" rx="9" fill="${fill}"/>`);
                pieces.push(`<text x="32" y="${y - 1}" text-anchor="middle" fill="${text}" font-size="10" font-family="Noto Sans KR, sans-serif" font-weight="700">${badge}</text>`);
                pieces.push(`<text x="56" y="${y}" fill="#0f172a" font-size="13" font-family="Noto Sans KR, sans-serif">${escapeHtml(column.name)}</text>`);
            } else {
                pieces.push(`<text x="18" y="${y}" fill="#0f172a" font-size="13" font-family="Noto Sans KR, sans-serif">${escapeHtml(column.name)}</text>`);
            }
            pieces.push(`<text x="282" y="${y}" text-anchor="end" fill="#64748b" font-size="11" font-family="Space Grotesk, sans-serif">${escapeHtml(column.type)}</text>`);
        });
        pieces.push("</g>");
    });

    pieces.push("</svg>");
    return pieces.join("");
}

function exportJson() {
    const blob = new Blob([JSON.stringify(state.project, null, 2)], { type: "application/json;charset=utf-8" });
    triggerDownload(blob, "data-table-project.json");
    showToast("프로젝트 JSON을 내보냈습니다.");
}

function exportCsv(tableId) {
    const table = getTableById(state.project, tableId);
    if (!table) {
        showToast("내보낼 테이블을 찾지 못했습니다.", "warning");
        return;
    }

    const lines = [];
    lines.push(table.columns.map((column) => escapeCsv(column.name)).join(","));
    table.rows.forEach((row) => {
        lines.push(
            table.columns
                .map((column) => escapeCsv(toCellString(row.cells[column.id] ?? "")))
                .join(","),
        );
    });
    const blob = new Blob([`\uFEFF${lines.join("\n")}`], { type: "text/csv;charset=utf-8" });
    triggerDownload(blob, `${table.name}.csv`);
    showToast(`${table.name} CSV를 내보냈습니다.`);
}

function escapeCsv(value) {
    const text = String(value ?? "");
    if (/[",\n]/.test(text)) {
        return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
}

function exportExcel() {
    try {
        const workbook = XLSX.utils.book_new();
        state.project.tables.forEach((table) => {
            const sheetData = [
                table.columns.map((column) => column.name),
                ...table.rows.map((row) => table.columns.map((column) => row.cells[column.id] ?? "")),
            ];
            const sheet = XLSX.utils.aoa_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(workbook, sheet, table.name.slice(0, 31) || "Sheet");
        });
        XLSX.writeFile(workbook, "data-table-project.xlsx");
        showToast("엑셀 파일을 내보냈습니다.");
    } catch (_error) {
        showToast("엑셀 내보내기에 실패했습니다.", "error");
    }
}

function parseDelimitedMatrix(text) {
    const source = String(text ?? "");
    if (!source.trim()) return [];
    const delimiter = detectDelimiter(source);
    const rows = [];
    let cell = "";
    let row = [];
    let inQuotes = false;

    for (let index = 0; index < source.length; index += 1) {
        const char = source[index];
        const next = source[index + 1];

        if (char === '"') {
            if (inQuotes && next === '"') {
                cell += '"';
                index += 1;
            } else {
                inQuotes = !inQuotes;
            }
            continue;
        }

        if (!inQuotes && char === delimiter) {
            row.push(cell);
            cell = "";
            continue;
        }

        if (!inQuotes && (char === "\n" || char === "\r")) {
            if (char === "\r" && next === "\n") {
                index += 1;
            }
            row.push(cell);
            rows.push(row);
            row = [];
            cell = "";
            continue;
        }

        cell += char;
    }

    row.push(cell);
    rows.push(row);

    return rows
        .map((cells) => cells.map((value) => String(value ?? "").trim()))
        .filter((cells) => cells.some((value) => value !== ""));
}

function detectDelimiter(text) {
    const sample = String(text ?? "")
        .split(/\r?\n/)
        .find((line) => line.trim());
    if (!sample) return "\t";
    const tabCount = (sample.match(/\t/g) || []).length;
    const commaCount = (sample.match(/,/g) || []).length;
    return tabCount >= commaCount ? "\t" : ",";
}

function appendMatrixToTable(table, matrix) {
    const normalizedColumnNames = table.columns.map((column) => normalizeKey(column.name));
    const firstRow = matrix[0] || [];
    const headerMapping = firstRow.map((cell) => normalizedColumnNames.indexOf(normalizeKey(cell)));
    const headerMatches = headerMapping.filter((index) => index >= 0).length;
    const useHeader = headerMatches > 0 && headerMatches >= Math.ceil(Math.min(firstRow.length, table.columns.length) / 2);
    const dataRows = useHeader ? matrix.slice(1) : matrix;
    let insertedCount = 0;
    let lastRowId = null;

    dataRows.forEach((cells) => {
        if (!cells.some((value) => String(value ?? "").trim() !== "")) return;
        const row = buildNewRow(table);

        if (useHeader) {
            cells.forEach((value, cellIndex) => {
                const targetIndex = headerMapping[cellIndex];
                const targetColumn = targetIndex >= 0 ? table.columns[targetIndex] : null;
                if (targetColumn) {
                    row.cells[targetColumn.id] = toCellString(value);
                }
            });
        } else {
            table.columns.forEach((column, columnIndex) => {
                if (cells[columnIndex] !== undefined) {
                    row.cells[column.id] = toCellString(cells[columnIndex]);
                }
            });
        }

        table.rows.push(row);
        insertedCount += 1;
        lastRowId = row.id;
    });

    return { insertedCount, lastRowId };
}

function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

async function handleJsonFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
        const text = await file.text();
        const parsed = JSON.parse(text);
        replaceProject(parsed, { recordHistory: true, toast: `${file.name} 을(를) 불러왔습니다.` });
    } catch (_error) {
        showToast("JSON 파일을 읽지 못했습니다.", "error");
    } finally {
        event.target.value = "";
    }
}

async function handleCsvFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
        const text = await file.text();
        const workbook = XLSX.read(text, { type: "string" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
        const table = buildTableFromMatrix(stripFileExtension(file.name), matrix, state.project);
        if (!table) throw new Error("empty");
        updateProject((project) => {
            project.tables.push(table);
        }, `${file.name} 을(를) 테이블로 가져왔습니다.`);
        setView("data", table.id);
    } catch (_error) {
        showToast("CSV 파일을 읽지 못했습니다.", "error");
    } finally {
        event.target.value = "";
    }
}

async function handleExcelFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(buffer), { type: "array" });
        const newTables = workbook.SheetNames.map((sheetName) => {
            const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: "" });
            return buildTableFromMatrix(sheetName, matrix, state.project);
        }).filter(Boolean);

        if (!newTables.length) {
            showToast("가져올 시트가 없습니다.", "warning");
            return;
        }

        updateProject((project) => {
            newTables.forEach((table) => project.tables.push(table));
        }, `${newTables.length}개 시트를 가져왔습니다.`);
        setView("data", newTables[0].id);
    } catch (_error) {
        showToast("엑셀 파일을 읽지 못했습니다.", "error");
    } finally {
        event.target.value = "";
    }
}

function buildTableFromMatrix(rawName, matrix, project) {
    const rows = Array.isArray(matrix) ? matrix : [];
    const nonEmptyRows = rows.filter((row) => Array.isArray(row) && row.some((value) => toCellString(value).trim() !== ""));
    if (nonEmptyRows.length === 0) return null;

    const headers = makeUniqueHeaders(nonEmptyRows[0]);
    const dataRows = nonEmptyRows.slice(1);
    const columnTypes = headers.map((header, index) => inferColumnDefinition(header, dataRows.map((row) => toCellString(row[index] ?? ""))));
    const columns = columnTypes.map((definition) => createColumn(definition));

    const firstColumnValues = dataRows.map((row) => toCellString(row[0] ?? "").trim()).filter(Boolean);
    const firstColumnUnique = firstColumnValues.length > 0 && new Set(firstColumnValues).size === firstColumnValues.length;

    let syntheticIdColumn = null;
    if (firstColumnUnique) {
        columns[0].pk = true;
        columns[0].nn = true;
        columns[0].uq = true;
    } else {
        syntheticIdColumn = createColumn({
            name: "id",
            type: "INT",
            pk: true,
            nn: true,
            uq: true,
            description: "자동 생성된 기본 키",
        });
        columns.unshift(syntheticIdColumn);
    }

    const table = {
        id: uid("table"),
        name: nextTableName(project, rawName || "Imported Table"),
        note: "Imported data",
        position: nextTablePosition(project.tables.length + 1),
        columns,
        rows: dataRows.map((row, index) => {
            const cells = {};
            if (syntheticIdColumn) {
                cells[syntheticIdColumn.id] = String(index + 1);
            }
            columnTypes.forEach((definition, definitionIndex) => {
                const column = columns.find((item) => item.name === definition.name);
                if (!column) return;
                cells[column.id] = toCellString(row[definitionIndex] ?? "");
            });
            return {
                id: uid("row"),
                cells,
            };
        }),
        filterPresets: [],
    };

    const normalized = normalizeTable(table);
    repairLegacyReferences({ tables: [...project.tables, normalized] });
    return normalized;
}

function makeUniqueHeaders(sourceHeaders) {
    const used = new Set();
    return sourceHeaders.map((header, index) => {
        const base = slugifyHeader(header) || `column_${index + 1}`;
        let candidate = base;
        let count = 2;
        while (used.has(candidate)) {
            candidate = `${base}_${count}`;
            count += 1;
        }
        used.add(candidate);
        return candidate;
    });
}

function slugifyHeader(value) {
    return String(value ?? "")
        .trim()
        .replace(/\s+/g, "_")
        .replace(/[^\w가-힣]/g, "_")
        .replace(/_+/g, "_")
        .replace(/^_+|_+$/g, "")
        .toLowerCase();
}

function inferColumnDefinition(name, values) {
    const filled = values.map((value) => String(value).trim()).filter(Boolean);
    const uniqueValues = Array.from(new Set(filled));

    if (filled.length && filled.every((value) => /^-?\d+$/.test(value))) {
        return { name, type: "INT" };
    }
    if (filled.length && filled.every((value) => /^-?\d+(\.\d+)?$/.test(value))) {
        return { name, type: "FLOAT" };
    }
    if (filled.length && filled.every((value) => /^(true|false|1|0|yes|no|y|n)$/i.test(value))) {
        return { name, type: "BOOLEAN" };
    }
    if (filled.length && filled.every((value) => /^\d{4}-\d{2}-\d{2}$/.test(value))) {
        return { name, type: "DATE" };
    }
    if (filled.length && uniqueValues.length >= 2 && uniqueValues.length <= 8 && filled.length >= uniqueValues.length * 2) {
        return { name, type: "ENUM", enumValues: uniqueValues };
    }
    if (
        filled.length &&
        filled.every((value) => {
            try {
                JSON.parse(value);
                return true;
            } catch (_error) {
                return false;
            }
        })
    ) {
        return { name, type: "JSON" };
    }
    return { name, type: "STRING" };
}

function stripFileExtension(filename) {
    return String(filename || "Imported Table").replace(/\.[^.]+$/, "");
}

function getWideTableDataStats(table, rows, query, sort) {
    const rowIssueMap = state.validation.cellIssueMap[table.id] || {};
    const sortColumn = sort.columnId ? getColumnById(table, sort.columnId) : null;
    const invalidRows = Object.keys(rowIssueMap).length;
    const invalidCells = Object.values(rowIssueMap).reduce((sum, columnMap) => sum + Object.keys(columnMap || {}).length, 0);

    return {
        visibleRows: rows.length,
        hiddenRows: Math.max(0, table.rows.length - rows.length),
        invalidRows,
        invalidCells,
        queryLabel: query.trim() || "No active search",
        sortLabel: sortColumn && sort.direction ? `${sortColumn.name} ${sort.direction}` : "No active sort",
    };
}

function renderSchemaViewWide() {
    const stats = getProjectStats();
    return `
        <div class="workspace-scroll" data-scroll-root="${getScrollKey()}">
            <div class="workspace-stack">
                <div class="schema-layout">
                    <section class="panel panel--summary">
                        <div class="panel-header">
                            <div class="panel-header__copy">
                                <div class="panel-eyebrow">Project pulse</div>
                                <h3 class="panel-title">Structure at a glance</h3>
                                <p class="panel-subtitle">요약 정보와 빠른 액션은 왼쪽에 고정하고, 오른쪽은 테이블 스키마 편집에 집중하도록 정리했습니다.</p>
                            </div>
                            <div class="toolbar-strip">
                                <span class="tag">Stable cells</span>
                                <span class="tag">Explicit FK</span>
                                <span class="tag">Autosave</span>
                            </div>
                        </div>
                        <div class="panel-body">
                            <div class="summary-grid summary-grid--schema">
                                <div class="summary-card summary-card--compact">
                                    <div class="summary-card__label">Tables</div>
                                    <div class="summary-card__value">${stats.tables}</div>
                                </div>
                                <div class="summary-card summary-card--compact">
                                    <div class="summary-card__label">Columns</div>
                                    <div class="summary-card__value">${stats.columns}</div>
                                </div>
                                <div class="summary-card summary-card--compact">
                                    <div class="summary-card__label">Rows</div>
                                    <div class="summary-card__value">${stats.rows}</div>
                                </div>
                                <div class="summary-card summary-card--compact">
                                    <div class="summary-card__label">Relations</div>
                                    <div class="summary-card__value">${stats.relations}</div>
                                </div>
                            </div>

                            <div class="rail-section">
                                <div class="panel-eyebrow">Quick actions</div>
                                <div class="rail-actions">
                                    <button class="solid-button" data-action="add-table">Add table</button>
                                    <button class="ghost-button" data-action="set-view" data-view="validation">Open validation</button>
                                    <button class="ghost-button" data-action="copy-sql">Copy SQL</button>
                                    <button class="ghost-button" data-action="export-sql">Export SQL</button>
                                </div>
                            </div>

                            <div class="rail-section">
                                <div class="panel-eyebrow">Structure guide</div>
                                <div class="detail-list">
                                    <div class="detail-item">
                                        <span class="detail-item__label">Editing</span>
                                        <span class="detail-item__value">컬럼은 drag 로 순서를 바꾸고 이름을 즉시 수정합니다.</span>
                                    </div>
                                    <div class="detail-item">
                                        <span class="detail-item__label">Relations</span>
                                        <span class="detail-item__value">FK 는 reference, cardinality, relation name 을 같이 유지합니다.</span>
                                    </div>
                                    <div class="detail-item">
                                        <span class="detail-item__label">Checks</span>
                                        <span class="detail-item__value">Validation 화면에서 타입, UQ, FK 문제를 바로 되짚습니다.</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </section>

                    <div class="schema-main">
                        ${
                            state.project.tables.length
                                ? `<div class="schema-table-list">${state.project.tables.map((table) => renderSchemaTablePanel(table)).join("")}</div>`
                                : `<section class="panel panel--fill">${renderEmptyState(
                                      "테이블이 없습니다",
                                      "새 테이블을 추가하고 컬럼, 관계, 샘플 데이터를 한 곳에서 설계해보세요.",
                                      "add-table",
                                      "첫 테이블 만들기",
                                      "",
                                      "compact",
                                  )}</section>`
                        }
                    </div>
                </div>
            </div>
        </div>
    `;
}

function renderDataViewWide() {
    const table = getActiveTable();
    if (!table) {
        return `
            <div class="workspace-scroll" data-scroll-root="${getScrollKey()}">
                ${renderEmptyState("열린 테이블이 없습니다", "왼쪽 사이드바에서 테이블을 선택해 데이터 편집으로 이동하세요.", "set-view", "Schema", "schema")}
            </div>
        `;
    }

    const query = state.ui.dataSearch[table.id] || "";
    const sort = state.ui.dataSort[table.id] || { columnId: null, direction: null };
    const rows = getVisibleRows(table, query, sort);
    const stats = getWideTableDataStats(table, rows, query, sort);
    const presets = table.filterPresets || [];
    const bulkPasteOpen = Boolean(state.ui.bulkPasteOpen[table.id]);
    const bulkPasteDraft = state.ui.bulkPasteDraft[table.id] || "";
    const selectedMap = state.ui.selectedRows[table.id] || {};
    const selectedVisibleCount = rows.filter((row) => selectedMap[row.id]).length;
    const totalSelectedCount = Object.keys(selectedMap).length;
    const bulkEditColumnId = state.ui.bulkEditColumn[table.id] || table.columns[0]?.id || "";
    const bulkEditValue = state.ui.bulkEditValue[table.id] || "";

    return `
        <div class="workspace-scroll" data-scroll-root="${getScrollKey()}">
            <div class="data-layout">
                <section class="panel panel--summary">
                    <div class="panel-header">
                        <div class="panel-header__copy">
                            <div class="panel-eyebrow">Data overview</div>
                            <h3 class="panel-title">${escapeHtml(table.name)}</h3>
                            <p class="panel-subtitle">검색, 정렬, 저장된 뷰, 빠른 액션을 왼쪽에 묶고 오른쪽은 표 편집에 집중하도록 정리했습니다.</p>
                        </div>
                        <div class="toolbar-strip">
                            <span class="tag tag--neutral">${table.columns.length} columns</span>
                            ${state.validation.tableIssueCounts[table.id] ? `<span class="tag tag--warning">${state.validation.tableIssueCounts[table.id]} issues</span>` : `<span class="tag">Valid</span>`}
                        </div>
                    </div>
                    <div class="panel-body">
                        <div class="summary-grid summary-grid--data">
                            <div class="summary-card summary-card--compact">
                                <div class="summary-card__label">Rows</div>
                                <div class="summary-card__value">${table.rows.length}</div>
                            </div>
                            <div class="summary-card summary-card--compact">
                                <div class="summary-card__label">Visible</div>
                                <div class="summary-card__value">${stats.visibleRows}</div>
                            </div>
                            <div class="summary-card summary-card--compact">
                                <div class="summary-card__label">Hidden</div>
                                <div class="summary-card__value">${stats.hiddenRows}</div>
                            </div>
                            <div class="summary-card summary-card--compact">
                                <div class="summary-card__label">Invalid cells</div>
                                <div class="summary-card__value">${stats.invalidCells}</div>
                            </div>
                        </div>

                        <div class="rail-section">
                            <div class="panel-eyebrow">View controls</div>
                            <div class="detail-list">
                                <label class="field-label" for="data-search-${table.id}">Search rows</label>
                                <input
                                    id="data-search-${table.id}"
                                    class="input"
                                    type="search"
                                    data-field="data-search"
                                    data-table-id="${table.id}"
                                    value="${escapeAttr(query)}"
                                    placeholder="Search every column"
                                >
                                <div class="detail-item">
                                    <span class="detail-item__label">Search</span>
                                    <span class="detail-item__value">${escapeHtml(stats.queryLabel)}</span>
                                </div>
                                <div class="detail-item">
                                    <span class="detail-item__label">Sort</span>
                                    <span class="detail-item__value">${escapeHtml(stats.sortLabel)}</span>
                                </div>
                                <div class="detail-item">
                                    <span class="detail-item__label">Invalid rows</span>
                                    <span class="detail-item__value">${stats.invalidRows}</span>
                                </div>
                            </div>
                        </div>

                        <div class="rail-section">
                            <div class="panel-eyebrow">Quick actions</div>
                            <div class="rail-actions">
                                <button class="solid-button" data-action="add-row" data-table-id="${table.id}">Add row</button>
                                <button class="ghost-button" data-action="duplicate-row" data-table-id="${table.id}">Duplicate last row</button>
                                <button class="ghost-button" data-action="save-filter-preset" data-table-id="${table.id}">Save view</button>
                                <button class="ghost-button" data-action="reset-data-view" data-table-id="${table.id}">Reset view</button>
                                <button class="ghost-button" data-action="toggle-bulk-paste" data-table-id="${table.id}">${bulkPasteOpen ? "Hide paste" : "Bulk paste"}</button>
                                <button class="ghost-button" data-action="export-csv" data-table-id="${table.id}">Export CSV</button>
                                <button class="danger-button" data-action="clear-rows" data-table-id="${table.id}">Clear rows</button>
                            </div>
                        </div>

                        ${
                            presets.length
                                ? `
                                    <div class="rail-section">
                                        <div class="panel-eyebrow">Saved views</div>
                                        <div class="preset-bar preset-bar--rail">${presets.map((preset) => renderFilterPresetChip(table, preset)).join("")}</div>
                                    </div>
                                `
                                : ""
                        }
                    </div>
                </section>

                <section class="panel panel--fill">
                    <div class="panel-header">
                        <div class="panel-header__copy">
                            <div class="panel-eyebrow">Data Editor</div>
                            <h3 class="panel-title">${escapeHtml(table.name)}</h3>
                            <p class="panel-subtitle">${escapeHtml(table.note || "샘플 데이터를 입력하고 정렬/검색/검증 상태를 확인합니다.")}</p>
                        </div>
                        <div class="toolbar-strip">
                            <span class="tag tag--neutral">${table.rows.length} rows</span>
                            <span class="tag tag--neutral">${table.columns.length} columns</span>
                            ${state.validation.tableIssueCounts[table.id] ? `<span class="tag tag--warning">${state.validation.tableIssueCounts[table.id]} issues</span>` : `<span class="tag">Valid</span>`}
                            <button class="ghost-button" data-action="set-view" data-view="schema">Open schema</button>
                        </div>
                    </div>
                    <div class="panel-body">
                        <div class="bulk-edit-bar">
                            <span class="tag tag--neutral">${totalSelectedCount} selected</span>
                            <select class="select" style="max-width: 220px;" data-field="bulk-edit-column" data-table-id="${table.id}">
                                ${table.columns.map((column) => `<option value="${column.id}" ${bulkEditColumnId === column.id ? "selected" : ""}>${escapeHtml(column.name)}</option>`).join("")}
                            </select>
                            <input
                                class="input"
                                style="max-width: 240px;"
                                type="text"
                                data-field="bulk-edit-value"
                                data-table-id="${table.id}"
                                value="${escapeAttr(bulkEditValue)}"
                                placeholder="Apply value to selected rows"
                            >
                            <button class="ghost-button" data-action="apply-bulk-edit" data-table-id="${table.id}">Apply to selected</button>
                            <button class="ghost-button" data-action="clear-selected-rows" data-table-id="${table.id}">Clear selection</button>
                            <button class="danger-button" data-action="delete-selected-rows" data-table-id="${table.id}" ${totalSelectedCount ? "" : "disabled"}>Delete selected</button>
                        </div>

                        ${
                            bulkPasteOpen
                                ? `
                                    <div class="paste-panel" style="margin-top: 16px;">
                                        <div class="panel-eyebrow">Bulk paste</div>
                                        <p class="muted-note">Paste TSV or CSV rows. If the first row matches column names, it will be treated as a header.</p>
                                        <textarea
                                            class="textarea"
                                            data-field="bulk-paste-text"
                                            data-table-id="${table.id}"
                                            placeholder="player_id\tnickname\tcreated_at&#10;3\tNewUser\t2026-05-21"
                                        >${escapeHtml(bulkPasteDraft)}</textarea>
                                        <div class="toolbar-strip" style="margin-top: 12px;">
                                            <button class="solid-button" data-action="apply-bulk-paste" data-table-id="${table.id}">Append rows</button>
                                            <button class="ghost-button" data-action="toggle-bulk-paste" data-table-id="${table.id}">Close</button>
                                        </div>
                                    </div>
                                `
                                : ""
                        }

                        <div class="data-table-wrap" style="margin-top: 16px;">
                            <table class="data-table">
                                <thead>
                                    <tr>
                                        <th style="width: 54px;">
                                            <input
                                                class="checkbox"
                                                type="checkbox"
                                                data-field="select-visible-rows"
                                                data-table-id="${table.id}"
                                                ${rows.length > 0 && selectedVisibleCount === rows.length ? "checked" : ""}
                                            >
                                        </th>
                                        <th style="width: 66px;">#</th>
                                        ${table.columns.map((column) => renderDataHeaderCell(table, column, sort)).join("")}
                                        <th style="width: 88px;"></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${
                                        rows.length
                                            ? rows.map((row, index) => renderDataRow(table, row, index)).join("")
                                            : `<tr><td colspan="${table.columns.length + 3}">${renderTableEmptyRow("No rows match the current filters.")}</td></tr>`
                                    }
                                </tbody>
                            </table>
                        </div>
                    </div>
                </section>
            </div>
        </div>
    `;
}

function renderValidationViewWide() {
    const issues = [...state.validation.issues].sort((left, right) => {
        if (left.level === right.level) return 0;
        return left.level === "error" ? -1 : 1;
    });

    return `
        <div class="workspace-scroll" data-scroll-root="${getScrollKey()}">
            <div class="workspace-stack">
                <div class="validation-layout">
                    <section class="panel panel--summary">
                        <div class="panel-header">
                            <div class="panel-header__copy">
                                <div class="panel-eyebrow">Validation</div>
                                <h3 class="panel-title">스키마와 데이터 상태를 한 번에 점검합니다.</h3>
                                <p class="panel-subtitle">컬럼 중복, 타입 불일치, 고유값 충돌, 잘못된 FK 참조를 즉시 확인할 수 있습니다.</p>
                            </div>
                            <div class="toolbar-strip">
                                <span class="tag tag--error">${state.validation.errorCount} errors</span>
                                <span class="tag tag--warning">${state.validation.warningCount} warnings</span>
                            </div>
                        </div>
                        <div class="panel-body">
                            <div class="summary-grid summary-grid--validation">
                                <div class="summary-card summary-card--compact">
                                    <div class="summary-card__label">Errors</div>
                                    <div class="summary-card__value">${state.validation.errorCount}</div>
                                </div>
                                <div class="summary-card summary-card--compact">
                                    <div class="summary-card__label">Warnings</div>
                                    <div class="summary-card__value">${state.validation.warningCount}</div>
                                </div>
                                <div class="summary-card summary-card--compact">
                                    <div class="summary-card__label">Affected Tables</div>
                                    <div class="summary-card__value">${Object.keys(state.validation.tableIssueCounts).length}</div>
                                </div>
                                <div class="summary-card summary-card--compact">
                                    <div class="summary-card__label">Clean Tables</div>
                                    <div class="summary-card__value">${Math.max(0, state.project.tables.length - Object.keys(state.validation.tableIssueCounts).length)}</div>
                                </div>
                            </div>
                        </div>
                    </section>

                    <section class="panel panel--fill">
                        <div class="panel-header">
                            <div class="panel-header__copy">
                                <div class="panel-eyebrow">Issue list</div>
                                <h3 class="panel-title">전체 이슈</h3>
                                <p class="panel-subtitle">오류는 먼저, 경고는 그 다음으로 정렬했습니다.</p>
                            </div>
                        </div>
                        <div class="panel-body">
                            ${
                                issues.length
                                    ? `<div class="validation-list">${issues.map((issue) => renderIssueRow(issue)).join("")}</div>`
                                    : renderEmptyState("이슈가 없습니다", "현재 프로젝트는 검증 기준을 통과했습니다.", null, "", "", "compact")
                            }
                        </div>
                    </section>
                </div>
            </div>
        </div>
    `;
}

function init() {
    loadProject();
    bindEvents();
    renderApp();
}

init();
