/**
 * GridSheet with preSave / postSave per endpoint (save & update)
 * ------------------------------------------------------------------------
 * - Suitable when validation/actions for save & update are different
 * - Explanations/notes on each section
 * 
 * @author sxcharacter@gmail.com
 */

// Error handling using console.error

/**
 * GridSheet - DataTables extension for interactive Excel-like editing
 * 
 * Main features:
 * - Cell editing (text, checkbox, select)
 * - Multi-cell and multi-row selection
 * - Copy/paste with Ctrl+C/V
 * - Fill handle drag (like Excel)
 * - Batch save mode with localStorage
 * - Direct save mode with API calls
 * 
 * @class GridSheet
 * @version 1.0.0
 * @author sxcharacter
 */
class GridSheet {

    /**
     * Constructor for GridSheet
     * 
     * @param {DataTable|Object} tableOrOptions - DataTable instance or options object
     * @param {Object} [maybeOptions=null] - Options if first parameter is table
     * 
     * @param {Array<string>} [options.allowedFields] - List of editable fields
     * @param {boolean} [options.enableColumnNumber=false] - Show automatic No. column
     * @param {boolean} [options.allowAddEmptyRow=false] - Add empty row at the end
     * @param {string} [options.tableSelector] - CSS selector for table
     * @param {Object} [options.emptyTable] - Empty table mode configuration
     * @param {boolean} [options.emptyTable.enabled=false] - Enable empty table mode
     * @param {string} [options.emptyTable.saveMode='direct'] - 'direct' or 'batch'
     * @param {string} [options.emptyTable.storageKey] - Key for localStorage
     * @param {Object} [options.endpoints] - Endpoint configurations
     * @param {Object} [options.endpoints.save] - Save endpoint config
     * @param {Object} [options.endpoints.update] - Update endpoint config
     * @param {Object} [options.endpoints.delete] - Delete endpoint config
     */
    constructor(tableOrOptions = {}, maybeOptions = null) {
        // ============================================
        // DEPENDENCY CHECKS
        // ============================================
        if (typeof jQuery === 'undefined') {
            throw new Error('GridSheet requires jQuery. Please include jQuery before this script.');
        }
        if (!jQuery.fn.dataTable) {
            throw new Error('GridSheet requires DataTables. Please include datatables.min.js before this script.');
        }
        if (!jQuery.fn.dataTable.KeyTable) {
            throw new Error('GridSheet requires KeyTable extension. Please include dataTables.keyTable.min.js before this script.');
        }

        // Support two calling methods:
        // 1. new GridSheet(table, options) - table as first parameter
        // 2. new GridSheet(options) - options.table contains table instance
        let options;
        if (maybeOptions !== null) {
            // Method 1: (table, options)
            options = maybeOptions;
            options.table = tableOrOptions;
        } else if (tableOrOptions && typeof tableOrOptions.row === 'function') {
            // Method 1 without options: only (table)
            options = { table: tableOrOptions };
        } else {
            // Method 2: (options)
            options = tableOrOptions;
        }
        this.allowedFields = options.allowedFields || false;
        this.selectOptions = options.selectOptions || {};
        this.allowAddEmptyRow = options.allowAddEmptyRow || false;
        this.allowDeleteRow = options.allowDeleteRow !== undefined ? options.allowDeleteRow : true;
        this.enableColumnNumber = options.enableColumnNumber || false;
        this.extraHeaders = options.extraHeaders || {};
        this.dataTableOptions = options.dataTableOptions || { keys: true, paging: true, order: [] };

        // Parameters for empty table (nested object for extensibility)
        // emptyTable: { enabled: true, initialRows: 5, saveMode: 'direct'|'batch', storageKey: 'myTable' }
        this.emptyTable = options.emptyTable || { enabled: false, initialRows: 1, saveMode: 'direct' };
        // Normalize: if only boolean is given, convert to object
        if (typeof this.emptyTable === 'boolean') {
            this.emptyTable = { enabled: this.emptyTable, initialRows: 1, saveMode: 'direct' };
        }
        // Set defaults
        this.emptyTable.enabled = this.emptyTable.enabled || false;
        this.emptyTable.initialRows = this.emptyTable.initialRows || 1;
        this.emptyTable.saveMode = this.emptyTable.saveMode || 'direct'; // 'direct' = save immediately, 'batch' = save to localStorage first
        this.emptyTable.storageKey = this.emptyTable.storageKey || 'GridSheet_data';
        this.emptyTable.saveButtonText = this.emptyTable.saveButtonText || 'Save';
        this.emptyTable.showClearButton = this.emptyTable.showClearButton !== undefined ? this.emptyTable.showClearButton : false;
        this.emptyTable.showSaveButton = this.emptyTable.showSaveButton !== undefined ? this.emptyTable.showSaveButton : false;
        this.emptyTable.autoRestore = this.emptyTable.autoRestore !== undefined ? this.emptyTable.autoRestore : false;

        // If emptyTable enabled, force paging, info and ordering to false
        // ordering: false is important so inserted rows don't change position on draw()
        if (this.emptyTable.enabled) {
            this.dataTableOptions.paging = false;
            this.dataTableOptions.info = false;
            this.dataTableOptions.ordering = false;
        }

        /**
         * Footer Configuration
         * For showing aggregate row (sum, avg, count) at the bottom of table
         * footer: { enabled: true, label: 'Total' }
         * Column configuration via data-footer="sum|avg|count" on th elements
         */
        this.footer = options.footer || { enabled: false };
        if (typeof this.footer === 'boolean') {
            this.footer = { enabled: this.footer };
        }
        this.footer.enabled = this.footer.enabled || false;
        this.footer.label = this.footer.label !== undefined ? this.footer.label : 'Total';

        /**
         * Column Configuration
         * Centralized column config in JavaScript instead of HTML data-* attributes
         * columns: { 0: { name: 'field', type: 'text', ... }, 1: { ... } }
         * Index 0 = first data column (No. column excluded if enableColumnNumber is true)
         */
        this.columns = options.columns || null;

        /**
         * Format Configuration
         * For formatting number, currency, percentage, etc.
         * If no format specified, value will be displayed as-is.
         */
        const defaultFormats = {
            number: {
                thousandSeparator: '',
                decimalSeparator: '.',
                decimalPlaces: null,  // null = preserve original
                prefix: '',
                suffix: ''
            },
            currency: {
                thousandSeparator: '.',
                decimalSeparator: ',',
                decimalPlaces: 0,
                prefix: '',
                suffix: ''
            },
            percentage: {
                thousandSeparator: '',
                decimalSeparator: ',',
                decimalPlaces: 0,
                prefix: '',
                suffix: '%'
            },
            date: {
                displayFormat: 'dd/mm/yyyy',  // Format for display
                storageFormat: 'yyyy-mm-dd'   // ISO format for storage
            },
            email: {
                validateOnBlur: true          // Validate email format on blur
            },
            textarea: {
                displayLength: 25,            // Max chars shown in cell before "..."
                maxLength: null,              // Max input length (null = unlimited)
                rows: 5,                      // Textarea rows in edit modal
                editMode: 'modal'             // 'modal' = modal overlay, 'inline' = inline popup
            }
        };
        // Merge user formats with defaults
        this.formats = {};
        if (options.formats) {
            Object.keys(defaultFormats).forEach(type => {
                this.formats[type] = { ...defaultFormats[type], ...(options.formats[type] || {}) };
            });
            // Also add any custom types from user
            Object.keys(options.formats).forEach(type => {
                if (!this.formats[type]) {
                    this.formats[type] = options.formats[type];
                }
            });
        } else {
            this.formats = defaultFormats;
        }

        /**
         * Centralized State Management
         * All state variables are managed in a single object for easier debugging and maintenance
         */
        this._state = {
            // Batch Mode Data
            pendingData: [],       // Array: pending data (new rows) for localStorage
            editedData: {},        // Object: edited data {rowId: {id, fields, timestamp}}
            deletedData: [],       // Array: deleted row IDs [rowId1, rowId2, ...]

            // Edit Mode
            isEditMode: false,     // Boolean: currently in cell edit mode?
            currentCell: null,     // DataTable Cell: currently focused cell
            focusedRow: null,      // Number: focused row index
            lastFocusedRow: null,  // Number: last focused row
            currentFocusedRow: null,
            previousFocusedRow: null,

            // Cell Selection
            selectionStart: null,  // {row, col}: selection start point
            selectionEnd: null,    // {row, col}: selection end point
            selectedCells: [],     // Array<{row, col}>: selected cells

            // Row Selection
            isRowSelecting: false, // Boolean: currently drag selecting row?
            rowSelectionStart: null, // Number: starting row index
            selectedRows: [],      // Array<Number>: selected row indices

            // Fill Handle
            fillHandleElement: null, // HTMLElement: fill handle element
            isFillDragging: false, // Boolean: currently drag filling?
            fillStartCells: [],    // Array: initial fill cells
            fillOriginalBounds: null, // Object: original bounds
            fillSourceData: [],    // Array: source data for fill
            fillTargetEnd: null,   // {row, col}: fill target end

            // Clipboard
            hasClipboardData: false, // Boolean: has data in clipboard?
            clipboardRowCount: 0,    // Number: how many rows in clipboard

            // Keyboard
            isShiftPressed: false, // Boolean: Shift key pressed?
            isDragging: false,     // Boolean: currently dragging?

            // UI State
            emptyRowExists: false, // Boolean: empty row exists?
            activeCellNode: null,  // HTMLElement: active cell node
            savingCell: false,     // Boolean: currently saving cell?
            isAddingEmptyRow: false, // Boolean: currently adding empty row?
            isInsertingRow: false, // Boolean: currently inserting row?
            isCheckboxChanging: false, // Boolean: currently changing checkbox?
            highlightedRow: null   // HTMLElement: highlighted row
        };

        // Internal cache for raw values to persist across DataTables redraws
        // Key format: "rowIndex_colIndex" -> rawValue
        // This prevents corruption when _applyColumnTypes is called after redraw
        this._rawValueCache = new Map();

        /**
         * Language/i18n Support
         * Default labels (English). User can override via options.language
         */
        const defaultLang = {
            // Context Menu
            copyRow: 'Copy Row',
            pasteRow: 'Paste Row',
            pasteRows: 'Paste {n} Rows',
            deleteRow: 'Delete Row',
            clearRow: 'Clear Row',
            copyRows: 'Copy {n} Rows',
            clearRows: 'Clear {n} Rows',
            deleteRows: 'Delete {n} Rows',
            insertAbove: 'Insert Row Above',
            insertBelow: 'Insert Row Below',
            copyCells: 'Copy Cells',
            pasteCells: 'Paste Cells',
            copy: 'Copy',
            paste: 'Paste',  // For single cell paste
            readonly: 'Readonly',
            readonlyRow: 'Readonly Row',
            clear: 'Clear',
            reset: 'Reset',

            // Dropdown
            noResult: 'No results found',

            // Buttons
            saveAll: 'Save',
            cancel: 'Cancel',
            discard: 'Discard',
            restore: 'Restore',

            // Restore Modal
            restoreTitle: 'Restore unsaved changes?',
            restoreMessage: 'You have pending changes:',
            restoreNewRow: 'new row',

            // Clear/Reset Modal
            clearTitle: 'Clear all changes?',
            clearMessage: 'This will discard all pending changes and reload the page to restore original data.'
        };
        this.lang = { ...defaultLang, ...options.language };

        // Backward compatibility: alias for direct access (will be deprecated)
        // Use _getState() and _setState() for new code
        this._pendingData = this._state.pendingData;
        this._editedData = this._state.editedData;
        this._deletedData = this._state.deletedData;

        // endpoints: {save: {...}, update: {...}, delete: {...}, batch: {...}}
        this.endpoints = options.endpoints || {
            save: { endpoint: "", preSave: null, postSave: null },
            update: { endpoint: "", preSave: null, postSave: null },
            delete: { endpoint: "", preSave: null, postSave: null },
            batch: { endpoint: "", preSave: null, postSave: null }
        };

        this._selectOptionCache = {};

        // Support 3 modes:
        // 1. tableSelector + emptyTable.enabled -> create empty table with empty rows
        // 2. tableSelector (string) -> inject No column + init DataTable internally
        // 3. table (DataTable instance) -> use directly, assume No column already exists
        if (options.tableSelector) {
            this.tableSelector = options.tableSelector;

            // If emptyTable.enabled, prepare tbody with empty rows
            if (this.emptyTable.enabled) {
                this._prepareEmptyTable(this.tableSelector);
            }

            // Inject No. column BEFORE init DataTable if enableColumnNumber is active
            if (this.enableColumnNumber) {
                this._injectColumnNumber(this.tableSelector);
            }

            // Apply columns config from JavaScript to HTML th elements
            // This must run BEFORE dateColumnDefs detection
            if (this.columns) {
                this._applyColumnConfig(this.tableSelector);
            }

            // Auto-generate columnDefs to disable auto-type detection for date columns
            // This prevents DataTables from automatically formatting date columns
            const tableNode = document.querySelector(this.tableSelector);
            if (tableNode) {
                const headers = tableNode.querySelectorAll('thead th');
                const dateColumnDefs = [];
                headers.forEach((th, idx) => {
                    const dataType = th.getAttribute('data-type');
                    if (dataType === 'date') {
                        // Just set type to string to disable auto date formatting
                        dateColumnDefs.push({ targets: idx, type: 'string' });
                    }
                });

                // Merge with existing columnDefs if any
                if (dateColumnDefs.length > 0) {
                    this.dataTableOptions = this.dataTableOptions || {};
                    this.dataTableOptions.columnDefs = this.dataTableOptions.columnDefs || [];
                    this.dataTableOptions.columnDefs = [...this.dataTableOptions.columnDefs, ...dateColumnDefs];

                    // Disable auto type detection for date columns - requires DataTables 2.1.6+
                    // This prevents DataTables from auto-formatting dates
                    this.dataTableOptions.typeDetect = false;
                }
            }

            // Init DataTable internally (use jQuery wrapper for better compatibility with older DataTables versions like v1.10)
            this.table = $(this.tableSelector).DataTable(this.dataTableOptions);
        } else if (options.table) {
            this.table = options.table;
        } else {
            throw new Error('GridSheet: tableSelector or table must be provided!');
        }

        this.initialize();

        // Apply data-type from header to cells for styling (e.g., number = right align)
        this._applyColumnTypes();

        // Auto-parse raw values for cells without data-raw-value attribute
        // This must run BEFORE formula calculation so formulas have correct values
        this._autoParseRawValues();

        // Initial formula calculation for any formula columns
        this._recalculateAllFormulas();

        // Initialize footer totals if enabled
        this._initFooter();

        // If batch mode, add Save button and load data from localStorage
        if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
            this._loadFromLocalStorage();
            this._addSaveButton();
        }
    }

    // ==================== STATE MANAGEMENT HELPERS ====================

    /**
     * Get value from centralized state
     * @param {string} key - Key in _state object
     * @returns {*} Value from state
     * 
     * @example
     * const cells = this._getState('selectedCells');
     * const isEditing = this._getState('isEditMode');
     */
    _getState(key) {
        return this._state[key];
    }

    /**
     * Set value in centralized state
     * @param {string} key - Key in _state object
     * @param {*} value - New value
     * 
     * @example
     * this._setState('isEditMode', true);
     * this._setState('selectedCells', [{row: 0, col: 1}]);
     */
    _setState(key, value) {
        this._state[key] = value;
    }

    /**
     * Get all state for debugging
     * @returns {Object} Copy of all state
     */
    _getAllState() {
        return { ...this._state };
    }

    /**
     * Reset state to default values
     * Useful for cleanup or reset selection
     * @param {Array<string>} [keys] - Optional: array of keys to reset. If empty, reset all.
     */
    _resetState(keys = null) {
        const defaults = {
            isEditMode: false,
            currentCell: null,
            selectionStart: null,
            selectionEnd: null,
            selectedCells: [],
            isRowSelecting: false,
            rowSelectionStart: null,
            selectedRows: [],
            isFillDragging: false,
            fillStartCells: [],
            fillOriginalBounds: null,
            fillSourceData: [],
            fillTargetEnd: null,
            isShiftPressed: false,
            isDragging: false
        };

        if (keys && keys.length > 0) {
            keys.forEach(key => {
                if (defaults.hasOwnProperty(key)) {
                    this._state[key] = defaults[key];
                }
            });
        } else {
            Object.keys(defaults).forEach(key => {
                this._state[key] = defaults[key];
            });
        }
    }

    // ==================== END STATE MANAGEMENT ====================

    // ==================== FORMAT HELPERS ====================

    /**
     * Format number for display according to format configuration
     * @param {string|number} value - Original value
     * @param {string} type - Format type (number, currency, percentage)
     * @returns {string} Formatted value
     */
    _formatValue(value, type) {
        if (value === null || value === undefined || value === '') return '';

        const format = this.formats[type];
        if (!format) return String(value);

        // Parse to number (remove existing formatting)
        let num = this._parseValue(String(value), type);
        if (isNaN(num)) return String(value);

        // Apply decimal places
        if (format.decimalPlaces !== null && format.decimalPlaces !== undefined) {
            num = num.toFixed(format.decimalPlaces);
        } else {
            num = String(num);
        }

        // Split integer and decimal parts
        const parts = num.split('.');
        let integerPart = parts[0];
        const decimalPart = parts[1] || '';

        // Apply thousand separator
        if (format.thousandSeparator) {
            integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, format.thousandSeparator);
        }

        // Combine with decimal separator
        let result = integerPart;
        if (decimalPart) {
            result += format.decimalSeparator + decimalPart;
        }

        // Add prefix and suffix
        return (format.prefix || '') + result + (format.suffix || '');
    }

    /**
     * Parse formatted value to original number
     * @param {string} value - Formatted value
     * @param {string} type - Format type (number, currency, percentage)
     * @returns {number} Numeric value
     */
    _parseValue(value, type) {
        if (value === null || value === undefined || value === '') return NaN;

        const format = this.formats[type];
        if (!format) return parseFloat(value);

        let cleaned = String(value);

        // Remove prefix and suffix
        if (format.prefix) {
            cleaned = cleaned.replace(new RegExp('^' + this._escapeRegex(format.prefix)), '');
        }
        if (format.suffix) {
            cleaned = cleaned.replace(new RegExp(this._escapeRegex(format.suffix) + '$'), '');
        }

        // Remove thousand separator
        if (format.thousandSeparator) {
            cleaned = cleaned.split(format.thousandSeparator).join('');
        }

        // Convert decimal separator to standard dot
        if (format.decimalSeparator && format.decimalSeparator !== '.') {
            cleaned = cleaned.replace(format.decimalSeparator, '.');
        }

        return parseFloat(cleaned.trim());
    }

    /**
     * Auto-parse raw values for cells that don't have data-raw-value attribute
     * This runs during initialization to extract numeric values from formatted display text
     * Uses format configuration to reverse-parse values like "Rp. 15.000.000" → 15000000
     */
    _autoParseRawValues() {
        const tableNode = document.querySelector(this.tableSelector);
        if (!tableNode) return;

        const headers = tableNode.querySelectorAll('thead th');
        const tbody = tableNode.querySelector('tbody');
        if (!tbody) return;

        const rows = tbody.querySelectorAll('tr');

        headers.forEach((th, colIdx) => {
            const dataType = th.getAttribute('data-type');
            // Only process columns that need raw values (number, currency, percentage)
            if (!['number', 'currency', 'percentage'].includes(dataType)) return;

            rows.forEach(row => {
                const cell = row.cells[colIdx];
                if (!cell) return;

                // Skip if data-raw-value already exists
                if (cell.hasAttribute('data-raw-value')) return;

                // Get display text and try to parse
                const displayText = cell.textContent.trim();
                if (!displayText) return;

                // Use existing _parseValue to extract raw number
                const rawValue = this._parseValue(displayText, dataType);

                // Only set if we got a valid number
                if (!isNaN(rawValue)) {
                    cell.setAttribute('data-raw-value', rawValue);
                }
            });
        });
    }

    /**
     * Apply column configuration from JavaScript to HTML th elements
     * Sets data-* attributes on th based on columns config
     * Also auto-generates allowedFields and selectOptions
     * @param {string} tableSelector - CSS selector for table
     */
    _applyColumnConfig(tableSelector) {
        const tableNode = document.querySelector(tableSelector);
        if (!tableNode || !this.columns) return;

        const headers = tableNode.querySelectorAll('thead th');
        const autoAllowedFields = [];
        const autoSelectOptions = {};

        // Calculate offset: if enableColumnNumber, headers[0] is No. column
        const offset = this.enableColumnNumber ? 1 : 0;

        // Process each column config
        Object.keys(this.columns).forEach(indexStr => {
            const colIndex = parseInt(indexStr, 10);
            const config = this.columns[colIndex];

            // Get actual th element (add offset for No. column)
            const th = headers[colIndex + offset];
            if (!th || !config) return;

            // Set data-name (required)
            if (config.name) {
                th.setAttribute('data-name', config.name);
            }

            // Set data-type
            if (config.type) {
                th.setAttribute('data-type', config.type);
            }

            // Set data-formula (for formula type)
            if (config.formula) {
                th.setAttribute('data-formula', config.formula);
            }

            // Set data-format (for display format override)
            if (config.format && typeof config.format === 'object') {
                // Store complex format in data attribute as JSON
                th.setAttribute('data-format-config', JSON.stringify(config.format));
                // Also merge into this.formats for the column type
                if (config.type && this.formats[config.type]) {
                    this.formats[config.type] = { ...this.formats[config.type], ...config.format };
                }
            } else if (config.format && typeof config.format === 'string') {
                th.setAttribute('data-format', config.format);
            }

            // Set data-footer (for aggregate)
            if (config.footer) {
                th.setAttribute('data-footer', config.footer);
            }

            // Set data-readonly
            if (config.readonly) {
                th.setAttribute('data-readonly', 'true');
            }

            // Set data-server-field
            if (config.serverField) {
                th.setAttribute('data-server-field', config.serverField);
            }

            // Set data-empty (v1.1.1)
            if (config.allowEmpty !== undefined) {
                th.setAttribute('data-empty', config.allowEmpty ? 'true' : 'false');
            }

            // Handle select options
            if (config.options && config.name) {
                if (typeof config.options === 'string') {
                    // Endpoint URL - add to autoSelectOptions
                    autoSelectOptions[config.name] = config.options;
                } else if (Array.isArray(config.options)) {
                    // Static array - store in data attribute
                    th.setAttribute('data-options', JSON.stringify(config.options));

                    // Convert array to {id: text} object and add to autoSelectOptions
                    const mappedOptions = {};
                    config.options.forEach(opt => {
                        if (typeof opt === 'object' && opt !== null) {
                            mappedOptions[opt.id !== undefined ? opt.id : opt.value] = opt.text;
                        } else {
                            mappedOptions[String(opt)] = String(opt);
                        }
                    });
                    autoSelectOptions[config.name] = mappedOptions;
                }
            }

            // Auto-generate allowedFields (exclude readonly columns)
            if (config.name && !config.readonly && config.type !== 'formula') {
                autoAllowedFields.push(config.name);
            }
        });

        // Auto-set allowedFields if not manually specified
        if (!this.allowedFields && autoAllowedFields.length > 0) {
            this.allowedFields = autoAllowedFields;
        }

        // Auto-set selectOptions if not manually specified
        if (Object.keys(autoSelectOptions).length > 0) {
            this.selectOptions = { ...autoSelectOptions, ...this.selectOptions };
        }
    }

    /**
     * Escape special regex characters
     */
    _escapeRegex(str) {
        return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    /**
     * Parse various date formats to ISO (yyyy-mm-dd)
     * Supports: dd/mm/yyyy, mm/dd/yyyy, yyyy-mm-dd, dd-mm-yyyy
     * @param {string} dateStr - Date string to parse
     * @returns {string} ISO date string (yyyy-mm-dd) or original if invalid
     */
    _parseDateToISO(dateStr) {
        if (!dateStr) return '';

        // Already in ISO format
        if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
            return dateStr;
        }

        // Parse dd/mm/yyyy or dd-mm-yyyy
        let match = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
        if (match) {
            const day = match[1].padStart(2, '0');
            const month = match[2].padStart(2, '0');
            const year = match[3];
            return `${year}-${month}-${day}`;
        }

        // Try to parse with Date object
        const date = new Date(dateStr);
        if (!isNaN(date.getTime())) {
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }

        return dateStr; // Return original if can't parse
    }

    /**
     * Format ISO date for display based on config
     * @param {string} isoDate - ISO date string (yyyy-mm-dd)
     * @returns {string} Formatted date string
     */
    _formatDateForDisplay(isoDate) {
        if (!isoDate) return '';

        // Support both 'format' and 'displayFormat' for backwards compatibility
        const format = (this.formats.date && (this.formats.date.displayFormat || this.formats.date.format)) || 'dd/mm/yyyy';

        // Parse ISO format - handle both with and without leading zeros
        const match = isoDate.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (!match) return isoDate;

        const year = match[1];
        const month = match[2].padStart(2, '0');
        const day = match[3].padStart(2, '0');

        // Apply display format - case insensitive replacement
        let result = format.toLowerCase();
        result = result.replace('yyyy', '{{YEAR}}');
        result = result.replace('mm', '{{MONTH}}');
        result = result.replace('dd', '{{DAY}}');
        result = result.replace('{{YEAR}}', year);
        result = result.replace('{{MONTH}}', month);
        result = result.replace('{{DAY}}', day);

        return result;
    }

    /**
     * Validate email format
     * @param {string} email - Email to validate
     * @returns {boolean} True if valid email
     */
    _isValidEmail(email) {
        if (!email) return true; // Empty is valid (not required check)
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return emailRegex.test(email);
    }

    /**
     * Validate time format (HH:MM or HH:MM:SS)
     * @param {string} time - Time to validate
     * @returns {boolean} True if valid time
     */
    _isValidTime(time) {
        if (!time) return true; // Empty is valid
        return /^([01]?\d|2[0-3]):[0-5]\d(:[0-5]\d)?$/.test(time);
    }

    /**
     * Validate date format (ISO: YYYY-MM-DD)
     * @param {string} date - Date to validate  
     * @returns {boolean} True if valid ISO date
     */
    _isValidDate(date) {
        if (!date) return true; // Empty is valid
        const iso = this._parseDateToISO(date);
        return /^\d{4}-\d{2}-\d{2}$/.test(iso);
    }

    /**
     * Validate number format (for number, currency, percentage)
     * @param {string} value - Value to validate
     * @param {string} type - Data type (number, currency, percentage)
     * @returns {boolean} True if valid number
     */
    _isValidNumber(value, type) {
        if (!value || !String(value).trim()) return true; // Empty is valid
        const cleaned = this._cleanValueForType(value, type);
        return cleaned !== '' && !isNaN(Number(cleaned));
    }

    /**
     * CENTRALIZED VALIDATION - validates field by type
     * Use this wrapper instead of individual validators
     * @param {string} value - Value to validate
     * @param {string} type - Data type from column header
     * @returns {boolean} True if valid
     */
    _validateField(value, type) {
        if (!value || !String(value).trim()) return true; // Empty is always valid

        switch (type) {
            case 'email':
                return this._isValidEmail(value);
            case 'time':
                return this._isValidTime(value);
            case 'date':
                return this._isValidDate(value);
            case 'number':
            case 'currency':
            case 'percentage':
                return this._isValidNumber(value, type);
            default:
                return true; // Unknown types are valid
        }
    }

    /**
     * CENTRALIZED pre-save validation
     * Validates value and calls preSave callback
     * @param {HTMLElement} cellNode - TD element
     * @param {HTMLElement} editorEl - Editor element (input/select/textarea)
     * @param {string} newValue - New value to validate
     * @param {string} rowId - Row ID ('new' or actual ID)
     * @param {string} fieldName - Field name being edited
     * @returns {boolean} True if valid and preSave allows save
     */
    _validateBeforeSave(cellNode, editorEl, newValue, rowId, fieldName) {
        const th = this.table.column(cellNode.cellIndex).header();
        const fieldType = th ? th.getAttribute('data-type') : null;
        const allowEmpty = th && th.getAttribute('data-empty') === 'true';

        let isValidFormat = true;

        // 0. Empty Value Validation - check if column allows empty
        if (newValue === '' || newValue === null || newValue === undefined) {
            if (!allowEmpty) {
                console.log('Empty value validation failed for', fieldName, '- column does not allow empty (no data-empty="true")');
                isValidFormat = false;
            }
        }

        // 1. Browser Native Validation
        if (isValidFormat && editorEl.checkValidity && !editorEl.checkValidity()) {
            isValidFormat = false;
        }

        // 2. Custom Type Validation using centralized wrapper
        if (isValidFormat) {
            isValidFormat = this._validateField(newValue, fieldType);
        }

        // 3. Call preSave callback with consistent payload structure
        const endpoint = rowId === 'new' ? this.endpoints?.save : this.endpoints?.update;
        if (endpoint && typeof endpoint.preSave === 'function') {
            const rowNode = cellNode.closest('tr');
            const rowIndex = rowNode ? this.table.row(rowNode).index() : null;
            const colIndex = cellNode.cellIndex;

            // Build validation errors array
            const validationErrors = [];
            if (!isValidFormat) {
                validationErrors.push({
                    rowId: rowId,
                    rowIndex: rowIndex,
                    field: fieldName,
                    colIndex: colIndex,
                    value: newValue,
                    type: fieldType,
                    errorType: (newValue === '' || newValue === null) ? 'empty' : 'format'
                });
            }

            // Consistent payload structure
            const payload = rowId === 'new'
                ? {
                    operation: 'insert',
                    data: {
                        row: { [fieldName]: newValue }
                    },
                    meta: {
                        timestamp: Date.now(),
                        hasErrors: !isValidFormat,
                        validationErrors: validationErrors,
                        rowCount: 1,
                        tempId: rowNode?.getAttribute('data-id') || 'temp_' + Date.now(),
                        fieldCount: 1,
                        rowIndex: rowIndex
                    }
                }
                : {
                    operation: 'update',
                    data: {
                        id: rowId,
                        fields: { [fieldName]: newValue }
                    },
                    meta: {
                        timestamp: Date.now(),
                        hasErrors: !isValidFormat,
                        validationErrors: validationErrors,
                        rowCount: 1,
                        fieldCount: 1,
                        rowIndex: rowIndex,
                        colIndex: colIndex
                    }
                };
            return endpoint.preSave(payload) !== false;
        }

        return isValidFormat;
    }

    /**
     * Handle validation error - show error state on cell
     * @param {HTMLElement} cellNode - TD element
     * @param {string} newValue - The invalid value to display
     */
    _handleValidationError(cellNode, newValue) {
        cellNode.innerHTML = '';
        cellNode.textContent = newValue;
        cellNode.setAttribute('data-raw-value', newValue);
        cellNode.classList.add('dt-error');
        cellNode.classList.remove('dt-cell-editing');
        console.log('Validation failed (Non-blocking):', newValue);
    }

    // ============================================================
    // SERVER-FIELD AUTO UPDATE METHODS
    // ============================================================

    /**
     * Get all columns that have data-server-field attribute
     * These columns will be auto-updated from server response
     * @returns {Array} Array of {index, serverField, format, name}
     */
    _getServerFieldColumns() {
        const headers = this.table.table().node().querySelectorAll('thead th[data-server-field]');
        return Array.from(headers).map(th => ({
            index: Array.from(th.parentNode.children).indexOf(th),
            serverField: th.getAttribute('data-server-field'),
            format: th.getAttribute('data-format'),
            name: th.getAttribute('data-name')
        }));
    }

    /**
     * Update row cells from server response data
     * Automatically maps response fields to columns with data-server-field attribute
     * @param {HTMLElement} rowNode - TR element to update
     * @param {Object} responseData - Server response data object
     */
    _updateRowFromServerResponse(rowNode, responseData) {
        if (!rowNode || !responseData) return;

        const serverFieldCols = this._getServerFieldColumns();
        if (serverFieldCols.length === 0) return;


        serverFieldCols.forEach(col => {
            const value = responseData[col.serverField];
            if (value !== undefined && value !== null) {
                const cell = rowNode.cells[col.index];
                if (cell) {
                    // Apply formatting if specified
                    let displayValue = value;
                    if (col.format) {
                        displayValue = this._formatServerFieldValue(value, col.format);
                    }

                    cell.textContent = displayValue;
                    cell.setAttribute('data-raw-value', value);
                }
            }
        });


        // Update row ID if returned (common for new rows)
        if (responseData.id) {
            rowNode.setAttribute('data-id', responseData.id);
            rowNode.removeAttribute('data-temp-id');
            rowNode.removeAttribute('data-pending');
            rowNode.removeAttribute('data-new-row');
            rowNode.classList.remove('dt-row-pending');
        }
    }

    /**
     * Format value for display in server-field columns
     * Uses existing format methods based on data-format attribute
     * @param {*} value - Raw value from server
     * @param {string} format - Format type (currency, date, percentage, number)
     * @returns {string} Formatted value
     */
    _formatServerFieldValue(value, format) {
        if (value === null || value === undefined) return '';

        switch (format) {
            case 'currency':
            case 'percentage':
            case 'number':
                // Use existing _formatValue for numeric formats
                return this._formatValue(value, format);
            case 'date':
                // Use existing date formatter
                return this._formatDateForDisplay(value);
            default:
                return String(value);
        }
    }


    /**
     * Get value from editor element (input/select/checkbox/dropdown)
     * @param {HTMLElement} editorEl - Editor element
     * @returns {string} Extracted value
     */
    _getEditorValue(editorEl) {
        const elType = editorEl.tagName.toLowerCase();

        if (elType === 'input') {
            if (editorEl.type === 'checkbox') {
                return editorEl.checked ? 'true' : 'false';
            } else if (editorEl.hasAttribute('data-selected-text')) {
                // Searchable dropdown
                return editorEl.getAttribute('data-selected-text');
            } else {
                return editorEl.value;
            }
        } else if (elType === 'select') {
            return editorEl.options[editorEl.selectedIndex]?.text || '';
        } else if (elType === 'textarea') {
            return editorEl.value;
        }

        return editorEl.value || '';
    }

    /**
     * Handle save for new row (routes to batch or direct mode)
     * @param {HTMLElement} cellNode - TD element
     * @param {string} fieldName - Field name
     * @param {string} newValue - New value
     * @returns {boolean} True if save was triggered
     */
    _handleNewRowSave(cellNode, fieldName, newValue) {
        const rowNode = cellNode.parentNode;

        if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
            return this._saveNewRowBatch(rowNode, fieldName, newValue);
        } else {
            const isFilled = this.isRowRequiredFieldsFilled(rowNode);

            if (rowNode && isFilled) {
                // Direct mode - save to server
                Promise.resolve(this.saveNewRow(rowNode))
                    .finally(() => {
                        if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                            this.addEmptyRow();
                        }
                    });
                return true;
            }
        }
        return false;
    }

    /**
     * Save new row to localStorage in batch mode
     * @param {HTMLElement} rowNode - TR element
     * @param {string} fieldName - Field that was edited
     * @param {string} newValue - New value
     * @returns {boolean} True if row was complete and saved
     */
    _saveNewRowBatch(rowNode, fieldName, newValue) {
        // Generate temp_id for new row if not already set
        let tempId = rowNode.getAttribute('data-id');
        if (tempId === 'new') {
            tempId = 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
            rowNode.setAttribute('data-id', tempId);
            rowNode.setAttribute('data-temp-id', tempId); // Also set data-temp-id for error cell lookup
            rowNode.removeAttribute('data-new-row');
        }

        // Mark as pending (visual indication)
        rowNode.setAttribute('data-pending', 'true');

        // ===== IMPORTANT: Save current field to memory FIRST =====
        // This ensures each field is tracked even before row is complete
        this._updateLocalStorageEntry(tempId, fieldName, newValue, false);

        // Check if row has all required fields filled
        const isComplete = this.isRowRequiredFieldsFilled(rowNode);

        if (!isComplete) {
            return false;
        }

        // Row is complete - merge memory data with DOM data and save to localStorage
        // First collect all fields from DOM
        const rowDataFromDom = this._collectRowDataForBatch(rowNode, tempId);

        // Merge with existing memory data (memory data takes priority for freshness)
        const existingMemoryData = this._pendingData.find(e => e._rowTempId === tempId) || {};
        const mergedData = { ...rowDataFromDom, ...existingMemoryData };
        mergedData._timestamp = Date.now();

        this._saveToPendingData(tempId, mergedData);
        this._updateBatchCount();

        // Add empty row if needed
        if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
            this.addEmptyRow();
        } else {
        }

        return true;
    }

    /**
     * Collect all field values from a row for batch save
     * @param {HTMLElement} rowNode - TR element
     * @param {string} tempId - Temporary row ID
     * @returns {Object} Row data object
     */
    _collectRowDataForBatch(rowNode, tempId) {
        const rowData = { _rowTempId: tempId, _timestamp: Date.now() };
        const headers = this.table.columns().header().toArray();

        Array.from(rowNode.children).forEach((cell, idx) => {
            if (this.enableColumnNumber && idx === 0) return; // Skip No. column
            const th = headers[idx];
            if (th) {
                const fname = th.getAttribute('data-name');
                if (fname) {
                    rowData[fname] = this._getCellValue(cell, idx);
                }
            }
        });

        return rowData;
    }

    /**
     * Save or update entry in pending data (localStorage)
     * @param {string} tempId - Temporary row ID
     * @param {Object} rowData - Row data to save
     */
    _saveToPendingData(tempId, rowData) {
        const existingIdx = this._pendingData.findIndex(e => e._rowTempId === tempId);
        if (existingIdx === -1) {
            this._pendingData.push(rowData);
        } else {
            Object.assign(this._pendingData[existingIdx], rowData);
            this._pendingData[existingIdx]._timestamp = Date.now();
        }
        localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
    }

    // ==================== END VALIDATION HELPERS ====================

    // ==================== CENTRALIZED CELL/ROW HELPERS ====================

    /**
     * Set cell value with correct formatting (CENTRALIZED)
     * All paste, update, save functions should use this helper
     * @param {HTMLElement} cellNode - TD element
     * @param {string|number} value - Raw value to set
     * @param {number} colIndex - Column index
     * @param {boolean} updateDataTable - Whether to update DataTable internal data (default: true)
     */
    _setCellValueWithFormatting(cellNode, value, colIndex, updateDataTable = true) {
        if (!cellNode) return;

        const th = this.table.column(colIndex).header();
        const dataType = th ? th.getAttribute('data-type') : null;


        // Clean value based on type first
        const cleanedValue = this._cleanValueForType(value, dataType);

        // Handle formatting based on type
        if (dataType === 'currency' || dataType === 'percentage' || dataType === 'number') {
            // Numeric formatted types
            const formattedValue = this._formatValue(cleanedValue, dataType);
            cellNode.setAttribute('data-raw-value', cleanedValue);
            cellNode.textContent = formattedValue;
            cellNode.setAttribute('data-type', dataType);
        } else if (dataType === 'date') {
            // Date type - store ISO in raw-value, display formatted date
            const isoValue = this._parseDateToISO(cleanedValue);
            const formattedDate = this._formatDateForDisplay(isoValue);
            cellNode.setAttribute('data-raw-value', isoValue);
            cellNode.textContent = formattedDate;
            cellNode.setAttribute('data-type', dataType);
            cellNode.setAttribute('data-formatted', 'true');
        } else if (dataType === 'email') {
            // Email type - just display, validation on edit
            cellNode.setAttribute('data-raw-value', cleanedValue);
            cellNode.textContent = cleanedValue;
            cellNode.setAttribute('data-type', dataType);
        } else if (dataType === 'textarea') {
            // Textarea type - truncate for display, full value on hover
            cellNode.setAttribute('data-raw-value', cleanedValue);
            cellNode.setAttribute('data-full-text', cleanedValue); // IMPORTANT: update full text for edit mode
            const displayLength = (this.formats.textarea && this.formats.textarea.displayLength) || 50;
            if (cleanedValue.length > displayLength) {
                cellNode.textContent = cleanedValue.substring(0, displayLength) + '...';
                cellNode.setAttribute('title', cleanedValue); // Tooltip for hover
            } else {
                cellNode.textContent = cleanedValue;
                cellNode.removeAttribute('title');
            }
            cellNode.setAttribute('data-type', dataType);
            cellNode.classList.add('dt-textarea-cell');
        } else if (dataType === 'readonly') {
            // Readonly type - just display, no edit allowed
            cellNode.textContent = cleanedValue;
            cellNode.setAttribute('data-type', dataType);
            cellNode.classList.add('dt-readonly-cell');
        } else if (dataType === 'checkbox') {
            // Checkbox type - render checkbox element
            const isChecked = String(cleanedValue).toLowerCase() === 'true' || cleanedValue === '1';
            cellNode.setAttribute('data-checkbox-value', isChecked ? 'true' : 'false');

            // Check if checkbox already exists
            let checkbox = cellNode.querySelector('input[type="checkbox"]');
            if (!checkbox) {
                checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.className = 'dt-checkbox-display';
                cellNode.textContent = '';
                cellNode.appendChild(checkbox);

                // Add change handler
                checkbox.addEventListener('change', (e) => {
                    e.stopPropagation();
                    const newValue = checkbox.checked ? 'true' : 'false';
                    cellNode.setAttribute('data-checkbox-value', newValue);
                    this._handleCellChange(cellNode, colIndex, newValue);
                });
            }
            checkbox.checked = isChecked;
        } else if (dataType === 'select') {
            // Select type - render with arrow structure inline
            cellNode.innerHTML = '';
            cellNode.style.position = 'relative';
            cellNode.style.paddingRight = '24px';
            const contentSpan = document.createElement('span');
            contentSpan.className = 'dt-select-content';
            contentSpan.textContent = cleanedValue;
            contentSpan.title = cleanedValue;
            cellNode.appendChild(contentSpan);
            const arrow = document.createElement('span');
            arrow.className = 'dt-select-arrow';
            arrow.innerHTML = '▼';
            cellNode.appendChild(arrow);
            cellNode.setAttribute('data-select-cell', 'true');
        } else if (dataType === 'time') {
            // Time type - store raw value for reading
            cellNode.setAttribute('data-raw-value', cleanedValue);
            cellNode.textContent = cleanedValue;
            cellNode.setAttribute('data-type', dataType);
        } else {
            // Text and other types - also set data-raw-value for consistency
            if (cleanedValue) {
                cellNode.setAttribute('data-raw-value', cleanedValue);
            }
            cellNode.textContent = cleanedValue;
        }

        // Update DataTables internal data
        if (updateDataTable) {
            const cell = this.table.cell(cellNode);
            if (cell) {
                // Store raw value in DataTables
                cell.data(cleanedValue);

                // IMPORTANT: cell.data() triggers a redraw which overwrites textContent
                // For formatted types, we need to re-apply the visual formatting after DataTables update
                if (dataType === 'currency' || dataType === 'percentage' || dataType === 'number') {
                    const formattedValue = this._formatValue(cleanedValue, dataType);
                    cellNode.textContent = formattedValue;
                    cellNode.setAttribute('data-raw-value', cleanedValue);
                    cellNode.setAttribute('data-type', dataType);
                } else if (dataType === 'date') {
                    // Re-apply date formatting after DataTables redraw
                    // Use requestAnimationFrame to ensure it runs AFTER DataTables finishes
                    const isoValue = this._parseDateToISO(cleanedValue);
                    const formattedDate = this._formatDateForDisplay(isoValue);

                    // Apply immediately first
                    cellNode.textContent = formattedDate;
                    cellNode.setAttribute('data-raw-value', isoValue);
                    cellNode.setAttribute('data-type', dataType);
                    cellNode.setAttribute('data-formatted', 'true');

                    // Then defer to ensure it persists after any async DataTables redraw
                    requestAnimationFrame(() => {
                        if (cellNode && document.body.contains(cellNode)) {
                            cellNode.textContent = formattedDate;
                            cellNode.setAttribute('data-raw-value', isoValue);
                        }
                    });
                } else if (dataType === 'textarea') {
                    // Re-apply truncation for textarea after DataTables redraw
                    const displayLength = (this.formats.textarea && this.formats.textarea.displayLength) || 50;
                    if (cleanedValue.length > displayLength) {
                        cellNode.textContent = cleanedValue.substring(0, displayLength) + '...';
                        cellNode.setAttribute('title', cleanedValue);
                    } else {
                        cellNode.textContent = cleanedValue;
                        cellNode.removeAttribute('title');
                    }
                    cellNode.setAttribute('data-raw-value', cleanedValue);
                    cellNode.setAttribute('data-full-text', cleanedValue);
                    cellNode.setAttribute('data-type', dataType);
                    cellNode.classList.add('dt-textarea-cell');
                } else if (dataType === 'checkbox') {
                    // Re-render checkbox after data update
                    const isChecked = String(cleanedValue).toLowerCase() === 'true' || cleanedValue === '1';
                    cellNode.setAttribute('data-checkbox-value', isChecked ? 'true' : 'false');
                    let checkbox = cellNode.querySelector('input[type="checkbox"]');
                    if (!checkbox) {
                        checkbox = document.createElement('input');
                        checkbox.type = 'checkbox';
                        checkbox.className = 'dt-checkbox-display';
                        cellNode.textContent = '';
                        cellNode.appendChild(checkbox);
                        // Add change handler
                        checkbox.addEventListener('change', (e) => {
                            e.stopPropagation();
                            const newValue = checkbox.checked ? 'true' : 'false';
                            cellNode.setAttribute('data-checkbox-value', newValue);
                            this._handleCellChange(cellNode, cell.index().column, newValue);
                        });
                    }
                    checkbox.checked = isChecked;
                } else if (dataType === 'time') {
                    // Re-apply time value after DataTables redraw
                    cellNode.textContent = cleanedValue;
                    cellNode.setAttribute('data-raw-value', cleanedValue);
                    cellNode.setAttribute('data-type', dataType);
                } else if (dataType === 'select') {
                    // Re-apply select value with arrow after DataTables redraw
                    cellNode.innerHTML = '';
                    cellNode.style.position = 'relative';
                    cellNode.style.paddingRight = '24px';
                    const sContentSpan = document.createElement('span');
                    sContentSpan.className = 'dt-select-content';
                    sContentSpan.textContent = cleanedValue;
                    sContentSpan.title = cleanedValue;
                    cellNode.appendChild(sContentSpan);
                    const sArrow = document.createElement('span');
                    sArrow.className = 'dt-select-arrow';
                    sArrow.innerHTML = '▼';
                    cellNode.appendChild(sArrow);
                    cellNode.setAttribute('data-select-cell', 'true');
                    if (cleanedValue) {
                        cellNode.setAttribute('data-raw-value', cleanedValue);
                    }
                } else {
                    // Default text/email/other types - re-apply data-raw-value
                    cellNode.textContent = cleanedValue;
                    if (cleanedValue) {
                        cellNode.setAttribute('data-raw-value', cleanedValue);
                    }
                }
            }
        }
    }

    /**
     * Get raw value from cell (CENTRALIZED)
     * Handles formatted cells, checkboxes, selects
     * @param {HTMLElement} cellNode - TD element
     * @param {number} colIndex - Column index
     * @returns {string} Raw value
     */
    _getCellRawValue(cellNode, colIndex) {
        if (!cellNode) return '';

        const th = this.table.column(colIndex).header();
        const dataType = th ? th.getAttribute('data-type') : null;

        // Check for data-raw-value (formatted cells)
        if (cellNode.hasAttribute('data-raw-value')) {
            return cellNode.getAttribute('data-raw-value');
        }

        // Check for checkbox
        if (dataType === 'checkbox') {
            const checkbox = cellNode.querySelector('input[type="checkbox"]');
            if (checkbox) {
                return checkbox.checked ? 'true' : 'false';
            }
            // Fallback to attribute
            if (cellNode.hasAttribute('data-checkbox-value')) {
                return cellNode.getAttribute('data-checkbox-value');
            }
            // Fallback to text content
            const text = cellNode.textContent.trim().toLowerCase();
            return (text === 'true' || text === '1') ? 'true' : 'false';
        }

        // Check for select (remove arrow indicator)
        if (dataType === 'select') {
            return cellNode.textContent.replace(/[▼▲]/g, '').trim();
        }

        // Check for textarea (may have data-full-text or truncated display)
        if (dataType === 'textarea') {
            // Check data-full-text first (set when textarea value is truncated for display)
            if (cellNode.hasAttribute('data-full-text')) {
                return cellNode.getAttribute('data-full-text');
            }
            // Fallback to text content (remove "..." if truncated)
            let text = cellNode.textContent.trim();
            if (text.endsWith('...')) {
                // If truncated, prefer data-full-text or title attribute
                const title = cellNode.getAttribute('title');
                if (title) return title;
            }
            return text;
        }

        // Default: return text content
        return cellNode.textContent.trim();
    }

    /**
     * Get all field values from a row (CENTRALIZED)
     * @param {HTMLElement} rowNode - TR element
     * @returns {Object} { fieldName: value, ... }
     */
    _getRowData(rowNode) {
        if (!rowNode) return {};

        const data = {};
        const cells = rowNode.querySelectorAll('td');

        cells.forEach((td, idx) => {
            const th = this.table.column(idx).header();
            const fieldName = th ? th.getAttribute('data-name') : null;

            if (fieldName) {
                data[fieldName] = this._getCellRawValue(td, idx);
            }
        });

        return data;
    }

    /**
     * Clean/validate value based on data type (CENTRALIZED)
     * @param {string|number} value - Raw input value
     * @param {string} dataType - Column data type
     * @returns {string} Cleaned value
     */
    _cleanValueForType(value, dataType) {
        if (value === null || value === undefined) return '';

        let cleaned = String(value).trim();

        switch (dataType) {
            case 'currency':
            case 'percentage':
            case 'number':
                // Remove non-numeric characters except minus and decimal
                const formatConfig = this.formats[dataType] || {};
                const decimalSep = formatConfig.decimalSeparator || ',';

                // Remove all non-numeric except - and decimal separator
                // Also remove thousand separators
                let numClean = '';
                let hasDecimal = false;
                let hasMinus = false;

                for (let i = 0; i < cleaned.length; i++) {
                    const char = cleaned[i];
                    if (/[0-9]/.test(char)) {
                        numClean += char;
                    } else if (char === '-' && numClean.length === 0 && !hasMinus) {
                        numClean += char;
                        hasMinus = true;
                    } else if ((char === decimalSep || char === '.' || char === ',') && !hasDecimal) {
                        // Only add decimal if we already have digits (avoid ".50000000" from "Rp. 50.000.000")
                        if (numClean.length > 0) {
                            // Count consecutive digits after this separator
                            const remainingStr = cleaned.substring(i + 1);
                            const digitsAfterMatch = remainingStr.match(/^[0-9]+/);
                            const digitsAfter = digitsAfterMatch ? digitsAfterMatch[0].length : 0;

                            // Check if there are more of the same separator after
                            const hasMoreSameSep = remainingStr.includes(char);

                            // It's a decimal separator if:
                            // 1. No more of same separator after AND
                            // 2. 1-2 digits follow (like .50) OR it's the last char
                            // It's a thousand separator if 3 digits follow
                            if (!hasMoreSameSep && digitsAfter > 0 && digitsAfter <= 2) {
                                numClean += '.';
                                hasDecimal = true;
                            }
                            // Otherwise skip - it's a thousand separator
                        }
                    }
                }
                return numClean || '';

            case 'checkbox':
                // Normalize to 'true' or 'false'
                const lower = cleaned.toLowerCase();
                return (lower === 'true' || lower === '1' || lower === 'yes') ? 'true' : 'false';

            case 'select':
                // Remove arrow indicators
                return cleaned.replace(/[▼▲]/g, '').trim();

            case 'date':
                // Parse various date formats to ISO (yyyy-mm-dd)
                return this._parseDateToISO(cleaned);

            case 'email':
                // Just trim, validation done separately
                return cleaned.toLowerCase();

            case 'textarea':
            case 'readonly':
            case 'text':
            default:
                return cleaned;
        }
    }

    /**
     * Handle cell change event (for batch mode, save tracking)
     * @param {HTMLElement} cellNode - TD element
     * @param {number} colIndex - Column index
     * @param {string} value - New value
     */
    _handleCellChange(cellNode, colIndex, value) {
        const rowNode = cellNode.closest('tr');
        if (!rowNode) return;

        const rowId = rowNode.getAttribute('data-id');
        const th = this.table.column(colIndex).header();
        const fieldName = th ? th.getAttribute('data-name') : null;

        if (!fieldName) return;

        // Handle batch mode
        if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
            // Check if this is truly a pending row
            const isInPendingData = this._pendingData.some(row => row._rowTempId === rowId);
            const hasPendingAttr = rowNode.getAttribute('data-pending') === 'true';
            const isActuallyPending = rowId === 'new' || (rowId && rowId.startsWith('temp_') && (isInPendingData || hasPendingAttr));

            if (isActuallyPending) {
                // New/pending row - mark as pending
                this._markRowAsPending(rowNode, fieldName, value);
            } else if (rowId && rowId !== 'new') {
                // Existing DB row OR saved temp row - track edit
                this._trackEditedField(rowId, fieldName, value);
            }
        } else {
            // Direct mode - update cell via API
            if (rowId && rowId !== 'new') {
                this.updateCell(fieldName, rowId, value);
            }
        }
    }

    /**
     * Mark row as pending and save to localStorage (CENTRALIZED)
     * @param {HTMLElement} rowNode - TR element
     * @param {string} fieldName - Field that changed
     * @param {string} value - New value
     */
    _markRowAsPending(rowNode, fieldName, value) {
        if (!rowNode) return;

        rowNode.setAttribute('data-pending', 'true');

        // Generate temp_id if needed
        let rowId = rowNode.getAttribute('data-id');
        if (rowId === 'new' || !rowId) {
            rowId = 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
            rowNode.setAttribute('data-id', rowId);
            rowNode.removeAttribute('data-new-row');
        }

        // Check if row is complete before saving to localStorage
        const isComplete = this.isRowRequiredFieldsFilled(rowNode);

        // Update in-memory and optionally localStorage
        this._updateLocalStorageEntry(rowId, fieldName, value, isComplete);

        // If row is complete, check if we need to add an empty row
        if (isComplete) {

            if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                this.addEmptyRow();
            } else {
            }
        }

        return rowId;
    }

    /**
     * Check if cell is editable (not readonly, not No column)
     * @param {HTMLElement} cellNode - TD element
     * @param {number} colIndex - Column index
     * @returns {boolean}
     */
    _isEditableCell(cellNode, colIndex) {
        if (!cellNode) return false;

        // Skip No column
        if (this.enableColumnNumber && colIndex === 0) return false;

        // Check if cell is readonly
        if (cellNode.classList.contains('readonly')) return false;

        // Check if row is readonly
        const rowNode = cellNode.closest('tr');
        if (rowNode && rowNode.classList.contains('readonly')) return false;

        return true;
    }

    /**
     * Create modal for textarea editing
     * @param {string} fieldName - Field name
     * @param {string} rowId - Row ID
     * @param {string} currentValue - Current value
     * @returns {HTMLElement} Modal element
     */
    _createTextareaModal(fieldName, rowId, currentValue) {
        const formatConfig = this.formats.textarea || {};
        const rows = formatConfig.rows || 5;
        const maxLength = formatConfig.maxLength;

        // Create modal overlay
        const overlay = document.createElement('div');
        overlay.className = 'dt-textarea-modal-overlay';

        // Create modal container
        const modal = document.createElement('div');
        modal.className = 'dt-textarea-modal';

        // Header
        const header = document.createElement('div');
        header.className = 'dt-textarea-modal-header';
        header.innerHTML = `<span>${fieldName}</span>`;
        if (maxLength) {
            const charCount = document.createElement('span');
            charCount.className = 'dt-textarea-char-count';
            charCount.textContent = `${currentValue.length}/${maxLength}`;
            header.appendChild(charCount);
        }
        modal.appendChild(header);

        // Textarea
        const textarea = document.createElement('textarea');
        textarea.className = 'dt-textarea-editor';
        textarea.value = currentValue;
        textarea.rows = rows;
        if (maxLength) textarea.maxLength = maxLength;
        modal.appendChild(textarea);

        // Update char count on input
        if (maxLength) {
            textarea.addEventListener('input', () => {
                const charCount = modal.querySelector('.dt-textarea-char-count');
                if (charCount) {
                    charCount.textContent = `${textarea.value.length}/${maxLength}`;
                }
            });
        }

        // Footer with buttons
        const footer = document.createElement('div');
        footer.className = 'dt-textarea-modal-footer';

        const cancelBtn = document.createElement('button');
        cancelBtn.className = 'dt-textarea-btn dt-textarea-btn-cancel';
        cancelBtn.textContent = 'Cancel';
        cancelBtn.addEventListener('click', () => {
            overlay.remove();
            this.isEditMode = false;
        });

        const saveBtn = document.createElement('button');
        saveBtn.className = 'dt-textarea-btn dt-textarea-btn-save';
        saveBtn.textContent = 'Save';
        saveBtn.addEventListener('click', () => {
            const newValue = textarea.value;

            // Update cell using centralized helper
            const colIndex = this.activeCellNode ? this.activeCellNode.cellIndex : -1;
            if (colIndex >= 0) {
                this._setCellValueWithFormatting(this.activeCellNode, newValue, colIndex, true);
                this._handleCellChange(this.activeCellNode, colIndex, newValue);

                // For new row: trigger save check (same as doSave flow)
                const rowNode = this.activeCellNode.parentNode;
                const rowId = rowNode ? rowNode.getAttribute('data-id') : null;
                if (rowId === 'new' || rowNode.getAttribute('data-new-row') === 'true') {
                    const th = this.table.column(colIndex).header();
                    const fieldName = th ? th.getAttribute('data-name') : null;
                    setTimeout(() => {
                        this._handleNewRowSave(this.activeCellNode, fieldName, newValue);
                    }, 0);
                }
            }

            overlay.remove();
            this.isEditMode = false;
        });

        footer.appendChild(cancelBtn);
        footer.appendChild(saveBtn);
        modal.appendChild(footer);

        overlay.appendChild(modal);

        // Auto-focus textarea
        setTimeout(() => textarea.focus(), 50);

        // Close on overlay click
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) {
                overlay.remove();
                this.isEditMode = false;
            }
        });

        // Handle Escape key
        textarea.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                overlay.remove();
                this.isEditMode = false;
            }
        });

        // Ctrl+Enter to save (useful for textarea)
        textarea.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && e.ctrlKey) {
                e.preventDefault();
                saveBtn.click();
            }
        });

        return overlay;
    }

    // ==================== END CENTRALIZED HELPERS ====================

    /**
     * Prepare empty table with empty rows
     * Existing data in tbody will be preserved
     * @param {string} selector - CSS selector for table
     */
    _prepareEmptyTable(selector) {
        const tableNode = document.querySelector(selector);
        if (!tableNode) {
            console.warn('Table not found:', selector);
            return;
        }

        // Ensure tbody exists
        let tbody = tableNode.querySelector('tbody');
        if (!tbody) {
            tbody = document.createElement('tbody');
            // Insert after thead
            const thead = tableNode.querySelector('thead');
            if (thead) {
                thead.after(tbody);
            } else {
                tableNode.appendChild(tbody);
            }
        }

        // Count existing data rows
        const existingRows = tbody.querySelectorAll('tr');
        const existingCount = existingRows.length;

        // Count columns from thead
        const thead = tableNode.querySelector('thead tr');
        const columnCount = thead ? thead.children.length : 0;

        // Calculate how many empty rows need to be added
        // initialRows = minimum total rows (existing + empty)
        // If there's existing data, only add enough empty rows to reach initialRows
        // Always add at least 1 empty row
        const emptyRowsToAdd = Math.max(1, this.emptyTable.initialRows - existingCount);

        // Create empty rows
        for (let i = 0; i < emptyRowsToAdd; i++) {
            const tr = document.createElement('tr');
            tr.setAttribute('data-id', 'new');
            tr.setAttribute('data-new-row', 'true');

            for (let j = 0; j < columnCount; j++) {
                const td = document.createElement('td');
                td.textContent = '';
                tr.appendChild(td);
            }

            tbody.appendChild(tr);
        }

        console.log(`Prepared emptyTable: ${existingCount} existing rows, ${emptyRowsToAdd} empty rows added`);
    }

    /**
     * Inject No. column to DOM table BEFORE DataTable init
     * @param {string} selector - CSS selector for table
     */
    _injectColumnNumber(selector) {
        const tableNode = document.querySelector(selector);
        if (!tableNode) {
            console.warn('Table not found:', selector);
            return;
        }

        // Inject header <th data-no="no">No.</th>
        const thead = tableNode.querySelector('thead tr');
        if (thead && !thead.querySelector('th[data-no="no"]')) {
            const thNo = document.createElement('th');
            thNo.setAttribute('data-no', 'no');
            thNo.textContent = '';
            thead.insertBefore(thNo, thead.firstChild);
        }

        // Inject cell <td> in each tbody row with sequential number
        const tbodyRows = tableNode.querySelectorAll('tbody tr');
        tbodyRows.forEach((row, index) => {
            const tdNo = document.createElement('td');
            tdNo.textContent = index + 1;
            row.insertBefore(tdNo, row.firstChild);
        });

        // Inject tfoot if exists
        const tfoot = tableNode.querySelector('tfoot tr');
        if (tfoot) {
            const tfNo = document.createElement('th');
            tfNo.textContent = '';
            tfoot.insertBefore(tfNo, tfoot.firstChild);
        }
    }

    /**
     * Apply data-type attribute from header to cells
     * For styling based on type (e.g., number = right align)
     * Also apply formatting for number/currency/percentage/date types
     */
    _applyColumnTypes() {
        const tableNode = this.table.table().node();
        const headers = tableNode.querySelectorAll('thead th');

        headers.forEach((th, colIndex) => {
            const dataType = th.getAttribute('data-type');
            if (dataType) {
                // Apply to all cells in this column
                const cells = tableNode.querySelectorAll(`tbody td:nth-child(${colIndex + 1})`);
                cells.forEach(td => {
                    td.setAttribute('data-type', dataType);

                    // Skip if cell is empty
                    const textContent = td.textContent.trim();
                    if (!textContent) {
                        return;
                    }

                    // Get the DataTables cell object and row index
                    const dtCell = this.table.cell(td);
                    if (!dtCell || !dtCell.index()) return;

                    const rowIndex = dtCell.index().row;
                    const cacheKey = `${rowIndex}_${colIndex}`;

                    // Determine the raw value using priority:
                    // 1. Check _rawValueCache (most reliable - persists across redraws)
                    // 2. Check DOM data-raw-value attribute
                    // 3. Check if DataTables internal data is clean number
                    // 4. Parse from textContent as last resort

                    if (dataType === 'date') {
                        // Date type handling
                        let isoValue;
                        const cachedValue = this._rawValueCache.get(cacheKey);
                        const existingRawValue = td.getAttribute('data-raw-value');
                        const dtData = dtCell.data();

                        if (cachedValue && this._isValidDate(cachedValue)) {
                            isoValue = cachedValue;
                        } else if (existingRawValue && this._isValidDate(existingRawValue)) {
                            isoValue = existingRawValue;
                            this._rawValueCache.set(cacheKey, isoValue);
                        } else if (dtData && this._isValidDate(String(dtData))) {
                            isoValue = String(dtData);
                            this._rawValueCache.set(cacheKey, isoValue);
                        } else {
                            isoValue = this._parseDateToISO(textContent);
                            this._rawValueCache.set(cacheKey, isoValue);
                        }

                        td.setAttribute('data-raw-value', isoValue);
                        td.setAttribute('data-formatted', 'true');
                        td.textContent = this._formatDateForDisplay(isoValue);

                    } else if (this.formats[dataType]) {
                        // Number types (currency, percentage, number)
                        let cleanedValue;
                        const cachedValue = this._rawValueCache.get(cacheKey);
                        const existingRawValue = td.getAttribute('data-raw-value');
                        const dtData = dtCell.data();

                        if (cachedValue !== undefined && cachedValue !== null && cachedValue !== '') {
                            // Use cached value (most reliable)
                            cleanedValue = cachedValue;
                        } else if (existingRawValue !== null && existingRawValue !== undefined && existingRawValue !== '') {
                            // Use DOM attribute and cache it
                            cleanedValue = existingRawValue;
                            this._rawValueCache.set(cacheKey, cleanedValue);
                        } else if (dtData !== null && dtData !== undefined) {
                            // Use DataTables internal data
                            const dtDataStr = String(dtData);
                            if (/^-?\d+\.?\d*$/.test(dtDataStr)) {
                                // It's a clean number like "50000000"
                                cleanedValue = dtDataStr;
                            } else {
                                // It's formatted string, need to clean
                                cleanedValue = this._cleanValueForType(dtDataStr, dataType);
                            }
                            this._rawValueCache.set(cacheKey, cleanedValue);
                        } else {
                            // Fallback: parse from textContent
                            cleanedValue = this._cleanValueForType(textContent, dataType);
                            this._rawValueCache.set(cacheKey, cleanedValue);
                        }

                        td.setAttribute('data-raw-value', cleanedValue);
                        td.textContent = this._formatValue(cleanedValue, dataType);
                    }
                });
            }
        });
    }

    /**
     * Initialize events on DataTable and keyboard
     */
    initialize() {
        // ===== Backward Compatibility: Create property aliases to _state =====
        // These allow existing code (this._selectedCells) to work while using centralized state
        // For new code, use this._getState() and this._setState()

        // Edit mode aliases
        Object.defineProperty(this, 'isEditMode', {
            get: () => this._state.isEditMode,
            set: (v) => this._state.isEditMode = v,
            configurable: true
        });
        Object.defineProperty(this, 'currentCell', {
            get: () => this._state.currentCell,
            set: (v) => this._state.currentCell = v,
            configurable: true
        });
        Object.defineProperty(this, 'focusedRow', {
            get: () => this._state.focusedRow,
            set: (v) => this._state.focusedRow = v,
            configurable: true
        });

        // Cell selection aliases
        Object.defineProperty(this, '_selectionStart', {
            get: () => this._state.selectionStart,
            set: (v) => this._state.selectionStart = v,
            configurable: true
        });
        Object.defineProperty(this, '_selectionEnd', {
            get: () => this._state.selectionEnd,
            set: (v) => this._state.selectionEnd = v,
            configurable: true
        });
        Object.defineProperty(this, '_selectedCells', {
            get: () => this._state.selectedCells,
            set: (v) => this._state.selectedCells = v,
            configurable: true
        });

        // Fill handle aliases
        Object.defineProperty(this, '_fillHandleElement', {
            get: () => this._state.fillHandleElement,
            set: (v) => this._state.fillHandleElement = v,
            configurable: true
        });
        Object.defineProperty(this, '_isFillDragging', {
            get: () => this._state.isFillDragging,
            set: (v) => this._state.isFillDragging = v,
            configurable: true
        });
        Object.defineProperty(this, '_fillStartCells', {
            get: () => this._state.fillStartCells,
            set: (v) => this._state.fillStartCells = v,
            configurable: true
        });
        Object.defineProperty(this, '_fillOriginalBounds', {
            get: () => this._state.fillOriginalBounds,
            set: (v) => this._state.fillOriginalBounds = v,
            configurable: true
        });
        Object.defineProperty(this, '_fillSourceData', {
            get: () => this._state.fillSourceData,
            set: (v) => this._state.fillSourceData = v,
            configurable: true
        });
        Object.defineProperty(this, '_fillTargetEnd', {
            get: () => this._state.fillTargetEnd,
            set: (v) => this._state.fillTargetEnd = v,
            configurable: true
        });

        // Row selection aliases
        Object.defineProperty(this, '_isRowSelecting', {
            get: () => this._state.isRowSelecting,
            set: (v) => this._state.isRowSelecting = v,
            configurable: true
        });
        Object.defineProperty(this, '_rowSelectionStart', {
            get: () => this._state.rowSelectionStart,
            set: (v) => this._state.rowSelectionStart = v,
            configurable: true
        });
        Object.defineProperty(this, '_selectedRows', {
            get: () => this._state.selectedRows,
            set: (v) => this._state.selectedRows = v,
            configurable: true
        });
        Object.defineProperty(this, '_hasClipboardData', {
            get: () => this._state.hasClipboardData,
            set: (v) => this._state.hasClipboardData = v,
            configurable: true
        });
        Object.defineProperty(this, '_clipboardRowCount', {
            get: () => this._state.clipboardRowCount,
            set: (v) => this._state.clipboardRowCount = v,
            configurable: true
        });

        // Keyboard state aliases
        Object.defineProperty(this, '_isShiftPressed', {
            get: () => this._state.isShiftPressed,
            set: (v) => this._state.isShiftPressed = v,
            configurable: true
        });
        Object.defineProperty(this, '_isDragging', {
            get: () => this._state.isDragging,
            set: (v) => this._state.isDragging = v,
            configurable: true
        });

        // UI state aliases
        Object.defineProperty(this, '_highlightedRow', {
            get: () => this._state.highlightedRow,
            set: (v) => this._state.highlightedRow = v,
            configurable: true
        });

        // Bind fill drag handlers
        this._onFillDragMove = this._handleFillDragMove.bind(this);
        this._onFillDragEnd = this._handleFillDragEnd.bind(this);

        // Bind row selection handlers
        this._onRowSelectMove = this._handleRowSelectMove.bind(this);
        this._onRowSelectEnd = this._handleRowSelectEnd.bind(this);

        // ===== Event Management System =====
        // Registry to track all event listeners for proper cleanup
        this._eventListeners = [];

        // Initialize row selection events
        this._initRowSelectionEvents();

        // Re-apply formatting after every DataTables draw (pagination, sort, filter)
        this.table.on('draw.dt', () => {
            this._applyColumnTypes();
            // Recalculate formula columns after column types are applied
            this._recalculateAllFormulas();
        });

        // Keyboard keyTable mode - SELECTION MODE (not direct edit)
        this.table.on('key-focus', (e, datatable, cell) => {
            // Cancel any pending refocus from select dropdown
            // This prevents focus jumping back to select cell after user navigates away
            this._selectRefocusCancelled = true;

            // Get previous row info before updating
            const prevRowIndex = this.focusedRow;
            const newRowIndex = cell.index().row;

            // Process pending save from checkbox/select/dropdown/delete if any
            if (this._pendingSave) {
                const ps = this._pendingSave;

                // Validate empty value if column doesn't allow empty
                if (ps.newValue === '' && ps.allowEmpty === false) {
                    console.log('Validation failed - column', ps.fieldName, 'does not allow empty');
                    // Restore old value
                    if (ps.oldValue !== undefined) {
                        this._setCellValueWithFormatting(ps.cellNode, ps.oldValue, ps.cellNode.cellIndex, false);
                    }
                    // Show error indicator
                    ps.cellNode.classList.add('dt-error');
                    setTimeout(() => ps.cellNode.classList.remove('dt-error'), 2000);
                    this._pendingSave = null;
                    // Continue with focus update, don't return
                } else {
                    // Update cell display
                    this._setCellValueWithFormatting(ps.cellNode, ps.newValue, ps.cellNode.cellIndex, true);

                    // Trigger save based on rowId
                    if (ps.rowId !== 'new') {
                        // EXISTING ROW: Update to server
                        this.updateCell(ps.fieldName, ps.rowId, ps.newValue);
                    } else {
                        // NEW ROW: Check if row is complete and save
                        const rowNode = ps.cellNode.parentNode;
                        if (rowNode && this.isRowRequiredFieldsFilled(rowNode)) {
                            // Direct mode: save new row
                            if (!(this.emptyTable.enabled && this.emptyTable.saveMode === 'batch')) {
                                Promise.resolve(this.saveNewRow(rowNode))
                                    .finally(() => {
                                        if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                                            this.addEmptyRow();
                                        }
                                    });
                            } else {
                                // Batch mode: route to batch save
                                this._handleNewRowSave(ps.cellNode, ps.fieldName, ps.newValue);
                            }
                        }
                    }

                    this._pendingSave = null; // Clear pending
                }
            }

            // ROW CHANGE DETECTION: Check if user moved from new row to another row
            // This handles the case where user filled all fields via normal input (blur saves)
            if (prevRowIndex !== undefined && prevRowIndex !== newRowIndex) {
                const prevRowNode = this.table.row(prevRowIndex).node();
                if (prevRowNode && prevRowNode.getAttribute('data-id') === 'new') {
                    // Check if previous new row is now complete
                    if (this.isRowRequiredFieldsFilled(prevRowNode)) {
                        // Direct mode: save new row
                        if (!(this.emptyTable.enabled && this.emptyTable.saveMode === 'batch')) {
                            Promise.resolve(this.saveNewRow(prevRowNode))
                                .finally(() => {
                                    if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                                        this.addEmptyRow();
                                    }
                                });
                        }
                    }
                }
            }

            this.focusedRow = cell.index().row;
            this.currentCell = cell;
            this.isEditMode = false; // Reset to selection mode

            // Remove row highlight when focus changes
            this._removeRowHighlight();

            // Reset multi-cell selection when focus changes (without shift)
            if (!this._isShiftPressed) {
                this._clearCellSelection();
                this._selectionStart = { row: cell.index().row, col: cell.index().column };
                this._selectedCells = [{ row: cell.index().row, col: cell.index().column }];
                // Show fill handle for single cell focus
                this._updateFillHandle(cell.index().row, cell.index().column);
            }

            let colIndex = cell.index().column;
            // Prevent edit on "No" column and automatically focus next column
            if (this.enableColumnNumber && colIndex === 0) {
                const rowIdx = cell.index().row;
                const nextCell = this.table.cell(rowIdx, 1);
                if (nextCell && nextCell.node()) {
                    nextCell.focus();
                }
                return;
            }

            // Add visual highlight for selection mode
            this._highlightCell(cell.node());
        })
            .on('key-blur', (e, datatable, cell) => {
                this.lastFocusedRow = cell.index().row;
                // Remove highlight on blur
                this._removeHighlight(cell.node());
                // Remove fill handle (green dot)
                this._removeFillHandle();
            });

        // Keyboard handler for Selection Mode
        this.table.on('key', (e, datatable, key, cell, originalEvent) => {
            // If already in edit mode, let editor handle
            if (this.isEditMode) return;

            const colIndex = cell.index().column;
            const rowIndex = cell.index().row;

            // Skip No column
            if (this.enableColumnNumber && colIndex === 0) return;

            // ===== EMPTY TABLE: Auto-add row is already handled by document keydown listener =====
            // Don't handle here to prevent double-trigger

            // Enter = enter Edit Mode with existing content
            if (key === 13) { // Enter
                originalEvent.preventDefault();

                // Check column type
                const th = this.table.column(colIndex).header();
                const fieldType = th ? th.getAttribute('data-type') : null;

                if (fieldType === 'checkbox') {
                    this._toggleCheckbox(cell);
                } else {
                    this.isEditMode = true;
                    this.activateEditCell(cell, false); // don't clear content
                }
                return;
            }

            // Delete/Backspace = clear cell content
            if (key === 46 || key === 8) { // Delete or Backspace
                originalEvent.preventDefault();
                this._clearCellContent(cell);
                return;
            }

            // Space = toggle checkbox (for checkbox columns only)
            if (key === 32) { // Spacebar
                const th = this.table.column(colIndex).header();
                const fieldType = th ? th.getAttribute('data-type') : null;

                if (fieldType === 'checkbox') {
                    originalEvent.preventDefault();
                    this._toggleCheckbox(cell);
                    return;
                }
            }

            // Printable characters (letters, numbers, etc.) = Replace mode
            // Key codes: 48-57 (0-9), 65-90 (A-Z), 96-111 (numpad), 186-222 (symbols)
            if (this._isPrintableKey(key)) {
                originalEvent.preventDefault();
                this.isEditMode = true;
                this.activateEditCell(cell, true); // true = clear content, type new

                // Dispatch pressed key to editor
                setTimeout(() => {
                    const editor = cell.node().querySelector('input, select');
                    if (editor && editor.tagName.toLowerCase() === 'input') {
                        // For printable characters, add to input
                        const char = originalEvent.key;
                        if (char.length === 1) {
                            editor.value = char;
                            // Trigger input event
                            editor.dispatchEvent(new Event('input', { bubbles: true }));
                        }
                    }
                }, 10);
                return;
            }
        });

        // Double click = enter Edit Mode
        this.table.on('dblclick', 'td', (e) => {
            if (e.target.tagName.toLowerCase() !== 'td') return;
            let cell = this.table.cell(e.target);
            if (cell) {
                this.isEditMode = true;
                this.activateEditCell(cell, false); // Edit existing content
            }
        });

        // Single click = Selection Mode (focus only, no edit)
        this.table.on('click', 'td', (e) => {
            // Skip if already in edit mode
            if (this.isEditMode) return;

            // Get td element - could be clicked on child element
            const tdElement = e.target.tagName.toLowerCase() === 'td' ? e.target : e.target.closest('td');
            if (!tdElement) return;

            let cell = this.table.cell(tdElement);
            if (cell) {
                const colIndex = cell.index().column;

                // If click on No column (column 0) and enableColumnNumber is active
                if (this.enableColumnNumber && colIndex === 0) {
                    const rowNode = tdElement.closest('tr');

                    // Toggle: if this row is already highlighted, release and focus the second column
                    if (rowNode && rowNode.classList.contains('row-highlight')) {
                        this._removeRowHighlight();
                        // Focus to the second column (column 1)
                        const rowIdx = cell.index().row;
                        const nextCell = this.table.cell(rowIdx, 1);
                        nextCell.focus();

                        // Update selection and fill handle for the new cell
                        this._selectionStart = { row: rowIdx, col: 1 };
                        this._selectedCells = [{ row: rowIdx, col: 1 }];
                        this._updateFillHandle(rowIdx, 1);
                    } else {
                        // Highlight new row
                        this._highlightRow(rowNode);
                    }
                    return;
                }

                const rowNode = tdElement.closest('tr');
                const wasHighlighted = rowNode && rowNode.classList.contains('row-highlight');

                // Remove row highlight when clicking normal cell
                this._removeRowHighlight();

                // Clear multi-row selection when clicking on a cell
                if (this._selectedRows && this._selectedRows.length > 0) {
                    this._clearRowSelection();
                }

                // Special case for select columns: single click activates edit mode and opens dropdown
                const th = this.table.column(colIndex).header();
                const fieldType = th ? th.getAttribute('data-type') : null;
                if (fieldType === 'select') {
                    this.isEditMode = true;
                    this.activateEditCell(cell, false);
                    return;
                }

                // Just focus, don't edit
                cell.focus();

                // If from row highlight, update selection and focus class manually
                if (wasHighlighted) {
                    const rowIdx = cell.index().row;
                    const cellNode = cell.node();

                    // Manual add focus class (KeyTable may not add it due to timing)
                    cellNode.classList.add('focus');

                    this._selectionStart = { row: rowIdx, col: colIndex };
                    this._selectedCells = [{ row: rowIdx, col: colIndex }];
                    this._updateFillHandle(rowIdx, colIndex);
                }
            }
        });

        // ===== Keydown listener for auto-add row in Selection Mode =====
        // Using capture phase on document to intercept BEFORE KeyTable handle
        this._lastAddRowTime = 0; // Debounce timestamp
        const tableNode = this.table.table().node();
        const tableWrapper = this.table.table().container();

        // Capture phase = true untuk intercept sebelum KeyTable
        // Using _addEvent for proper cleanup tracking
        this._addEvent(document, 'keydown', (e) => {
            // Skip if emptyTable is not enabled
            if (!this.emptyTable.enabled) return;

            // Only handle Arrow Down and Tab
            if (e.key !== 'ArrowDown' && e.key !== 'Tab') return;

            // Skip if in Edit Mode (already handled by editor in attachEditorEvent)
            if (this.isEditMode) return;

            // Find cell with 'focus' class in this table (DOM-based)
            const focusedCellNode = tableNode.querySelector('td.focus');
            if (!focusedCellNode) return;

            // Skip if currently inserting row (prevent auto-add during insert)
            if (this._isInsertingRow) return;

            // Ensure focus is in this table
            if (!tableWrapper.contains(document.activeElement) && !tableNode.contains(focusedCellNode)) {
                return;
            }

            // Get row from focused cell
            const focusedRow = focusedCellNode.closest('tr');
            if (!focusedRow) return;

            // Calculate row index from DOM - BEFORE KeyTable navigation
            const tbody = tableNode.querySelector('tbody');
            const allRows = Array.from(tbody.querySelectorAll('tr'));
            const currentRowIndex = allRows.indexOf(focusedRow);
            const totalRows = allRows.length;

            // Check if CURRENTLY on last row (before navigation)
            const isCurrentlyLastRow = currentRowIndex === totalRows - 1;

            // For Tab, also check if last cell
            const allCellsInRow = Array.from(focusedRow.querySelectorAll('td'));
            const currentColIndex = allCellsInRow.indexOf(focusedCellNode);
            const totalCols = allCellsInRow.length;
            const isCurrentlyLastCol = currentColIndex === totalCols - 1;

            // Debounce: prevent rapid add (within 300ms)
            const now = Date.now();
            if (now - this._lastAddRowTime < 300) return;

            // Arrow Down on CURRENT last row
            if (e.key === 'ArrowDown' && isCurrentlyLastRow) {
                e.preventDefault();
                e.stopPropagation();
                this._lastAddRowTime = now;

                this._addNewEmptyRow();
                setTimeout(() => {
                    const newRowIdx = this.table.rows().count() - 1;
                    const firstEditableCol = this.enableColumnNumber ? 1 : 0;
                    const newCell = this.table.cell(newRowIdx, firstEditableCol);
                    if (newCell && newCell.node()) {
                        newCell.focus();
                        this.currentCell = newCell;
                    }
                }, 50);
                return;
            }

            // Tab on last cell of CURRENT last row
            if (e.key === 'Tab' && isCurrentlyLastRow && isCurrentlyLastCol) {
                e.preventDefault();
                e.stopPropagation();
                this._lastAddRowTime = now;

                this._addNewEmptyRow();
                setTimeout(() => {
                    const newRowIdx = this.table.rows().count() - 1;
                    const firstEditableCol = this.enableColumnNumber ? 1 : 0;
                    const newCell = this.table.cell(newRowIdx, firstEditableCol);
                    if (newCell && newCell.node()) {
                        newCell.focus();
                        this.currentCell = newCell;
                    }
                }, 50);
                return;
            }
        }, true); // CAPTURE PHASE

        // ===== Keyboard shortcuts for Copy/Paste Row =====
        this._addEvent(document, 'keydown', (e) => {
            // Skip if in Edit Mode
            if (this.isEditMode) return;

            // Check if there's a highlighted row
            const highlightedRow = tableNode.querySelector('tr.row-highlight');

            // Find focused cell - check multiple possible focus classes and currentCell
            let focusedCellNode = tableNode.querySelector('td.focus, td.dt-focused, td.focus-cell');

            // Fallback: use currentCell if available
            if (!focusedCellNode && this.currentCell) {
                try {
                    focusedCellNode = this.currentCell.node();
                } catch (e) {
                    focusedCellNode = null;
                }
            }

            // Fallback: check _selectedCells for single cell
            if (!focusedCellNode && this._selectedCells && this._selectedCells.length === 1) {
                const sel = this._selectedCells[0];
                try {
                    focusedCellNode = this.table.cell(sel.row, sel.col).node();
                } catch (e) {
                    focusedCellNode = null;
                }
            }

            // Ctrl+C: Copy - priority: 1) multi-row selection, 2) multi-cell selection, 3) highlighted row, 4) single cell
            if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
                // Check if there's multi-row selection (from No. column)
                if (this._selectedRows && this._selectedRows.length > 0) {
                    e.preventDefault();
                    this._copyMultipleRows();
                    return;
                }

                // Check if there's multi-cell selection
                if (this._selectedCells && this._selectedCells.length > 1) {
                    e.preventDefault();
                    this._copyCellsToClipboard();
                    return;
                }

                // Highlighted row: copy entire row
                if (highlightedRow) {
                    e.preventDefault();
                    this._copyRowToClipboard(highlightedRow);
                    return;
                }

                // Single cell focused: copy cell value
                if (focusedCellNode) {
                    e.preventDefault();
                    this._copySingleCell(focusedCellNode);
                    return;
                }

                // No focus, let browser handle (regular text copy if any)
                return;
            }

            // Ctrl+V: Paste Row (only for empty or pending rows)
            if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
                // Handle multi-row selection - paste to multiple selected rows
                if (this._selectedRows && this._selectedRows.length > 0) {
                    e.preventDefault();
                    this._pasteMultipleRows();
                    return;
                }

                // Check if multi-cell selection exists - paste to selected cells
                if (this._selectedCells && this._selectedCells.length > 1) {
                    e.preventDefault();
                    this._pasteToSelectedCells();
                    return;
                }

                // Single cell paste: if there's a focused cell (not row highlight)
                if (focusedCellNode && !highlightedRow) {
                    // Check if cell or column is readonly
                    const isCellReadonly = focusedCellNode.getAttribute('data-readonly') === 'true';
                    const rowNode = focusedCellNode.closest('tr');
                    const isRowReadonlyCell = rowNode && rowNode.getAttribute('data-readonly-row') === 'true';
                    const colIdx = focusedCellNode.cellIndex;
                    const th = this.table.column(colIdx).header();
                    const fieldType = th ? th.getAttribute('data-type') : null;
                    const isColumnReadonly = fieldType === 'readonly' || fieldType === 'formula';

                    if (isCellReadonly || isRowReadonlyCell || isColumnReadonly) {
                        e.preventDefault();
                        return;
                    }

                    e.preventDefault();
                    const cellIdx = this.table.cell(focusedCellNode).index();
                    this._pasteSingleCell(focusedCellNode, cellIdx);
                    return;
                }

                // Determine target row for row paste
                let rowNode = highlightedRow;
                if (!rowNode && focusedCellNode) {
                    rowNode = focusedCellNode.closest('tr');
                }
                if (!rowNode) return;

                // Allow paste to any row (empty, pending, or existing)
                e.preventDefault();
                this._pasteRowFromClipboard(rowNode);
            }
        });

        // ===== Track Shift key state =====
        this._isShiftPressed = false;
        this._addEvent(document, 'keydown', (e) => {
            if (e.key === 'Shift') {
                this._isShiftPressed = true;
            }
        });
        this._addEvent(document, 'keyup', (e) => {
            if (e.key === 'Shift') {
                this._isShiftPressed = false;
            }
        });

        // ===== Shift+Arrow untuk extend selection =====
        this._addEvent(document, 'keydown', (e) => {
            if (!this._isShiftPressed) return;
            if (this.isEditMode) return;
            if (!['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) return;

            // Ensure there's a selection start
            if (!this._selectionStart) {
                if (this.currentCell) {
                    this._selectionStart = {
                        row: this.currentCell.index().row,
                        col: this.currentCell.index().column
                    };
                } else {
                    return;
                }
            }

            e.preventDefault();

            // Calculate new end position
            let endRow = this._selectionEnd ? this._selectionEnd.row : this._selectionStart.row;
            let endCol = this._selectionEnd ? this._selectionEnd.col : this._selectionStart.col;

            const maxRow = this.table.rows().count() - 1;
            const maxCol = this.table.columns().count() - 1;

            switch (e.key) {
                case 'ArrowUp':
                    endRow = Math.max(0, endRow - 1);
                    break;
                case 'ArrowDown':
                    endRow = Math.min(maxRow, endRow + 1);
                    break;
                case 'ArrowLeft':
                    endCol = Math.max(this.enableColumnNumber ? 1 : 0, endCol - 1);
                    break;
                case 'ArrowRight':
                    endCol = Math.min(maxCol, endCol + 1);
                    break;
            }

            this._selectionEnd = { row: endRow, col: endCol };
            this._updateCellSelection();
        });

        // ===== Mouse drag handler untuk drag selection =====
        this._isDragging = false;

        // Mousedown: start selection
        this._addEvent(tableNode, 'mousedown', (e) => {
            const td = e.target.closest('td');
            if (!td) return;

            // Ignore if it's a right click
            if (e.button !== 0) return;

            const rowNode = td.closest('tr');
            if (!rowNode) return;

            const rowIdx = this.table.row(rowNode).index();
            const colIdx = Array.from(rowNode.children).indexOf(td);

            // Skip No. column - let click handler handle it
            if (this.enableColumnNumber && colIdx === 0) {
                return;
            }

            // Skip if row is highlighted - let click handler remove highlight first
            if (rowNode.classList.contains('row-highlight')) {
                return;
            }

            if (this._isShiftPressed && this._selectionStart) {
                // Shift+Click: extend selection dari start ke clicked cell
                e.preventDefault();
                // Clear row selection when extending cell selection
                if (this._selectedRows && this._selectedRows.length > 0) {
                    this._clearRowSelection();
                }
                this._selectionEnd = { row: rowIdx, col: colIdx };
                this._updateCellSelection();
            } else {
                // Start new selection / drag
                this._clearCellSelection();
                // Clear row selection when starting new cell selection
                if (this._selectedRows && this._selectedRows.length > 0) {
                    this._clearRowSelection();
                }
                this._selectionStart = { row: rowIdx, col: colIdx };
                this._selectionEnd = null;
                this._isDragging = true;

                // Focus the first cell with KeyTable
                this.table.cell(rowIdx, colIdx).focus();

                // Store anchor cell for visual distinction
                this._anchorCell = { row: rowIdx, col: colIdx };
            }
        });

        // Mousemove: extend selection while dragging
        this._addEvent(tableNode, 'mousemove', (e) => {
            if (!this._isDragging) return;
            if (!this._selectionStart) return;

            const td = e.target.closest('td');
            if (!td) return;

            const rowNode = td.closest('tr');
            if (!rowNode) return;

            const rowIdx = this.table.row(rowNode).index();
            const colIdx = Array.from(rowNode.children).indexOf(td);

            // Only update if position changed
            if (!this._selectionEnd ||
                this._selectionEnd.row !== rowIdx ||
                this._selectionEnd.col !== colIdx) {
                this._selectionEnd = { row: rowIdx, col: colIdx };
                this._updateCellSelection();
            }
        });

        // Mouseup: end selection
        this._addEvent(document, 'mouseup', (e) => {
            if (this._isDragging) {
                this._isDragging = false;
            }
        });

        // On load, only add row if on last page
        // and no empty row exists (avoid duplication if startEmpty is active)
        if (this.allowAddEmptyRow === true && this.isOnLastPage() && !this.hasEmptyRow()) {
            this.addEmptyRow();
        }


        // Global Keydown Handler for Escape (Clear Selection)
        // Note: Context menu escape is handled in showContextMenu
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                // Ignore if in edit mode (handled by editor)
                if (this.isEditMode) return;

                // Ignore if context menu is open (handled by context menu)
                if (document.getElementById('custom-contextmenu')) return;

                // Close dropdown if open
                const dropdown = document.querySelector('.dt-dropdown-options');
                if (dropdown) {
                    dropdown.remove();
                    const searchDropdown = document.querySelector('.dt-searchable-dropdown');
                    if (searchDropdown) searchDropdown.remove();
                    return;
                }

                // Clear all selections
                this._clearCellSelection();
                this._clearRowSelection();
                this._removeFillHandle();
                this._removeRowHighlight();

                // Remove visual focus classes
                const tableNode = this.table.table().node();
                if (tableNode) {
                    tableNode.querySelectorAll('.focus, .cell-focus, .dt-cell-selected, .selected').forEach(el => {
                        el.classList.remove('focus', 'cell-focus', 'dt-cell-selected', 'selected');
                    });
                }

                // Blur DataTables KeyTable focus
                if (this.table.cell && this.table.cell.blur) {
                    this.table.cell.blur();
                }

                this.currentCell = null;
            }
        });

        let contextMenuTimeout = null;
        this.table.on('contextmenu', 'td', (e) => {
            e.preventDefault();
            e.stopPropagation();

            // Debounce context menu untuk mencegah double right-click
            if (contextMenuTimeout) return;
            contextMenuTimeout = setTimeout(() => { contextMenuTimeout = null; }, 300);

            // Check if this is a regular cell (not No column)
            const td = e.currentTarget;
            const colIdx = td.cellIndex;
            const isNoColumn = this.enableColumnNumber && colIdx === 0;

            // Check if clicked cell is part of current selection
            const rowNode = td.closest('tr');
            const rowIdx = this.table.row(rowNode).index();

            // Check if cell is in multi-cell selection
            const isInCellSelection = this._selectedCells && this._selectedCells.some(
                c => c.row === rowIdx && c.col === colIdx
            );

            // Check if row is in multi-row selection
            const isInRowSelection = this._selectedRows && this._selectedRows.includes(rowIdx);

            // ===== HANDLING FOR No. COLUMN =====
            if (isNoColumn) {
                // If row is NOT in current row selection, switch to this row
                if (!isInRowSelection) {

                    // Clear old row selection
                    this._clearRowSelection();

                    // Clear cell selection
                    this._clearCellSelection();

                    // Select this row
                    this._selectedRows = [rowIdx];
                    this._highlightSelectedRows();
                } else {
                }
                // Show context menu (handled below)
            }
            // ===== HANDLING FOR REGULAR CELLS =====
            else if (!isNoColumn) {
                // If clicked cell is INSIDE current selection, keep the selection
                if (isInCellSelection || isInRowSelection) {
                    // Don't clear anything, just show context menu
                }
                // If clicked cell is OUTSIDE current selection, clear and focus new cell
                else {

                    // Clear multi-row selection
                    if (this._selectedRows && this._selectedRows.length > 0) {
                        this._clearRowSelection();
                    }

                    // Clear row highlight
                    this._removeRowHighlight();

                    // Clear ALL focus classes from cells
                    const tableNode = this.table.table().node();
                    tableNode.querySelectorAll('td.focus, td.cell-focus, td.cell-selected, td.dt-cell-selected').forEach(c => {
                        c.classList.remove('focus', 'cell-focus', 'cell-selected', 'dt-cell-selected');
                    });

                    // Clear range selection highlight (removes dt-range-* classes)
                    this._clearCellSelectionHighlight();

                    // Remove old fill handle
                    this._removeFillHandle();

                    // Clear multi-cell selection
                    this._selectedCells = [];

                    // Focus the clicked cell
                    const cell = this.table.cell(td);
                    if (cell) {
                        const cellIdx = cell.index();

                        // Blur KeyTable internal focus first
                        if (this.table.cell.blur) {
                            this.table.cell.blur();
                        }

                        // Add focus classes to clicked cell
                        td.classList.add('focus');
                        td.classList.add('dt-cell-range-selected');
                        td.classList.add('dt-range-top', 'dt-range-bottom', 'dt-range-left', 'dt-range-right');

                        // Update internal state
                        this.currentCell = cell;
                        this._selectionStart = { row: cellIdx.row, col: cellIdx.column };
                        this._selectedCells = [{ row: cellIdx.row, col: cellIdx.column }];

                        // Update fill handle (dot at bottom-right)
                        this._updateFillHandle(cellIdx.row, cellIdx.column);
                    }
                }
            }

            // Show context menu for all columns
            this._buildContextMenu(e, td);
        });

        // On every page change, check if on last page, then add empty row if not exists
        this.table.on('draw', () => {
            // PATCH: Check anti storm!
            if (
                this.allowAddEmptyRow &&
                !this.isAddingEmptyRow && // only if not in the process of adding
                !this.isCheckboxChanging && // skip during checkbox change
                this.isOnLastPage() &&
                this.isShouldAddEmptyRow() &&
                !this.hasEmptyRow()
            ) {
                this.addEmptyRow();
            }

        });

        // Render formatted columns (date, textarea, readonly, email) - MUST run first before other renders
        this._renderFormattedColumns();

        // Render checkbox columns on initial load
        this._renderCheckboxColumns();

        // Render select columns with arrow indicator
        this._renderSelectColumns();

        // Re-render all formatted columns when table redraws (paging, sorting, etc.)
        this.table.on('draw.dt', () => {
            this._renderFormattedColumns();
            this._renderCheckboxColumns();
            this._renderSelectColumns();
            this._renderNumericColumns();
        });

        // Initialize custom tooltip for long text cells
        this._initTooltip();
    }

    // ==================== TOOLTIP METHODS ====================

    /**
     * Initialize custom tooltip element and event handlers
     * Called once during GridSheet initialization
     */
    _initTooltip() {
        // Create shared tooltip element
        let tooltip = document.querySelector('.dt-tooltip');
        if (!tooltip) {
            tooltip = document.createElement('div');
            tooltip.className = 'dt-tooltip';
            document.body.appendChild(tooltip);
        }
        this._tooltip = tooltip;

        // Add hover events for textarea cells and long text
        const tooltipTableNode = this.table.table().node();
        tooltipTableNode.addEventListener('mouseenter', (e) => {
            const td = e.target.closest('td');
            if (!td || this.isEditMode) return;

            // Check cell type
            const dataType = td.getAttribute('data-type');
            const isFormulaCell = td.classList.contains('dt-formula-cell');

            // For formula cells, show the formula from data-formula-display attribute
            if (isFormulaCell) {
                const formulaText = td.getAttribute('data-formula-display');
                if (formulaText) {
                    this._showTooltip(td, formulaText);
                }
                return;
            }

            // Get FULL text - prefer full-text over raw-value over truncated textContent
            let fullText = td.getAttribute('data-full-text') ||
                td.getAttribute('data-raw-value') ||
                td.textContent.trim();

            // Remove default browser tooltip (but not for formula cells)
            td.removeAttribute('title');

            // Show tooltip for textarea type OR text longer than 50 chars
            // BUT only if text is not empty
            if (fullText.length > 0 && (dataType === 'textarea' || fullText.length > 50)) {
                this._showTooltip(td, fullText);
            }
        }, true);

        tooltipTableNode.addEventListener('mouseleave', (e) => {
            const td = e.target.closest('td');
            if (td) {
                this._hideTooltip();
            }
        }, true);

        // Hide tooltip on scroll (any scroll in page)
        window.addEventListener('scroll', () => {
            this._hideTooltip();
        }, true);

        // Hide tooltip on resize
        window.addEventListener('resize', () => {
            this._hideTooltip();
        });
    }

    /**
     * Show tooltip with text at cell position
     * @param {HTMLElement} td - Cell element
     * @param {string} text - Text to display
     */
    _showTooltip(td, text) {
        this._tooltip.textContent = text;

        // Position tooltip below the cell and match cell width
        const rect = td.getBoundingClientRect();
        this._tooltip.style.left = rect.left + 'px';
        this._tooltip.style.top = (rect.bottom + 8) + 'px';
        // Set tooltip width to match cell width for text wrapping
        this._tooltip.style.width = rect.width + 'px';
        this._tooltip.style.maxWidth = Math.max(rect.width, 200) + 'px';

        // Adjust if tooltip goes off-screen
        setTimeout(() => {
            const tooltipRect = this._tooltip.getBoundingClientRect();
            if (tooltipRect.right > window.innerWidth) {
                this._tooltip.style.left = (window.innerWidth - tooltipRect.width - 10) + 'px';
            }
            if (tooltipRect.bottom > window.innerHeight) {
                this._tooltip.style.top = (rect.top - tooltipRect.height - 8) + 'px';
            }
        }, 10);

        this._tooltip.classList.add('visible');
    }

    /**
     * Hide tooltip
     */
    _hideTooltip() {
        if (this._tooltip) {
            this._tooltip.classList.remove('visible');
        }
    }

    // ==================== END TOOLTIP METHODS ====================

    /**
     * Render checkbox inputs for columns with data-type="checkbox"
     * Converts true/false text to actual checkbox elements
     */
    _renderCheckboxColumns() {
        const tableNode = this.table.table().node();
        const headers = tableNode.querySelectorAll('thead th');

        // Find checkbox columns
        const checkboxColIndices = [];
        headers.forEach((th, idx) => {
            if (th.getAttribute('data-type') === 'checkbox') {
                checkboxColIndices.push(idx);
            }
        });

        if (checkboxColIndices.length === 0) return;

        // Render checkboxes for each row
        const rows = tableNode.querySelectorAll('tbody tr');
        rows.forEach(row => {
            checkboxColIndices.forEach(colIdx => {
                const td = row.cells[colIdx];
                if (!td) return;

                // Skip if already has checkbox
                if (td.querySelector('input[type="checkbox"]')) return;

                const value = td.textContent.trim().toLowerCase();
                const isChecked = value === 'true' || value === '1';

                // Check if row is readonly
                const isRowReadonly = row.getAttribute('data-readonly-row') === 'true';

                // Check if column is readonly
                const colHeader = headers[colIdx];
                const isColReadonly = colHeader && colHeader.getAttribute('data-readonly') === 'true';

                // Create checkbox
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.className = 'dt-checkbox-display';
                checkbox.checked = isChecked;

                // Disable checkbox if row or column is readonly
                if (isRowReadonly || isColReadonly) {
                    checkbox.disabled = true;
                }

                // Add click handler for direct toggle
                checkbox.addEventListener('change', (e) => {
                    e.stopPropagation();

                    // Check if row is readonly
                    const rowNode = td.closest('tr');
                    if (rowNode && rowNode.getAttribute('data-readonly-row') === 'true') {
                        console.log('Cannot change checkbox - row is readonly');
                        // Revert checkbox to previous state
                        checkbox.checked = !checkbox.checked;
                        return;
                    }

                    // Check if column is readonly
                    const colIdxForCheck = td.cellIndex;
                    const thForCheck = this.table.column(colIdxForCheck).header();
                    if (thForCheck && thForCheck.getAttribute('data-readonly') === 'true') {
                        console.log('Cannot change checkbox - column is readonly');
                        checkbox.checked = !checkbox.checked;
                        return;
                    }

                    const newValue = checkbox.checked ? 'true' : 'false';
                    td.setAttribute('data-checkbox-value', newValue);

                    // Get row and field info for saving
                    const rowId = rowNode ? rowNode.getAttribute('data-id') : null;
                    const colIdx = td.cellIndex;
                    const th = this.table.column(colIdx).header();
                    const fieldName = th ? th.getAttribute('data-name') : null;

                    // Check if this is a new row (need to preserve attributes after draw)
                    const isNewRow = rowNode && rowNode.getAttribute('data-new-row') === 'true';


                    // Set flag to prevent addEmptyRow during checkbox change
                    this.isCheckboxChanging = true;

                    // Simpan index cell sebelum draw untuk restore focus nanti
                    const cellIdxObj = this.table.cell(td).index();

                    // Update cell data triggers redraw!
                    const cell = this.table.cell(td);
                    if (cell) {
                        cell.data(newValue).draw(false);
                    }

                    // Restore data-new-row attribute if this was a new row
                    if (isNewRow) {
                        setTimeout(() => {
                            const newRowNode = this.table.row(cellIdxObj.row).node();
                            if (newRowNode) {
                                newRowNode.setAttribute('data-new-row', 'true');
                                if (rowId === 'new') {
                                    newRowNode.setAttribute('data-id', 'new');
                                }
                            }
                            // Reset flag after attributes are restored
                            this.isCheckboxChanging = false;
                        }, 20);
                    } else {
                        // Reset flag for non-new rows
                        setTimeout(() => {
                            this.isCheckboxChanging = false;
                        }, 20);
                    }

                    // Save based on mode
                    if (rowId && fieldName) {
                        if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
                            // BATCH MODE
                            const hasOtherContent = this._rowHasContent(rowNode);

                            if (rowId === 'new') {
                                // New row - generate temp_id first
                                if (hasOtherContent) {
                                    const tempId = 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
                                    rowNode.setAttribute('data-id', tempId);
                                    rowNode.setAttribute('data-pending', 'true');
                                    // Save per-cell to localStorage
                                    this._updateLocalStorageEntry(tempId, fieldName, newValue);
                                }
                            } else if (rowId.startsWith('temp_')) {
                                // Existing temp row
                                rowNode.setAttribute('data-pending', 'true');
                                // Save per-cell to localStorage
                                this._updateLocalStorageEntry(rowId, fieldName, newValue);
                            } else {
                                // Existing row from DB
                                if (hasOtherContent) {
                                    rowNode.setAttribute('data-edited', 'true');
                                    this._trackEditedRow(rowId, fieldName, newValue);
                                }
                            }
                            this._updateBatchCount();
                        } else {
                            // DIRECT MODE - Save immediately on toggle
                            if (rowId !== 'new') {
                                this.updateCell(fieldName, rowId, newValue);
                            }
                        }
                    }

                    // Re-focus the cell after save/redraw
                    // Wajib pakai timeout karena draw() async render di DOM
                    setTimeout(() => {
                        if (cellIdxObj) {
                            const newCell = this.table.cell(cellIdxObj.row, cellIdxObj.column);
                            const newTd = newCell.node();

                            if (newTd) {
                                // Restore visual focus
                                newTd.classList.add('dt-cell-selected');
                                newTd.classList.add('focus'); // KeyTable class

                                // Restore internal state
                                this.currentCell = newCell;
                                this.activeCellNode = newTd;
                            }
                        }
                    }, 50);

                    // Re-render to restore checkbox
                    setTimeout(() => this._renderCheckboxColumns(), 10);
                });

                // Clear cell and add checkbox
                td.textContent = '';
                td.style.textAlign = 'center';
                td.appendChild(checkbox);

                // Store value as data attribute
                td.setAttribute('data-checkbox-value', isChecked ? 'true' : 'false');
            });
        });
    }

    /**
     * Render select columns with dropdown arrow indicator
     * Adds visual indicator that cell is a dropdown
     */
    _renderSelectColumns() {
        const tableNode = this.table.table().node();
        const headers = tableNode.querySelectorAll('thead th');

        // Find select columns
        const selectColIndices = [];
        headers.forEach((th, idx) => {
            if (th.getAttribute('data-type') === 'select') {
                selectColIndices.push(idx);
            }
        });

        if (selectColIndices.length === 0) return;

        // Add arrow indicator to each select cell
        const rows = tableNode.querySelectorAll('tbody tr');
        rows.forEach(row => {
            selectColIndices.forEach(colIdx => {
                const td = row.querySelectorAll('td')[colIdx];
                if (!td) return;

                // Skip if already has arrow or is being edited
                if (td.querySelector('.dt-select-arrow')) return;
                if (td.classList.contains('dt-cell-editing')) return;

                // Add arrow wrapper
                const content = td.textContent.trim();
                td.innerHTML = '';
                td.style.position = 'relative';
                td.style.paddingRight = '24px';

                // Content span
                const contentSpan = document.createElement('span');
                contentSpan.className = 'dt-select-content';
                contentSpan.textContent = content;
                contentSpan.title = content; // Show full text on hover
                td.appendChild(contentSpan);

                // Arrow indicator
                const arrow = document.createElement('span');
                arrow.className = 'dt-select-arrow';
                arrow.innerHTML = '▼';
                td.appendChild(arrow);

                // Mark td as select type for click handling
                td.setAttribute('data-select-cell', 'true');
            });
        });
    }

    /**
     * Render formatted columns (date, textarea, readonly, email)
     * Applies proper formatting on initial load
     */
    _renderFormattedColumns() {
        const tableNode = this.table.table().node();
        const headers = tableNode.querySelectorAll('thead th');

        // Find columns that need formatting
        const formattedCols = [];
        headers.forEach((th, idx) => {
            const dataType = th.getAttribute('data-type');
            if (['date', 'textarea', 'readonly', 'email'].includes(dataType)) {
                formattedCols.push({ idx, dataType });
            }
        });

        if (formattedCols.length === 0) return;

        // Render for each row
        const rows = tableNode.querySelectorAll('tbody tr');
        rows.forEach(row => {
            formattedCols.forEach(({ idx, dataType }) => {
                const td = row.cells[idx];
                if (!td) return;

                // Skip if already formatted
                if (td.getAttribute('data-formatted')) return;

                const rawValue = td.textContent.trim();
                td.setAttribute('data-raw-value', rawValue);
                td.setAttribute('data-type', dataType);
                td.setAttribute('data-formatted', 'true');

                switch (dataType) {
                    case 'date':
                        // Format date for display
                        const isoDate = this._parseDateToISO(rawValue);
                        const formattedDate = this._formatDateForDisplay(isoDate);
                        // Store formatted date in data-raw-value for CSS pseudo-element display
                        td.setAttribute('data-raw-value', formattedDate);

                        // Update DataTables internal data
                        try {
                            const cell = this.table.cell(td);
                            if (cell && cell.node()) {
                                cell.data(formattedDate);
                            }
                        } catch (e) {
                            // Ignore errors during initial render
                        }
                        break;

                    case 'textarea':
                        // Store full text for tooltip (before truncation)
                        td.setAttribute('data-full-text', rawValue);

                        // Truncate with ellipsis for display
                        const displayLength = (this.formats.textarea && this.formats.textarea.displayLength) || 50;
                        if (rawValue.length > displayLength) {
                            td.textContent = rawValue.substring(0, displayLength) + '...';
                            // Do NOT set title - we use custom tooltip instead
                        }
                        td.classList.add('dt-textarea-cell');
                        break;

                    case 'readonly':
                        // Add readonly styling
                        td.classList.add('dt-readonly-cell');
                        break;

                    case 'email':
                        // Just mark as email
                        break;
                }
            });
        });
    }

    /**
     * Render numeric column formatting (currency, number, percentage, time)
     * Called on page change to apply number formatting and alignment
     */
    _renderNumericColumns() {
        const tableNode = this.table.table().node();
        const headers = tableNode.querySelectorAll('thead th');

        // Find numeric and time columns that need formatting
        const numericTimeColumns = [];
        headers.forEach((th, idx) => {
            const dataType = th.getAttribute('data-type');
            if (['currency', 'number', 'percentage', 'time'].includes(dataType)) {
                numericTimeColumns.push({ idx, dataType });
            }
        });

        if (numericTimeColumns.length === 0) return;

        // Render for each row
        const rows = tableNode.querySelectorAll('tbody tr');
        rows.forEach(row => {
            numericTimeColumns.forEach(({ idx, dataType }) => {
                const td = row.cells[idx];
                if (!td) return;

                // Skip if already formatted in this render
                if (td.getAttribute('data-numeric-formatted')) return;

                td.setAttribute('data-numeric-formatted', 'true');

                if (dataType === 'currency' || dataType === 'number' || dataType === 'percentage') {
                    // Right-align numeric types
                    td.style.textAlign = 'right';

                    // Get row index for cache key
                    const dtCell = this.table.cell(td);
                    if (!dtCell || !dtCell.index()) return;

                    const rowIndex = dtCell.index().row;
                    const cacheKey = `${rowIndex}_${idx}`;

                    // PRIORITY: Check cache first (set by _applyColumnTypes)
                    const cachedValue = this._rawValueCache.get(cacheKey);
                    if (cachedValue !== undefined && cachedValue !== null && cachedValue !== '') {
                        // Use cached raw value - already processed correctly
                        td.setAttribute('data-raw-value', cachedValue);
                        td.textContent = this._formatValue(cachedValue, dataType);
                        return;
                    }

                    // Check existing data-raw-value attribute
                    const existingRawValue = td.getAttribute('data-raw-value');
                    if (existingRawValue !== null && existingRawValue !== '') {
                        // Use existing raw value and cache it
                        this._rawValueCache.set(cacheKey, existingRawValue);
                        td.textContent = this._formatValue(existingRawValue, dataType);
                        return;
                    }

                    // Fallback: Parse from textContent
                    const rawValue = td.textContent.trim();
                    if (rawValue) {
                        // Use _cleanValueForType which handles all edge cases correctly
                        const cleanedValue = this._cleanValueForType(rawValue, dataType);

                        if (cleanedValue !== '' && !isNaN(parseFloat(cleanedValue))) {
                            // Store in cache and DOM
                            this._rawValueCache.set(cacheKey, cleanedValue);
                            td.setAttribute('data-raw-value', cleanedValue);
                            td.textContent = this._formatValue(cleanedValue, dataType);
                        }
                    }
                } else if (dataType === 'time') {
                    // Right-align time columns
                    td.style.textAlign = 'right';
                }
            });
        });

        // Recalculate formula columns after numeric rendering
        this._recalculateAllFormulas();
    }

    // ============================================
    // FOOTER AGGREGATE METHODS
    // ============================================

    /**
     * Initialize footer row with aggregate cells
     * Called once during GridSheet initialization if footer.enabled is true
     */
    _initFooter() {
        if (!this.footer.enabled) return;

        const tableNode = this.table.table().node();
        let tfoot = tableNode.querySelector('tfoot');

        // Create tfoot if it doesn't exist
        if (!tfoot) {
            tfoot = document.createElement('tfoot');
            tableNode.appendChild(tfoot);
        }

        // Clear existing footer rows
        tfoot.innerHTML = '';

        // Create footer row
        const footerRow = document.createElement('tr');
        footerRow.classList.add('dt-footer-totals');

        const headers = tableNode.querySelectorAll('thead th');
        const colCount = headers.length;

        headers.forEach((th, index) => {
            const td = document.createElement('td');
            const footerType = th.getAttribute('data-footer');
            const dataType = th.getAttribute('data-type');

            // First column: show label
            if (index === 0) {
                td.textContent = this.footer.label || 'Total';
                td.style.fontWeight = 'bold';
            } else if (footerType) {
                // Column with aggregate function
                td.setAttribute('data-footer-type', footerType);
                td.setAttribute('data-footer-col', index);

                // Copy data-type for formatting
                if (dataType) {
                    td.setAttribute('data-type', dataType);
                }

                // For formula columns, get format from data-format
                const dataFormat = th.getAttribute('data-format');
                if (dataFormat) {
                    td.setAttribute('data-format', dataFormat);
                }

                td.classList.add('dt-footer-aggregate');
                td.textContent = '-'; // Placeholder
            }

            // Apply right-align for numeric types
            if (['number', 'currency', 'percentage', 'formula'].includes(dataType)) {
                td.style.textAlign = 'right';
            }

            footerRow.appendChild(td);
        });

        tfoot.appendChild(footerRow);
        this._footerRow = footerRow;

        // Calculate initial values
        this._updateFooterTotals();
    }

    /**
     * Get columns configured for footer aggregate
     * @returns {Array} Array of {index, type, format, element}
     */
    _getFooterColumns() {
        const tableNode = this.table.table().node();
        const headers = tableNode.querySelectorAll('thead th');
        const columns = [];

        headers.forEach((th, index) => {
            const footerType = th.getAttribute('data-footer');
            if (footerType) {
                columns.push({
                    index,
                    type: footerType.toLowerCase(), // sum, avg, count
                    dataType: th.getAttribute('data-type'),
                    format: th.getAttribute('data-format') || th.getAttribute('data-type'),
                    element: th
                });
            }
        });

        return columns;
    }

    /**
     * Calculate aggregate value for a column
     * @param {number} colIndex - Column index
     * @param {string} aggregateType - 'sum', 'avg', or 'count'
     * @returns {number|string} Calculated value
     */
    _calculateColumnAggregate(colIndex, aggregateType) {
        const values = [];

        // Use DataTables API to get all rows (including all pages)
        const tableNode = this.table.table().node();
        const rows = tableNode.querySelectorAll('tbody tr');

        rows.forEach(row => {
            // Skip rows without any cells
            const cells = row.querySelectorAll('td');
            if (cells.length === 0) return;

            const cell = cells[colIndex];
            if (!cell) return;

            // Get raw value (prefer data-raw-value over textContent)
            let rawValue = cell.getAttribute('data-raw-value');

            if (rawValue === null || rawValue === undefined || rawValue === '') {
                // Try to parse textContent
                let text = cell.textContent.trim();
                if (text === '' || text === '-') return; // Skip empty cells

                // Remove formatting (thousand separators, currency symbols, etc.)
                text = text.replace(/[^\d.,\-]/g, '');
                text = text.replace(/\./g, ''); // Remove thousand separators
                text = text.replace(',', '.'); // Convert decimal separator
                rawValue = text;
            }

            const numValue = parseFloat(rawValue);
            if (!isNaN(numValue)) {
                values.push(numValue);
            } else if (aggregateType === 'count' && cell.textContent.trim() !== '') {
                // For COUNT, include non-empty text values
                values.push(1);
            }
        });


        if (values.length === 0) return 0;

        switch (aggregateType) {
            case 'sum':
                return values.reduce((a, b) => a + b, 0);
            case 'avg':
                // Round to 2 decimal places to avoid long decimals
                const avg = values.reduce((a, b) => a + b, 0) / values.length;
                return Math.round(avg * 100) / 100;
            case 'count':
                return values.length;
            default:
                return 0;
        }
    }

    /**
     * Update all footer aggregate cells with recalculated values
     * Should be called after any data change (edit, paste, fill, delete)
     */
    _updateFooterTotals() {

        if (!this.footer.enabled || !this._footerRow) {
            return;
        }

        const columns = this._getFooterColumns();

        columns.forEach(col => {
            const footerCell = this._footerRow.cells[col.index];
            if (!footerCell) {
                return;
            }

            const value = this._calculateColumnAggregate(col.index, col.type);

            // Store old value for comparison
            const oldText = footerCell.textContent;

            // Format the value based on column type
            if (col.type === 'count') {
                footerCell.textContent = value;
            } else {
                // Get format config
                const formatType = col.format || footerCell.getAttribute('data-format');
                const format = this.formats[formatType];

                if (format) {
                    // For AVG, always show 2 decimal places for accuracy
                    const isAvg = col.type === 'avg';
                    footerCell.textContent = this._formatNumberDirect(value, format, isAvg);
                } else {
                    // No format config, just display the value
                    footerCell.textContent = col.type === 'avg' ? value.toFixed(2) : value;
                }
            }

            // Store raw value
            footerCell.setAttribute('data-raw-value', value);

        });
    }

    /**
     * Format a number directly without parsing (for footer aggregates)
     * @param {number} num - Numeric value
     * @param {Object} format - Format config object
     * @param {boolean} isAvg - If true, force 2 decimal places for average
     * @returns {string} Formatted string
     */
    _formatNumberDirect(num, format, isAvg = false) {
        if (num === null || num === undefined || isNaN(num)) return '';

        // Round according to decimal places
        // For AVG, always use 2 decimal places for accuracy
        let decimalPlaces;
        if (isAvg) {
            decimalPlaces = 2;
        } else if (format.decimalPlaces !== null && format.decimalPlaces !== undefined) {
            decimalPlaces = format.decimalPlaces;
        } else {
            decimalPlaces = 0;
        }

        const formatted = num.toFixed(decimalPlaces);

        // Split integer and decimal parts
        const parts = formatted.split('.');
        let integerPart = parts[0];
        const decimalPart = parts[1] || '';

        // Apply thousand separator
        if (format.thousandSeparator) {
            integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, format.thousandSeparator);
        }

        // Combine with decimal separator
        let result = integerPart;
        if (decimalPart) {
            result += format.decimalSeparator + decimalPart;
        }

        // Add prefix and suffix
        return (format.prefix || '') + result + (format.suffix || '');
    }

    // ============================================
    // FORMULA COLUMN METHODS
    // ============================================

    /**
     * Get all column values from a row as key-value object
     * @param {HTMLElement} rowNode - TR element
     * @returns {Object} Object with fieldName as key, value as value
     */
    _getRowFieldValues(rowNode) {
        const values = {};
        const tableNode = this.table.table().node();
        const headers = tableNode.querySelectorAll('thead th');

        headers.forEach((th, index) => {
            const fieldName = th.getAttribute('data-name');
            const dataType = th.getAttribute('data-type');

            // Skip formula columns - they don't have data-name and shouldn't be referenced
            if (fieldName && dataType !== 'formula') {
                const cell = rowNode.cells[index];
                if (!cell) return;

                // Get raw value (prefer data-raw-value for numeric types)
                const rawValue = cell.getAttribute('data-raw-value') || cell.textContent || '';

                // Parse to numeric if data type is numeric
                if (['number', 'currency', 'percentage'].includes(dataType)) {
                    const cleaned = rawValue.replace(/[^\d.-]/g, '');
                    values[fieldName] = parseFloat(cleaned) || 0;
                } else {
                    values[fieldName] = rawValue.trim();
                }
            }
        });

        return values;
    }

    /**
     * Get list of formula columns with their configuration
     * @returns {Array} Array of {index, formula, format, element}
     */
    _getFormulaColumns() {
        const tableNode = this.table.table().node();
        const headers = tableNode.querySelectorAll('thead th[data-type="formula"]');

        return Array.from(headers).map(th => ({
            index: Array.from(th.parentNode.children).indexOf(th),
            name: th.getAttribute('data-name'),  // Column name for formula chaining
            formula: th.getAttribute('data-formula'),
            format: th.getAttribute('data-format'),
            element: th
        }));
    }

    /**
     * Safe expression evaluation without arbitrary code execution
     * @param {string} expression - Mathematical expression to evaluate
     * @returns {*} Result of evaluation
     */
    _safeEval(expression) {
        // Only allow: numbers, operators, parentheses, Math functions, strings, quotes
        // This is a simple approach - for production may want more robust sandboxing
        try {
            return Function('"use strict"; return (' + expression + ')')();
        } catch (e) {
            console.error('SafeEval error:', expression, e);
            return NaN;
        }
    }

    /**
     * Parse and handle formula functions (SUM, ROUND, CONCAT)
     * @param {string} expression - Formula expression
     * @param {Object} rowData - Row data object {fieldName: value}
     * @returns {string} Expression with functions converted to JavaScript
     */
    _parseFormulaFunctions(expression, rowData) {
        // SUM(a, b, c) → (a + b + c)
        expression = expression.replace(/SUM\s*\(([^)]+)\)/gi, (match, args) => {
            const values = args.split(',').map(arg => {
                const field = arg.trim();
                // If it's a string literal, keep it
                if (/^['"].*['"]$/.test(field)) return field;
                // If it's a field reference, get value
                if (rowData[field] !== undefined) return rowData[field];
                // Otherwise keep as is (could be a number)
                return field;
            });
            return `(${values.join(' + ')})`;
        });

        // ROUND(value, decimals) → Math.round((value) * 10^decimals) / 10^decimals
        expression = expression.replace(/ROUND\s*\(([^,]+),\s*(\d+)\)/gi, (match, value, decimals) => {
            const field = value.trim();
            let num = field;
            if (rowData[field] !== undefined) {
                num = rowData[field];
            }
            const d = parseInt(decimals);
            const multiplier = Math.pow(10, d);
            return `(Math.round((${num}) * ${multiplier}) / ${multiplier})`;
        });

        // CONCAT(a, b, c) → (a + b + c) for strings
        expression = expression.replace(/CONCAT\s*\(([^)]+)\)/gi, (match, args) => {
            const parts = args.split(',').map(arg => {
                const field = arg.trim();
                // If it's a string literal, keep it
                if (/^['"].*['"]$/.test(field)) return field;
                // If it's a field reference, wrap in quotes
                if (rowData[field] !== undefined) {
                    const val = rowData[field];
                    return typeof val === 'string' ? `"${val}"` : val;
                }
                // Otherwise assume it's a literal string and wrap in quotes
                return `"${field}"`;
            });
            return `(${parts.join(' + ')})`;
        });

        return expression;
    }

    /**
     * Parse and evaluate formula expression with row data
     * @param {string} formula - Formula expression (e.g., "qty * price")
     * @param {Object} rowData - Row data object {fieldName: value}
     * @returns {*} Result or error code (#DIV/0!, #REF!, #VALUE!, #ERROR)
     */
    _evaluateFormula(formula, rowData) {
        try {
            let expression = formula;

            // Step 1: Parse function calls (SUM, ROUND, CONCAT)
            expression = this._parseFormulaFunctions(expression, rowData);

            // Step 2: Replace column references with values
            // Sort by length (longest first) to avoid partial replacements
            const fieldNames = Object.keys(rowData).sort((a, b) => b.length - a.length);

            fieldNames.forEach(field => {
                const value = rowData[field];
                const regex = new RegExp(`\\b${field}\\b`, 'g');

                if (typeof value === 'string') {
                    // Wrap strings in quotes
                    expression = expression.replace(regex, `"${value}"`);
                } else {
                    expression = expression.replace(regex, value);
                }
            });

            // Step 3: Check for unresolved references (field names that weren't replaced)
            // Ignore Math.xxx and string literals
            const cleanedExpr = expression
                .replace(/Math\.\w+/g, '')
                .replace(/"[^"]*"/g, '')
                .replace(/'[^']*'/g, '');

            if (/\b[a-zA-Z_][a-zA-Z0-9_]*\b/.test(cleanedExpr)) {
                return '#REF!';
            }

            // Step 4: Check for division by zero
            if (/\/\s*0(?!\d)/.test(expression)) {
                return '#DIV/0!';
            }

            // Step 5: Evaluate the expression
            const result = this._safeEval(expression);

            // Step 6: Check for NaN result (invalid math operation)
            if (typeof result === 'number' && isNaN(result)) {
                return '#VALUE!';
            }

            return result;

        } catch (e) {
            console.error('Formula evaluation error:', formula, e);
            return '#ERROR';
        }
    }
    /**
     * Recalculate all formula cells in a single row
     * @param {HTMLElement} rowNode - TR element
     */
    _recalculateRowFormulas(rowNode) {
        const formulaColumns = this._getFormulaColumns();
        if (formulaColumns.length === 0) return;

        const rowData = this._getRowFieldValues(rowNode);

        // Process formulas in order - each formula result is added to rowData
        // so subsequent formulas can reference it (formula chaining)
        formulaColumns.forEach(col => {
            const cell = rowNode.cells[col.index];
            if (!cell) return;

            // Check if all source fields referenced in formula are empty
            // Extract field names from formula (word boundaries)
            const fieldRefs = col.formula.match(/\b[a-zA-Z_][a-zA-Z0-9_]*\b/g) || [];
            // Filter to only actual column fields (not functions like SUM, ROUND, CONCAT)
            const functions = ['SUM', 'ROUND', 'CONCAT', 'Math'];
            const sourceFields = fieldRefs.filter(f =>
                !functions.includes(f) && rowData.hasOwnProperty(f)
            );

            // Check if ALL source fields are empty/null/undefined
            const allEmpty = sourceFields.length > 0 && sourceFields.every(field => {
                const val = rowData[field];
                return val === null || val === undefined || val === '' || val === 0;
            });

            if (allEmpty) {
                // Row has no data - show empty cell
                cell.textContent = '';
                cell.removeAttribute('data-raw-value');
                cell.classList.remove('dt-formula-error');
                // Add zero to rowData so it can be referenced by other formulas
                if (col.name) {
                    rowData[col.name] = 0;
                }
            } else {
                const result = this._evaluateFormula(col.formula, rowData);
                const isError = typeof result === 'string' && result.startsWith('#');

                if (isError) {
                    // Display error code
                    cell.textContent = result;
                    cell.classList.add('dt-formula-error');
                    cell.removeAttribute('data-raw-value');
                    // Add 0 for error so other formulas don't break
                    if (col.name) {
                        rowData[col.name] = 0;
                    }
                } else {
                    cell.classList.remove('dt-formula-error');

                    if (typeof result === 'number' && col.format) {
                        // Format numeric result
                        cell.textContent = this._formatValue(result, col.format);
                        cell.setAttribute('data-raw-value', result);
                    } else {
                        // Display as-is (string or unformatted number)
                        cell.textContent = result;
                        if (typeof result === 'number') {
                            cell.setAttribute('data-raw-value', result);
                        }
                    }

                    // Add result to rowData so subsequent formulas can reference it
                    if (col.name) {
                        rowData[col.name] = typeof result === 'number' ? result : 0;
                    }
                }
            }

            // Apply readonly styling (formula cells are not editable)
            cell.classList.add('dt-readonly-cell', 'dt-formula-cell');

            // Build formula display text using header text (not field names)
            // e.g., "Qty * Price" instead of "qty * price"
            let displayFormula = col.formula;
            const tableNode = this.table.table().node();
            const headers = tableNode.querySelectorAll('thead th');
            headers.forEach(th => {
                const fieldName = th.getAttribute('data-name');
                if (fieldName) {
                    const headerText = th.textContent.trim();
                    // Replace field names with header text (case insensitive, word boundary)
                    const regex = new RegExp(`\\b${fieldName}\\b`, 'gi');
                    displayFormula = displayFormula.replace(regex, headerText);
                }
            });

            // Store formula for custom tooltip (use data attribute, not title)
            cell.setAttribute('data-formula-display', `= ${displayFormula}`);
            cell.removeAttribute('title'); // Prevent native browser tooltip
        });
    }

    /**
     * Recalculate formulas for all visible rows
     */
    _recalculateAllFormulas() {
        const tableNode = this.table.table().node();
        const rows = tableNode.querySelectorAll('tbody tr');
        rows.forEach(row => this._recalculateRowFormulas(row));

        // Update footer totals after formula recalculation
        this._updateFooterTotals();
    }

    /**
     * Initialize multi-row selection events for No. column
     */
    _initRowSelectionEvents() {
        const tableNode = this.table.table().node();

        // Mousedown on No. column starts row selection (left-click only)
        this._addEvent(tableNode, 'mousedown', (e) => {
            // Skip right-click to preserve selection for context menu
            if (e.button === 2) return;

            const td = e.target.closest('td');
            if (!td) return;

            const row = td.closest('tr');
            if (!row || !row.parentElement || row.parentElement.tagName !== 'TBODY') return;

            // Only handle No. column (first column)
            if (this.enableColumnNumber && td.cellIndex === 0) {
                e.preventDefault();

                const rowIdx = this.table.row(row).index();

                // Shift+Click: Extend selection from last selected row
                if (e.shiftKey && this._selectedRows.length > 0 && this._rowSelectionStart !== null) {
                    // Calculate range from selection start to clicked row
                    const startRow = Math.min(this._rowSelectionStart, rowIdx);
                    const endRow = Math.max(this._rowSelectionStart, rowIdx);

                    // Clear previous highlights
                    this._clearRowSelection();

                    // Set new range selection
                    this._selectedRows = [];
                    for (let i = startRow; i <= endRow; i++) {
                        this._selectedRows.push(i);
                    }

                    this._highlightSelectedRows();

                    // Don't start drag for shift+click
                    return;
                }

                // Normal click: Start new selection
                this._isRowSelecting = true;
                this._rowSelectionStart = rowIdx;

                // ===== Clear previous row selection first =====
                this._clearRowSelection();

                // Start new selection
                this._selectedRows = [rowIdx];

                // ===== THOROUGHLY Clear all cell selections =====
                this._clearCellSelection();
                this._removeFillHandle();

                // Additionally clear focus/cell selection classes that might be missed
                const tableNode = this.table.table().node();
                tableNode.querySelectorAll('td.focus, td.cell-focus, td.cell-selected, td.dt-cell-selected').forEach(c => {
                    c.classList.remove('focus', 'cell-focus', 'cell-selected', 'dt-cell-selected');
                });

                // Blur KeyTable internal focus
                if (this.table.cell.blur) {
                    this.table.cell.blur();
                }

                // Reset cell selection state
                this.currentCell = null;
                this._selectionStart = null;

                // Highlight row
                this._highlightSelectedRows();

                // Add move/end listeners (these are temporary and will be removed in _onRowSelectEnd)
                document.addEventListener('mousemove', this._onRowSelectMove);
                document.addEventListener('mouseup', this._onRowSelectEnd);
            }
        });

        // Keyboard shortcuts for multi-row operations
        this._addEvent(document, 'keydown', (e) => {
            // Ctrl+C - Copy selected rows
            if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
                if (this._selectedRows.length > 0) {
                    e.preventDefault();
                    this._copyMultipleRows();
                }
            }

            // Ctrl+V - Paste (REMOVED: This was duplicate! Main handler is in keyboard shortcuts section above)
            // Duplicate handler was causing paste to run twice

            // Shift+Arrow - Extend row selection
            if (e.shiftKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
                if (this._selectedRows.length > 0) {
                    e.preventDefault();
                    e.stopImmediatePropagation(); // Prevent KeyTable from handling

                    // Clear any cell focus/selection
                    this._clearCellSelection();
                    const tableNode = this.table.table().node();
                    tableNode.querySelectorAll('td.focus, td.dt-focused, td.cell-focus').forEach(c => {
                        c.classList.remove('focus', 'dt-focused', 'cell-focus');
                    });
                    if (this.table.cell.blur) {
                        this.table.cell.blur();
                    }
                    this.currentCell = null;

                    // Get current selection bounds
                    const minRow = Math.min(...this._selectedRows);
                    const maxRow = Math.max(...this._selectedRows);
                    const rowCount = this.table.rows().count();

                    // Initialize selection start if not set
                    if (this._rowSelectionStart === null || this._rowSelectionStart === undefined) {
                        this._rowSelectionStart = minRow;
                    }

                    let newRow;
                    if (e.key === 'ArrowUp') {
                        // Extend upward
                        newRow = minRow - 1;
                        if (newRow >= 0) {
                            // Add row to selection
                            const startRow = Math.min(this._rowSelectionStart, newRow);
                            const endRow = Math.max(this._rowSelectionStart, maxRow);
                            this._selectedRows = [];
                            for (let i = startRow; i <= endRow; i++) {
                                this._selectedRows.push(i);
                            }
                        }
                    } else if (e.key === 'ArrowDown') {
                        // Extend downward
                        newRow = maxRow + 1;
                        if (newRow < rowCount) {
                            // Add row to selection
                            const startRow = Math.min(this._rowSelectionStart, minRow);
                            const endRow = Math.max(this._rowSelectionStart, newRow);
                            this._selectedRows = [];
                            for (let i = startRow; i <= endRow; i++) {
                                this._selectedRows.push(i);
                            }
                        }
                    }

                    this._highlightSelectedRows();
                }
            }
        });
    }

    /**
     * Handle mouse move during row selection
     */
    _handleRowSelectMove(e) {
        if (!this._isRowSelecting) return;

        const tableNode = this.table.table().node();
        const td = document.elementFromPoint(e.clientX, e.clientY)?.closest('td');
        if (!td) return;

        const row = td.closest('tr');
        if (!row || !tableNode.contains(row)) return;

        const rowIdx = this.table.row(row).index();
        if (rowIdx === undefined || rowIdx < 0) return;

        // Calculate range
        const startRow = Math.min(this._rowSelectionStart, rowIdx);
        const endRow = Math.max(this._rowSelectionStart, rowIdx);

        // Update selected rows
        this._selectedRows = [];
        for (let i = startRow; i <= endRow; i++) {
            this._selectedRows.push(i);
        }

        this._highlightSelectedRows();
    }

    /**
     * Handle mouse up to end row selection
     */
    _handleRowSelectEnd(e) {
        this._isRowSelecting = false;
        document.removeEventListener('mousemove', this._onRowSelectMove);
        document.removeEventListener('mouseup', this._onRowSelectEnd);
    }

    /**
     * Highlight selected rows
     */
    _highlightSelectedRows() {
        const tableNode = this.table.table().node();

        // Clear existing row highlights
        tableNode.querySelectorAll('tr.dt-row-selected').forEach(tr => {
            tr.classList.remove('dt-row-selected');
        });

        // Highlight selected rows
        this._selectedRows.forEach(rowIdx => {
            const rowNode = this.table.row(rowIdx).node();
            if (rowNode) {
                rowNode.classList.add('dt-row-selected');
            }
        });
    }

    /**
     * Clear row selection
     */
    _clearRowSelection() {
        const tableNode = this.table.table().node();

        // Clear dt-row-selected class
        tableNode.querySelectorAll('tr.dt-row-selected').forEach(tr => {
            tr.classList.remove('dt-row-selected');
        });

        // Also clear row-highlight class
        tableNode.querySelectorAll('tr.row-highlight').forEach(tr => {
            tr.classList.remove('row-highlight');
        });

        this._selectedRows = [];
    }

    // ==================== EDITOR METHODS ====================

    /**
     * Activate edit mode on cell (show input/select)
     * @param {object} cell - DataTables cell object
     * @param {boolean} clearContent - true = clear content (replace mode), false = edit existing content
     */
    activateEditCell(cell, clearContent = false) {
        // Hide custom tooltip immediately when entering edit mode
        this._hideTooltip();

        const cellNode = cell.node();

        // Prevent double-activation: if cell is already in edit mode
        // But for dropdown, show options on re-click
        const existingDropdown = cellNode.querySelector('.dt-searchable-dropdown');
        if (existingDropdown) {
            const optionsList = existingDropdown.querySelector('.dt-dropdown-options');
            const searchInput = existingDropdown.querySelector('.dt-dropdown-search');
            if (optionsList && searchInput) {
                optionsList.style.display = 'block';
                searchInput.focus();
                searchInput.select();
            }
            return;
        }
        if (cellNode.querySelector('input, select')) {
            return;
        }

        this.activeCellNode = cellNode;

        // Get current value - special handling for checkbox
        let currentValue = '';
        if (!clearContent) {
            const existingCheckbox = cellNode.querySelector('input[type="checkbox"]');
            if (existingCheckbox) {
                currentValue = existingCheckbox.checked ? 'true' : 'false';
            } else if (cellNode.getAttribute('data-checkbox-value')) {
                currentValue = cellNode.getAttribute('data-checkbox-value');
            } else {
                // Check for select column with span content
                const selectContent = cellNode.querySelector('.dt-select-content');
                if (selectContent) {
                    currentValue = selectContent.textContent.trim();
                } else {
                    // Check for textarea full text first (stored separately from truncated display)
                    const fullText = cellNode.getAttribute('data-full-text');
                    if (fullText !== null) {
                        currentValue = fullText;
                    } else {
                        // Check for formatted value (has data-raw-value)
                        const rawValue = cellNode.getAttribute('data-raw-value');
                        if (rawValue !== null) {
                            currentValue = rawValue;
                        } else {
                            currentValue = this.activeCellNode.textContent.trim();
                        }
                    }
                }
            }
        }
        let rowNode = this.activeCellNode.closest('tr');
        let rowId = rowNode ? rowNode.getAttribute('data-id') : null;
        if (!rowId) return;

        let colIndex = cell.index().column;

        if (this.enableColumnNumber && colIndex === 0) { return false; }

        let colHeader = this.table.column(colIndex).header();
        let fieldName = colHeader.getAttribute('data-name');
        let fieldType = colHeader.getAttribute('data-type');

        // ===== READONLY CHECK =====
        // Check 1: Column header has data-readonly="true"
        if (colHeader.getAttribute('data-readonly') === 'true') {
            console.log('Column is readonly:', fieldName);
            return;
        }
        // Check 2: Cell has data-readonly="true" (individual cell readonly)
        if (cellNode.getAttribute('data-readonly') === 'true') {
            console.log('Cell is readonly');
            return;
        }
        // Check 3: Row has data-readonly-row="true" (entire row readonly)
        if (rowNode.getAttribute('data-readonly-row') === 'true') {
            console.log('Row is readonly');
            return;
        }

        // Check if checkbox - toggle immediately
        if (fieldType === 'checkbox') {
            this._toggleCheckbox(cell);
            return;
        }

        this.showCellEditor(rowId, fieldName, fieldType, currentValue);
    }

    /**
     * Show editor on cell
     */
    async showCellEditor(rowId, fieldName, fieldType, currentValue) {
        let editor;
        // Handle text, number, currency, percentage as input fields
        if (fieldType === 'text' || fieldType === 'number' || fieldType === 'currency' || fieldType === 'percentage') {
            // For currency/percentage, use text input to avoid browser number restrictions
            const inputType = (fieldType === 'currency' || fieldType === 'percentage') ? 'text' : fieldType;
            editor = this.createElement('input', {
                type: inputType,
                value: currentValue,
                className: 'fac-class'
            }, {
                name: fieldName,
                id: rowId,
                fieldtype: fieldType  // Store original type for formatting
            });

            // Add input filter for currency and percentage - only allow valid numeric input
            if (fieldType === 'currency' || fieldType === 'percentage') {
                const formatConfig = this.formats[fieldType] || {};
                const decimalSep = formatConfig.decimalSeparator || ',';

                editor.addEventListener('keypress', (e) => {
                    const char = e.key;
                    const currentValue = editor.value;
                    const cursorPos = editor.selectionStart;

                    // Allow: numbers (0-9)
                    if (/[0-9]/.test(char)) return;

                    // Allow: minus sign only at the beginning
                    if (char === '-' && cursorPos === 0 && !currentValue.includes('-')) return;

                    // Allow: decimal separator only once
                    if (char === decimalSep && !currentValue.includes(decimalSep)) return;

                    // Allow: navigation keys (handled by browser)
                    if (['Backspace', 'Delete', 'ArrowLeft', 'ArrowRight', 'Tab', 'Enter', 'Escape'].includes(e.key)) return;

                    // Block all other characters
                    e.preventDefault();
                });

                // Also filter on paste
                editor.addEventListener('paste', (e) => {
                    e.preventDefault();
                    const pastedText = (e.clipboardData || window.clipboardData).getData('text');
                    const formatConfig = this.formats[fieldType] || {};
                    const decimalSep = formatConfig.decimalSeparator || ',';

                    // Clean pasted text - only keep valid characters
                    let cleaned = '';
                    let hasDecimal = editor.value.includes(decimalSep);
                    let hasMinus = editor.value.includes('-');

                    for (let i = 0; i < pastedText.length; i++) {
                        const char = pastedText[i];
                        if (/[0-9]/.test(char)) {
                            cleaned += char;
                        } else if (char === '-' && cleaned.length === 0 && !hasMinus && editor.selectionStart === 0) {
                            cleaned += char;
                            hasMinus = true;
                        } else if (char === decimalSep && !hasDecimal) {
                            cleaned += char;
                            hasDecimal = true;
                        }
                    }

                    // Insert cleaned text at cursor position
                    const start = editor.selectionStart;
                    const end = editor.selectionEnd;
                    const before = editor.value.substring(0, start);
                    const after = editor.value.substring(end);
                    editor.value = before + cleaned + after;
                    editor.setSelectionRange(start + cleaned.length, start + cleaned.length);
                });
            }
        } else if (fieldType === 'select') {
            // Create searchable dropdown container
            const dropdown = document.createElement('div');
            dropdown.className = 'dt-searchable-dropdown';

            // Search input
            const searchInput = document.createElement('input');
            searchInput.type = 'text';
            searchInput.className = 'dt-dropdown-search';
            searchInput.placeholder = 'Search...';
            dropdown.appendChild(searchInput);

            // Options container (fixed positioned - append to body to escape overflow:hidden)
            const optionsList = document.createElement('div');
            optionsList.className = 'dt-dropdown-options';
            optionsList.style.display = 'none'; // Hide initially
            document.body.appendChild(optionsList); // Append to body for fixed positioning

            // Store reference to cell for positioning
            const activeCellNode = this.activeCellNode;

            // Helper to position dropdown using fixed positioning
            // Smart positioning: opens upward if not enough space below
            const positionDropdown = () => {
                if (activeCellNode) {
                    const cellRect = activeCellNode.getBoundingClientRect();
                    const dropdownHeight = 200; // max-height from CSS
                    const viewportHeight = window.innerHeight;
                    const spaceBelow = viewportHeight - cellRect.bottom;
                    const spaceAbove = cellRect.top;

                    // Use exact width to match cell width precisely
                    optionsList.style.width = cellRect.width + 'px';
                    optionsList.style.minWidth = 'unset';
                    optionsList.style.left = cellRect.left + 'px';

                    // Check if enough space below, otherwise open upward
                    if (spaceBelow >= dropdownHeight || spaceBelow >= spaceAbove) {
                        // Open downward (normal)
                        optionsList.style.top = cellRect.bottom + 'px';
                        optionsList.style.bottom = 'auto';
                        optionsList.style.maxHeight = Math.min(dropdownHeight, spaceBelow - 10) + 'px';
                    } else {
                        // Open upward (not enough space below)
                        optionsList.style.top = 'auto';
                        optionsList.style.bottom = (viewportHeight - cellRect.top) + 'px';
                        optionsList.style.maxHeight = Math.min(dropdownHeight, spaceAbove - 10) + 'px';
                    }
                }
            };

            // Show dropdown helper
            const showDropdown = () => {
                positionDropdown();
                optionsList.style.display = 'block';
            };

            // Hide dropdown helper
            const hideDropdown = () => {
                optionsList.style.display = 'none';
            };

            // Cleanup function to remove dropdown from body and remove event listeners
            const cleanupDropdown = () => {
                if (optionsList.parentNode === document.body) {
                    document.body.removeChild(optionsList);
                }
                // Remove event listeners
                window.removeEventListener('resize', handleResizeScroll);
                window.removeEventListener('scroll', handleResizeScroll, true);
            };

            // Handle resize/scroll - close dropdown to avoid misalignment
            // BUT: don't close if scroll is inside the dropdown itself
            const handleResizeScroll = (e) => {
                // Check if scroll is inside the dropdown
                if (e && e.type === 'scroll' && e.target) {
                    // If scrolling inside optionsList, don't close
                    if (optionsList.contains(e.target) || e.target === optionsList) {
                        return; // Allow scroll inside dropdown
                    }
                }
                hideDropdown();
                // Trigger blur to close editor
                searchInput.blur();
            };

            // Add event listeners for resize and scroll
            // Defer to avoid catching the initial CSS reflow scroll that happens
            // when the cell transitions to edit mode (padding changes)
            setTimeout(() => {
                window.addEventListener('resize', handleResizeScroll);
                window.addEventListener('scroll', handleResizeScroll, true); // Use capture for scroll
            }, 200);


            // Fetch and populate options
            let options = this.selectOptions[fieldName];
            if (typeof options === 'string') {
                // Support both absolute URLs (http/https) and relative paths (.php, etc)
                options = await this.fetchSelectOptions(options);
            }
            if (!options) options = {};

            // Store options for filtering
            const allOptions = Object.entries(options);

            // Track selected value (not filter text!)
            let selectedText = currentValue;
            searchInput.value = currentValue;
            searchInput.setAttribute('data-selected-text', currentValue);

            // Render options function
            const renderOptions = (filterText = '') => {
                optionsList.innerHTML = '';
                const filtered = allOptions.filter(([val, text]) =>
                    text.toLowerCase().includes(filterText.toLowerCase())
                );

                filtered.forEach(([value, text]) => {
                    const optDiv = document.createElement('div');
                    optDiv.className = 'dt-dropdown-option';
                    if (text === selectedText) optDiv.classList.add('selected');
                    optDiv.textContent = text;
                    optDiv.setAttribute('data-value', value);
                    optDiv.addEventListener('mousedown', (e) => {
                        e.preventDefault(); // Prevent blur before click
                        e.stopPropagation();
                        // Update both display and stored value
                        searchInput.value = text;
                        selectedText = text;
                        searchInput.setAttribute('data-selected-text', text);
                        searchInput.setAttribute('data-selected-value', value);
                        optionsList.style.display = 'none';

                        // Store reference before blur
                        const cellToFocus = this.activeCellNode;
                        const tableRef = this.table;
                        const self = this;

                        // Get row info for new row check
                        const rowNode = cellToFocus ? cellToFocus.parentNode : null;
                        const thisRowId = rowNode ? rowNode.getAttribute('data-id') : null;
                        const colIndex = cellToFocus ? cellToFocus.cellIndex : -1;

                        // Blur synchronously
                        searchInput.blur();

                        // For new row: trigger save check (same as doSave flow)
                        if (thisRowId === 'new' || (rowNode && rowNode.getAttribute('data-new-row') === 'true')) {
                            // Update cell value first
                            if (colIndex >= 0) {
                                self._setCellValueWithFormatting(cellToFocus, text, colIndex, true);
                            }

                            setTimeout(() => {
                                if (rowNode && self.isRowRequiredFieldsFilled(rowNode)) {
                                    // Direct mode: save new row
                                    if (!(self.emptyTable.enabled && self.emptyTable.saveMode === 'batch')) {
                                        Promise.resolve(self.saveNewRow(rowNode))
                                            .finally(() => {
                                                if (self.isOnLastPage() && !self.isAddingEmptyRow && !self.hasEmptyRow()) {
                                                    self.addEmptyRow();
                                                }
                                            });
                                    } else {
                                        // Batch mode
                                        const th = tableRef.column(colIndex).header();
                                        const thFieldName = th ? th.getAttribute('data-name') : null;
                                        self._handleNewRowSave(cellToFocus, thFieldName, text);
                                    }
                                }
                            }, 50);
                        }

                        // Immediately add focus classes (before any navigation can happen)
                        if (cellToFocus) {
                            cellToFocus.classList.add('focus', 'dt-cell-selected');
                            self._highlightCell(cellToFocus);

                            // Update internal state on next frame
                            requestAnimationFrame(() => {
                                // Only refocus if we're still on this cell (no navigation happened)
                                if (self.activeCellNode === cellToFocus) {
                                    const cell = tableRef.cell(cellToFocus);
                                    if (cell) {
                                        self.currentCell = cell;
                                        cell.focus();
                                    }
                                }
                            });
                        }
                    });
                    optionsList.appendChild(optDiv);
                });

                if (filtered.length === 0) {
                    const noResult = document.createElement('div');
                    noResult.className = 'dt-dropdown-no-result';
                    noResult.textContent = this.lang.noResult;
                    optionsList.appendChild(noResult);
                }
            };

            // Initial render
            renderOptions();

            // Track highlighted option index for arrow key nav
            let highlightIndex = -1;

            const updateHighlight = () => {
                const opts = optionsList.querySelectorAll('.dt-dropdown-option');
                opts.forEach((opt, i) => {
                    opt.classList.toggle('highlighted', i === highlightIndex);
                });
                // Scroll into view
                if (highlightIndex >= 0 && opts[highlightIndex]) {
                    opts[highlightIndex].scrollIntoView({ block: 'nearest' });
                }
            };

            // Search/filter on input (but don't change selectedText)
            searchInput.addEventListener('input', (e) => {
                renderOptions(e.target.value);
                showDropdown();
                highlightIndex = -1;
            });

            // Show options on focus
            searchInput.addEventListener('focus', () => {
                showDropdown();
                searchInput.select();
                highlightIndex = -1;
            });

            // Click on dropdown container opens options
            dropdown.addEventListener('click', () => {
                showDropdown();
                searchInput.focus();
                searchInput.select();
            });

            // Click on search input also opens options
            searchInput.addEventListener('click', (e) => {
                e.stopPropagation();
                showDropdown();
            });

            // Arrow key navigation - use capture phase to intercept before KeyTable
            searchInput.addEventListener('keydown', (e) => {
                const opts = optionsList.querySelectorAll('.dt-dropdown-option');
                const optCount = opts.length;

                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    e.stopPropagation();
                    e.stopImmediatePropagation();
                    if (optionsList.style.display !== 'block') {
                        showDropdown();
                    }
                    highlightIndex = Math.min(highlightIndex + 1, optCount - 1);
                    updateHighlight();
                    return false;
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    e.stopPropagation();
                    e.stopImmediatePropagation();
                    highlightIndex = Math.max(highlightIndex - 1, 0);
                    updateHighlight();
                    return false;
                } else if (e.key === 'Enter') {
                    e.preventDefault();
                    e.stopPropagation();
                    e.stopImmediatePropagation();

                    const opts = optionsList.querySelectorAll('.dt-dropdown-option');
                    let selectedOpt = null;

                    if (highlightIndex >= 0 && opts[highlightIndex]) {
                        // Select highlighted option
                        selectedOpt = opts[highlightIndex];
                    } else if (opts.length > 0) {
                        // No highlight - select first visible option
                        selectedOpt = opts[0];
                    }

                    if (selectedOpt) {
                        const text = selectedOpt.textContent;
                        const value = selectedOpt.getAttribute('data-value');
                        searchInput.value = text;
                        selectedText = text;
                        searchInput.setAttribute('data-selected-text', text);
                        searchInput.setAttribute('data-selected-value', value);
                    }
                    // Don't dispatch change event - blur handler will handle save

                    optionsList.style.display = 'none';

                    // Store reference before blur
                    const cellToFocus = this.activeCellNode;
                    const tableRef = this.table;
                    const self = this;

                    // Get selected text for new row save
                    const selectedValue = selectedOpt ? selectedOpt.textContent : selectedText;

                    // Get row info for new row check
                    const rowNode = cellToFocus ? cellToFocus.parentNode : null;
                    const thisRowId = rowNode ? rowNode.getAttribute('data-id') : null;
                    const colIndex = cellToFocus ? cellToFocus.cellIndex : -1;

                    // Blur synchronously
                    searchInput.blur();

                    // For new row: trigger save check (same as doSave flow)
                    if (thisRowId === 'new' || (rowNode && rowNode.getAttribute('data-new-row') === 'true')) {
                        // Update cell value first
                        if (colIndex >= 0) {
                            self._setCellValueWithFormatting(cellToFocus, selectedValue, colIndex, true);
                        }

                        setTimeout(() => {
                            if (rowNode && self.isRowRequiredFieldsFilled(rowNode)) {
                                // Direct mode: save new row
                                if (!(self.emptyTable.enabled && self.emptyTable.saveMode === 'batch')) {
                                    Promise.resolve(self.saveNewRow(rowNode))
                                        .finally(() => {
                                            if (self.isOnLastPage() && !self.isAddingEmptyRow && !self.hasEmptyRow()) {
                                                self.addEmptyRow();
                                            }
                                        });
                                } else {
                                    // Batch mode
                                    const th = tableRef.column(colIndex).header();
                                    const thFieldName = th ? th.getAttribute('data-name') : null;
                                    self._handleNewRowSave(cellToFocus, thFieldName, selectedValue);
                                }
                            }
                        }, 50);
                    }

                    // Immediately add focus classes
                    if (cellToFocus) {
                        cellToFocus.classList.add('focus', 'dt-cell-selected');
                        self._highlightCell(cellToFocus);

                        // Update internal state on next frame
                        requestAnimationFrame(() => {
                            // Only refocus if we're still on this cell
                            if (self.activeCellNode === cellToFocus) {
                                const cell = tableRef.cell(cellToFocus);
                                if (cell) {
                                    self.currentCell = cell;
                                    cell.focus();
                                }
                            }
                        });
                    }
                    return false;
                } else if (e.key === 'Escape') {
                    e.preventDefault();
                    e.stopImmediatePropagation();
                    optionsList.style.display = 'none';
                    searchInput.value = selectedText;
                    // Blur to exit edit mode
                    setTimeout(() => searchInput.blur(), 50);
                    return false;
                }
            }, true); // Use capture phase!

            // On blur: restore selected value if user typed filter but didn't select
            const self = this;
            searchInput.addEventListener('blur', () => {
                setTimeout(() => {
                    // Restore to selected value (not filter text)
                    searchInput.value = selectedText;
                    optionsList.style.display = 'none';
                    highlightIndex = -1;
                    // Cleanup dropdown from body
                    cleanupDropdown();
                    // Re-render arrows after a slight delay for cell update
                    setTimeout(() => self._renderSelectColumns(), 50);
                }, 100);
            });

            // Store reference for editor handling
            searchInput.setAttribute('name', fieldName);
            searchInput.setAttribute('id', rowId);
            searchInput.setAttribute('data-dropdown', 'true');

            editor = dropdown;
        } else if (fieldType === 'checkbox') {
            // Checkbox type - create centered checkbox
            editor = this.createElement('input', {
                type: 'checkbox',
                className: 'fac-class dt-checkbox-editor'
            }, {
                name: fieldName,
                id: rowId
            });
            // Set checked based on current value (true/false/1/0)
            const isChecked = currentValue === 'true' || currentValue === '1' || currentValue === true;
            editor.checked = isChecked;
        } else if (fieldType === 'date') {
            // Date type - use native HTML5 date picker
            // Convert display format back to ISO for date input
            const isoValue = this._parseDateToISO(currentValue);
            editor = this.createElement('input', {
                type: 'date',
                value: isoValue,
                className: 'fac-class dt-date-editor'
            }, {
                name: fieldName,
                id: rowId,
                fieldtype: fieldType
            });

            // Date follows save-on-commit pattern - NO change event listener
            // Save is handled by blur event in setupEditorEvents()
        } else if (fieldType === 'time') {
            // Time type - use native HTML5 time picker
            editor = this.createElement('input', {
                type: 'time',
                value: currentValue,
                className: 'fac-class dt-time-editor'
            }, {
                name: fieldName,
                id: rowId,
                fieldtype: fieldType
            });

            // Time follows save-on-commit pattern - NO change event listener
            // Save is handled by blur event in attachEditorEvent()
        } else if (fieldType === 'email') {
            // Email type - input with validation
            editor = this.createElement('input', {
                type: 'email',
                value: currentValue,
                className: 'fac-class dt-email-editor'
            }, {
                name: fieldName,
                id: rowId,
                fieldtype: fieldType
            });

            // Add validation on blur
            editor.addEventListener('blur', () => {
                if (editor.value && !this._isValidEmail(editor.value)) {
                    editor.classList.add('dt-invalid');
                    editor.setAttribute('title', 'Invalid email format');
                } else {
                    editor.classList.remove('dt-invalid');
                    editor.removeAttribute('title');
                }
            });
        } else if (fieldType === 'textarea') {
            // Check edit mode preference
            const textareaConfig = this.formats.textarea || {};
            const editMode = textareaConfig.editMode || 'modal';

            if (editMode === 'modal') {
                // MODAL MODE - Show textarea in a modal overlay
                const overlay = this._createTextareaModal(fieldName, rowId, currentValue);
                document.body.appendChild(overlay);

                // Focus the textarea inside modal
                const modalTextarea = overlay.querySelector('.dt-textarea-editor');
                if (modalTextarea) {
                    setTimeout(() => modalTextarea.focus(), 50);
                }

                // Don't set editor - modal handles its own save/close
                return;
            }

            // INLINE MODE - Original inline editing
            editor = document.createElement('textarea');
            editor.className = 'fac-class dt-inline-textarea';
            editor.value = currentValue;
            editor.setAttribute('name', fieldName);
            editor.setAttribute('id', rowId);
            editor.setAttribute('data-fieldtype', fieldType);

            // Position the textarea near the cell (for fixed positioning)
            const cellRect = this.activeCellNode.getBoundingClientRect();
            editor.style.top = cellRect.top + 'px';
            editor.style.left = cellRect.left + 'px';
            // Match cell width for consistency - reduce slightly for margins/borders
            editor.style.width = (cellRect.width - 8) + 'px';
            editor.style.minWidth = 'unset';

            // Auto-resize based on content - expand to fit
            const adjustHeight = () => {
                // Reset height to auto to get accurate scrollHeight
                editor.style.height = 'auto';
                // Set height to scrollHeight (includes padding due to box-sizing: border-box)
                editor.style.height = Math.max(editor.scrollHeight, 40) + 'px';
            };
            editor.addEventListener('input', adjustHeight);
            // Initial height adjustment after render
            setTimeout(adjustHeight, 0);

            // Handle Enter (save) vs Shift+Enter (newline) vs Tab (save and move)
            editor.addEventListener('keydown', (e) => {
                // Enter without Shift = save and close edit mode, stay on this cell
                if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    e.stopPropagation();

                    // Mark as handled to prevent blur handler from processing
                    editor._handled = true;

                    const newValue = editor.value;

                    // Remove textarea from DOM
                    if (editor.parentNode) {
                        editor.parentNode.removeChild(editor);
                    }

                    this.activeCellNode.classList.remove('dt-cell-editing');
                    this.isEditMode = false;

                    // Update cell value if changed
                    if (newValue !== currentValue) {
                        this._setCellValueWithFormatting(this.activeCellNode, newValue, this.activeCellNode.cellIndex, true);

                        // Trigger server update via updateCell
                        if (rowId !== 'new') {
                            this.updateCell(fieldName, rowId, newValue);
                        }
                    } else {
                        // Restore original value to display
                        this._setCellValueWithFormatting(this.activeCellNode, currentValue, this.activeCellNode.cellIndex, false);
                    }

                    // For new row - trigger save check
                    if (rowId === 'new') {
                        setTimeout(() => {
                            this._handleNewRowSave(this.activeCellNode, fieldName, newValue);
                        }, 0);
                    }

                    // Keep focus on this cell with proper indicator
                    setTimeout(() => {
                        if (this.currentCell) {
                            this.currentCell.focus();
                            // Add selected class for green border
                            this.activeCellNode.classList.add('dt-cell-selected');
                        }
                    }, 10);
                    return;
                }

                // Tab = save and move to next cell
                if (e.key === 'Tab') {
                    e.preventDefault();
                    editor.blur(); // Trigger save via blur handler
                    return;
                }

                // All Arrow keys = stay in edit mode (for navigating text)
                if (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
                    e.stopPropagation(); // Prevent DataTables from moving focus
                    // Allow default textarea behavior (cursor movement)
                    return;
                }

                // Escape to cancel
                if (e.key === 'Escape') {
                    e.preventDefault();
                    editor._handled = true; // Mark as handled
                    editor.value = currentValue; // Restore original
                    editor.blur();
                    return;
                }
            });

            // Blur handler to save when moving cells
            const activeCellRef = this.activeCellNode;
            const self = this;
            editor.addEventListener('blur', () => {
                // Skip if already handled by Enter/Escape
                if (editor._handled) {
                    return;
                }

                const newValue = editor.value;

                // Remove textarea from DOM
                if (editor.parentNode) {
                    editor.parentNode.removeChild(editor);
                }

                activeCellRef.classList.remove('dt-cell-editing');
                self.isEditMode = false;

                // Cleanup scroll/wheel listeners
                if (editor._cleanupScroll) {
                    editor._cleanupScroll();
                }

                // Update cell value - always restore display (even if unchanged)
                if (newValue !== currentValue) {
                    self._setCellValueWithFormatting(activeCellRef, newValue, activeCellRef.cellIndex, true);

                    // Trigger server update via updateCell or handle new row
                    if (rowId !== 'new') {
                        self.updateCell(fieldName, rowId, newValue);
                    } else {
                        // For new row - trigger save check
                        setTimeout(() => {
                            self._handleNewRowSave(activeCellRef, fieldName, newValue);
                        }, 0);
                    }
                } else {
                    // Restore original value to cell display (since we cleared it for editor)
                    self._setCellValueWithFormatting(activeCellRef, currentValue, activeCellRef.cellIndex, false);

                    // For new row - still check if row is complete
                    if (rowId === 'new') {
                        setTimeout(() => {
                            self._handleNewRowSave(activeCellRef, fieldName, currentValue);
                        }, 0);
                    }
                }
            });

            // Close on scroll/wheel (like tooltip/dropdown behavior)
            const closeOnScroll = () => {
                if (editor.parentNode) {
                    // Mark as handled to skip blur save
                    editor._handled = true;

                    // Cleanup scroll listeners first
                    if (editor._cleanupScroll) {
                        editor._cleanupScroll();
                    }

                    // Remove textarea from DOM
                    editor.parentNode.removeChild(editor);

                    // Restore cell state
                    activeCellRef.classList.remove('dt-cell-editing');
                    self.isEditMode = false;

                    // Restore original value (cancel edit)
                    self._setCellValueWithFormatting(activeCellRef, currentValue, activeCellRef.cellIndex, false);
                }
            };

            const wheelHandler = (e) => {
                // Ignore wheel events inside the textarea itself
                if (e.target === editor || editor.contains(e.target)) {
                    return;
                }
                closeOnScroll();
            };

            const scrollHandler = (e) => {
                // Ignore scroll events from the textarea itself
                if (e.target === editor || editor.contains(e.target)) {
                    return;
                }
                closeOnScroll();
            };

            // Store cleanup function for blur handler
            editor._cleanupScroll = () => {
                window.removeEventListener('wheel', wheelHandler, true);
                window.removeEventListener('scroll', scrollHandler, true);
                document.removeEventListener('scroll', scrollHandler, true);
            };

            // Listen on wheel and scroll events (capture phase)
            window.addEventListener('wheel', wheelHandler, true);
            window.addEventListener('scroll', scrollHandler, true);
            document.addEventListener('scroll', scrollHandler, true);
        } else if (fieldType === 'readonly') {
            // Readonly type - don't allow editing
            console.log('Readonly cell - editing not allowed');
            this.isEditMode = false;
            return;
        }

        if (editor) {
            // Add class for edit mode styling
            this.activeCellNode.classList.add('dt-cell-editing');
            this.activeCellNode.classList.remove('dt-cell-selected');

            this.activeCellNode.textContent = '';
            this.activeCellNode.appendChild(editor);

            // Handle focus for searchable dropdown vs regular editor
            if (editor.classList.contains('dt-searchable-dropdown')) {
                const searchInput = editor.querySelector('.dt-dropdown-search');
                if (searchInput) {
                    searchInput.focus();
                    searchInput.select();
                    this.attachEditorEvent(searchInput);
                }
            } else {
                editor.focus();
                // Select all text for input (for easy replacement)
                if (editor.tagName.toLowerCase() === 'input') {
                    editor.select();
                }
                // SKIP attachEditorEvent for textarea as it already has its own blur handler
                if (editor.tagName.toLowerCase() !== 'textarea') {
                    this.attachEditorEvent(editor);
                }
            }
        }
    }

    /**
     * Helper to create input/select element with dataset attributes
     */
    createElement(elementType, attrs = {}, dataset = {}) {
        const el = document.createElement(elementType);
        Object.assign(el, attrs);
        Object.assign(el.dataset, dataset);
        return el;
    }

    /**
     * Fetch select options via endpoint, cache in memory
     */
    async fetchSelectOptions(endpointUrl) {
        if (this._selectOptionCache[endpointUrl]) return this._selectOptionCache[endpointUrl];

        try {
            const response = await fetch(endpointUrl, {
                method: 'GET',
                headers: this.getExtraHeaders()
            });
            if (!response.ok) throw new Error('Failed to fetch select options!');
            const data = await response.json();
            let optionsObj = {};
            if (Array.isArray(data)) {
                data.forEach(opt => { optionsObj[opt.id] = opt.text; });
            } else {
                optionsObj = data;
            }
            this._selectOptionCache[endpointUrl] = optionsObj;
            return optionsObj;
        } catch (error) {
            console.error(error);
            return {};
        }
    }

    /**
     * Attach events to editor input
     */
    attachEditorEvent(editorEl) {
        let fieldName = editorEl.getAttribute('data-name') || editorEl.getAttribute('name');
        // Get rowId from editor attribute OR from row node (more reliable for batch mode temp_ IDs)
        let rowId = editorEl.getAttribute('data-id') || editorEl.getAttribute('id');

        // Also check the actual row node's data-id (more reliable for updated temp_ IDs)
        const rowNode = this.activeCellNode?.parentNode;
        if (rowNode && rowNode.getAttribute('data-id')) {
            rowId = rowNode.getAttribute('data-id');
        }

        this.savingCell = false;

        // Store local reference to cell node (to avoid stale reference when changing pages)
        const localCellNode = this.activeCellNode;

        let oldValue = "";
        if (editorEl.tagName.toLowerCase() === 'input') {
            if (editorEl.type === 'checkbox') {
                oldValue = editorEl.checked ? 'true' : 'false';
            } else if (editorEl.hasAttribute('data-selected-text')) {
                // Searchable dropdown - use data-selected-text for consistency
                oldValue = editorEl.getAttribute('data-selected-text') || editorEl.value;
            } else {
                oldValue = editorEl.value;
            }
        } else if (editorEl.tagName.toLowerCase() === 'select') {
            oldValue = editorEl.options[editorEl.selectedIndex]?.text || '';
        }


        function debounce(fn, wait) {
            let t;
            return function (...args) {
                clearTimeout(t);
                t = setTimeout(() => fn.apply(this, args), wait);
            };
        }

        const doSave = (fromEvent) => {
            if (this.savingCell) return;

            // Check if editor and cell still exist in DOM
            if (!document.body.contains(editorEl) || !document.body.contains(localCellNode)) {
                this.savingCell = false;
                return;
            }

            this.savingCell = true;

            // Get new value using centralized helper
            const newValue = this._getEditorValue(editorEl);

            // Centralized validation (includes preSave callback)
            const isValid = this._validateBeforeSave(localCellNode, editorEl, newValue, rowId, fieldName);

            if (!isValid) {
                this._handleValidationError(localCellNode, newValue);
                this.savingCell = false;
                return;
            }

            // Clear error state if validation passed
            localCellNode.classList.remove('dt-error');

            // Debug log

            // Skip save if value unchanged (for existing rows only)
            if (rowId !== 'new' && newValue === oldValue) {
                this._setCellValueWithFormatting(localCellNode, newValue, localCellNode.cellIndex, true);
                this.savingCell = false;
                return;
            }

            // Update cell display with new value
            this._setCellValueWithFormatting(localCellNode, newValue, localCellNode.cellIndex, true);
            setTimeout(() => this._renderSelectColumns(), 10);

            // Recalculate formula columns for this row
            const rowNode = localCellNode.closest('tr');
            if (rowNode) {
                this._recalculateRowFormulas(rowNode);
            }
            // Update footer totals after data change
            this._updateFooterTotals();

            // Route to appropriate save handler
            if (rowId !== 'new') {
                // EXISTING ROW: Update via updateCell (handles both direct and batch mode)
                Promise.resolve(this.updateCell(fieldName, rowId, newValue))
                    .finally(() => {
                        this.savingCell = false;
                    });
            } else {
                // NEW ROW: Check if row is complete after DOM update
                // Use setTimeout(0) to ensure _setCellValueWithFormatting has updated the DOM
                setTimeout(() => {
                    this._handleNewRowSave(localCellNode, fieldName, newValue);
                }, 0);
                this.savingCell = false;
            }

            return true;
        };

        const debouncedSave = debounce(doSave, 80);

        editorEl.addEventListener('keydown', e => {
            const isSelect = editorEl.tagName.toLowerCase() === 'select';

            if (isSelect) {
                // For select: 
                // - Enter = open dropdown (if browser supports showPicker)
                // - Tab = save and move
                // - ArrowUp/Down = navigate options (let default)
                if (e.key === 'Enter') {
                    e.preventDefault();
                    e.stopPropagation();
                    // Open dropdown with showPicker (modern browsers) or click simulation
                    if (typeof editorEl.showPicker === 'function') {
                        editorEl.showPicker();
                    } else {
                        // Fallback: simulate mouse event to open dropdown
                        const event = new MouseEvent('mousedown', { bubbles: true });
                        editorEl.dispatchEvent(event);
                    }
                    return;
                }
                if (e.key === 'Tab') {
                    let saveRes = doSave(e);
                    if (saveRes === false) {
                        e.preventDefault();
                        e.stopPropagation();
                        setTimeout(() => editorEl.focus(), 1);
                    }
                    this.isEditMode = false; // Exit edit mode
                    return;
                }
                // ArrowUp/Down let default for option navigation
                return;
            } else {
                // ===== GOOGLE SHEETS BEHAVIOR for INPUT =====

                // Escape = Cancel, exit edit mode without save
                if (e.key === 'Escape') {
                    e.preventDefault();
                    e.stopPropagation();
                    // Restore original value
                    editorEl.value = oldValue;
                    doSave(e);
                    this.isEditMode = false;
                    return;
                }

                // Enter = Save and exit edit mode (stay on same cell or move down)
                // For checkbox: Enter should toggle before saving
                if (e.key === 'Enter') {
                    e.preventDefault();
                    e.stopPropagation();

                    // Special handling for checkbox
                    if (editorEl.type === 'checkbox') {
                        editorEl.checked = !editorEl.checked;
                        editorEl.dispatchEvent(new Event('change', { bubbles: true }));
                        // Re-focus the cell so subsequent key events work
                        setTimeout(() => {
                            const cell = this.table.cell(localCellNode);
                            if (cell && cell.node()) {
                                cell.focus();
                                localCellNode.classList.add('dt-cell-selected');
                            }
                        }, 50);
                    } else {
                        doSave(e);
                        // Re-focus the cell after save (stay on same cell in selection mode)
                        setTimeout(() => {
                            const cell = this.table.cell(localCellNode);
                            if (cell && cell.node()) {
                                cell.focus();
                                localCellNode.classList.add('dt-cell-selected');
                            }
                        }, 50);
                    }
                    this.isEditMode = false;
                    return;
                }

                // Space for checkbox = toggle
                if (e.key === ' ' && editorEl.type === 'checkbox') {
                    e.preventDefault();
                    e.stopPropagation();
                    editorEl.checked = !editorEl.checked;
                    editorEl.dispatchEvent(new Event('change', { bubbles: true }));
                    return;
                }

                // Tab = Save and move to next cell
                if (e.key === 'Tab') {
                    doSave(e);
                    this.isEditMode = false;

                    // Empty Table: Tab on last cell of last row = ALWAYS add new row
                    if (this.emptyTable.enabled) {
                        const cell = this.table.cell(localCellNode);
                        if (cell) {
                            const colIndex = cell.index().column;
                            const rowIndex = cell.index().row;
                            const totalRows = this.table.rows().count();
                            const totalCols = this.table.columns().count();
                            const isLastRow = rowIndex === totalRows - 1;
                            const isLastCol = colIndex === totalCols - 1;

                            if (isLastRow && isLastCol) {
                                e.preventDefault();
                                e.stopPropagation();
                                this._addNewEmptyRow();
                                setTimeout(() => {
                                    const newRowIdx = this.table.rows().count() - 1;
                                    const firstEditableCol = this.enableColumnNumber ? 1 : 0;
                                    const newCell = this.table.cell(newRowIdx, firstEditableCol);
                                    if (newCell && newCell.node()) {
                                        newCell.focus();
                                    }
                                }, 50);
                            }
                        }
                    }
                    return;
                }

                // Arrow Left/Right = Navigate cursor in text (LET DEFAULT, don't save)
                if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
                    e.stopPropagation(); // Stop bubbling to DataTables but let default behavior
                    return; // Don't save, let cursor move
                }

                // Arrow Up/Down = Save and navigate to cell above/below
                if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
                    doSave(e);
                    this.isEditMode = false;

                    // Empty Table: Arrow Down on last row = ALWAYS add new row
                    if (this.emptyTable.enabled && e.key === 'ArrowDown') {
                        const cell = this.table.cell(localCellNode);
                        if (cell) {
                            const rowIndex = cell.index().row;
                            const totalRows = this.table.rows().count();
                            const isLastRow = rowIndex === totalRows - 1;

                            if (isLastRow) {
                                e.preventDefault();
                                e.stopPropagation();
                                this._addNewEmptyRow();
                                setTimeout(() => {
                                    const newRowIdx = this.table.rows().count() - 1;
                                    const firstEditableCol = this.enableColumnNumber ? 1 : 0;
                                    const newCell = this.table.cell(newRowIdx, firstEditableCol);
                                    if (newCell && newCell.node()) {
                                        newCell.focus();
                                    }
                                }, 50);
                            }
                        }
                    }
                    return;
                }
            }
        });

        // For select: DON'T save immediately when value changes
        // Save will be done on blur (when moving to another column)
        if (editorEl.tagName.toLowerCase() === 'select') {
            editorEl.addEventListener('change', e => {
                // Only update visual, save will be on blur
                this.isEditMode = false;
            });
        }

        // For checkbox: DON'T save immediately when checked changes
        // Save will be done on blur (when moving to another column)
        if (editorEl.tagName.toLowerCase() === 'input' && editorEl.type === 'checkbox') {
            editorEl.addEventListener('change', e => {
                // Update visual only, save will be on blur
                const newValue = editorEl.checked ? 'true' : 'false';
                localCellNode.classList.remove('dt-cell-editing');
                // Use centralized helper to render checkbox display correctly
                this._setCellValueWithFormatting(localCellNode, newValue, localCellNode.cellIndex, true);

                // Mark value as changed for blur handler
                editorEl.setAttribute('data-value-changed', 'true');
            });
        }

        // For searchable dropdown: DON'T save immediately when value changes
        // Save will be done on blur (when moving to another column)
        let dropdownValueChanged = false;
        if (editorEl.tagName.toLowerCase() === 'input' && editorEl.hasAttribute('data-dropdown')) {
            editorEl.addEventListener('change', e => {
                dropdownValueChanged = true;
                this.isEditMode = false;
            });
        }

        editorEl.addEventListener('blur', e => {
            // Remove editing class
            localCellNode.classList.remove('dt-cell-editing');
            this.isEditMode = false;

            // For checkbox and select, save is deferred until moving to another cell
            // Save pending data to be saved when focus moves to another cell
            const editorType = editorEl.type || editorEl.tagName.toLowerCase();
            const isCheckbox = editorType === 'checkbox';
            const isSelect = editorEl.tagName.toLowerCase() === 'select';
            const isDropdown = editorEl.hasAttribute('data-dropdown');

            if (isCheckbox || isSelect || isDropdown) {
                // Get new value
                const newValue = this._getEditorValue(editorEl);

                // Update cell display IMMEDIATELY (so it's not empty)
                this._setCellValueWithFormatting(localCellNode, newValue, localCellNode.cellIndex, true);

                // SAVE IMMEDIATELY when value changes (Save on Commit)
                if (newValue !== oldValue && rowId !== 'new') {
                    this.updateCell(fieldName, rowId, newValue);
                }
                return;
            }

            // For other types (text, number, etc), save immediately on blur
            let saveRes = doSave(e);
            if (saveRes === false) setTimeout(() => editorEl.focus(), 1);
        });
    }

    /**
     * Check if all fields in row are filled
     * Columns with data-empty="true" in header can be empty
     */
    isRowRequiredFieldsFilled(row) {
        const results = Array.from(row.children).map((cell, index) => {
            // First column (No) can be empty if enableColumnNumber is active
            if (this.enableColumnNumber && index === 0) {
                return { index, pass: true, reason: 'No. column' };
            }

            // Check if this column can be empty (data-empty="true")
            const header = this.table.column(index).header();
            const fieldName = header ? header.getAttribute('data-name') : 'unknown';

            if (header && header.getAttribute('data-empty') === 'true') {
                return { index, fieldName, pass: true, reason: 'data-empty allowed' };
            }

            // Get column type
            const dataType = header ? header.getAttribute('data-type') : null;
            const isReadonly = header ? header.getAttribute('data-readonly') === 'true' : false;

            // Readonly columns don't need user input - always considered filled
            // Check both data-type="readonly" AND data-readonly="true"
            if (dataType === 'readonly' || isReadonly) {
                return { index, fieldName, pass: true, reason: 'readonly' };
            }

            // Checkbox column: always considered filled (has true/false value)
            if (dataType === 'checkbox') {
                // Check if checkbox exists or cell has data-checkbox-value attribute
                const checkbox = cell.querySelector('input[type="checkbox"]');
                if (checkbox || cell.hasAttribute('data-checkbox-value')) {
                    return { index, fieldName, pass: true, reason: 'checkbox' };
                }
                // Fallback: check text content for 'true' or 'false'
                const text = cell.textContent.trim().toLowerCase();
                const pass = text === 'true' || text === 'false';
                return { index, fieldName, pass, reason: 'checkbox-text', value: text };
            }

            // Use centralized _getCellRawValue helper for all other types
            const cellValue = this._getCellRawValue(cell, index);
            const pass = cellValue !== '' && cellValue !== null && cellValue !== undefined;
            return { index, fieldName, dataType, pass, value: cellValue };
        });

        // Log results
        results.forEach(r => {
            if (!r.pass) {
            } else {
            }
        });

        const allFilled = results.every(r => r.pass);
        return allFilled;
    }

    // ==================== SAVE OPERATIONS ====================

    /**
     * Update cell value and send to server or save to localStorage
     * 
     * Behavior based on mode:
     * - Direct mode: Fetch to update endpoint immediately
     * - Batch mode: Save to localStorage, send on "Save All"
     * 
     * @param {string} fieldName - Name of field/column to update
     * @param {string|number} rowId - Row ID from database or temp_xxx for new row
     * @param {string} value - New value for cell
     * @returns {void}
     * 
     * @example
     * // Update cell di direct mode
     * this.updateCell('name', 1, 'John Doe');
     * 
     * @example
     * // Update cell di batch mode (temp row)
     * this.updateCell('email', 'temp_123', 'john@example.com');
     */
    updateCell(fieldName, rowId, value) {
        let updateConf = this.endpoints.update || {};
        let endpointUrl = updateConf.endpoint || '';

        // DEBUG: Log all conditions

        // allowedFields filter
        if (this.allowedFields && Array.isArray(this.allowedFields)) {
            if (!this.allowedFields.includes(fieldName)) {
                console.log('updateCell: field not in allowedFields, skipping:', fieldName);
                return;
            }
        }

        // ========== BATCH MODE: Update localStorage ==========
        if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
            // Check if rowId is temp ID (pending NEW row)
            // IMPORTANT: Also check if this temp row is actually in _pendingData
            // If it was already saved but kept temp_ ID, it should be treated as an existing row
            if (rowId && rowId.startsWith('temp_')) {
                // Check if this temp row exists in pending data (not yet saved)
                const isInPendingData = this._pendingData.some(row => row._rowTempId === rowId);

                // Also check if row has data-pending attribute
                const rowNode = this.table.table().node().querySelector(`tr[data-id="${rowId}"]`);
                const hasPendingAttr = rowNode?.getAttribute('data-pending') === 'true';

                // Only treat as pending if it's actually in pending data OR has pending attribute
                if (isInPendingData || hasPendingAttr) {
                    this._updateLocalStorageEntry(rowId, fieldName, value);

                    // If row is now complete, check if we need to add an empty row
                    const isComplete = rowNode ? this.isRowRequiredFieldsFilled(rowNode) : false;
                    if (isComplete && rowNode) {
                        // Remove data-new-row since it's no longer a new empty row
                        rowNode.removeAttribute('data-new-row');

                        // Add empty row if needed

                        if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                            this.addEmptyRow();
                        } else {
                        }
                    }
                    return;
                } else {
                    // Row has temp_ ID but was already saved (not in pending data)
                    // Treat as existing row - fall through to _trackEditedRow below
                }
            }

            // If rowId is not temp_ OR is temp_ but already saved (not in pending), treat as EXISTING row
            if (rowId && rowId !== 'new') {
                this._trackEditedRow(rowId, fieldName, value);
                return;
            }
        }

        // ========== DIRECT MODE: Send to server ==========
        if (!endpointUrl) {
            console.warn('Update endpoint not configured!');
            return;
        }

        // Get row and column indexes for meta
        const rowNode = this.table.table().node().querySelector(`tr[data-id="${rowId}"]`);
        const rowIndex = rowNode ? this.table.row(rowNode).index() : null;

        // Get column index from field name
        const headers = this.table.table().node().querySelectorAll('thead th');
        let colIndex = null;
        headers.forEach((th, idx) => {
            if (th.getAttribute('data-name') === fieldName) {
                colIndex = idx;
            }
        });

        // Consistent payload structure
        const payload = {
            operation: 'update',
            data: {
                id: rowId,
                fields: {
                    [fieldName]: value
                }
            },
            meta: {
                timestamp: Date.now(),
                hasErrors: false,
                validationErrors: [],
                rowCount: 1,
                fieldCount: 1,
                rowIndex: rowIndex,
                colIndex: colIndex
            }
        };

        if (typeof updateConf.preSave === 'function') {
            let result = updateConf.preSave(payload);
            if (result === false) return; // Cancel update process
        }

        fetch(endpointUrl, {
            method: 'POST',
            headers: {
                ...this.getExtraHeaders(),
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        })
            .then(async res => {
                if (!res.ok) {
                    const txt = await res.text();
                    throw new Error('Server error: ' + txt);
                }
                return res.json();
            })
            .then(data => {
                if (typeof updateConf.postSave === 'function') updateConf.postSave(data);
            })
            .catch(err => {
                console.error('[DataTable] Update failed:', err);
            });
    }

    /**
     * Update multiple fields from one row in one API call
     * 
     * Used primarily for paste operations for efficiency (1 call per row).
     * Payload sent: { id: rowId, fields: { field1: value1, ... } }
     * 
     * @param {string|number} rowId - Row ID to update
     * @param {Object} fieldsData - Object containing {fieldName: value, ...}
     * @returns {void}
     * 
     * @example
     * this.updateRowBatch(1, {
     *     name: 'John Doe',
     *     email: 'john@example.com',
     *     office: 'Jakarta'
     * });
     */
    updateRowBatch(rowId, fieldsData) {
        let updateConf = this.endpoints.update || {};
        let endpointUrl = updateConf.endpoint || '';

        // DEBUG logging

        if (!rowId || !fieldsData || Object.keys(fieldsData).length === 0) {
            return;
        }

        // Filter by allowedFields
        let filteredData = fieldsData;
        if (this.allowedFields && Array.isArray(this.allowedFields)) {
            filteredData = {};
            this.allowedFields.forEach(fieldName => {
                if (fieldsData[fieldName] !== undefined) {
                    filteredData[fieldName] = fieldsData[fieldName];
                }
            });
        }

        if (Object.keys(filteredData).length === 0) {
            return;
        }

        // BATCH MODE: Update localStorage
        if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
            if (rowId.startsWith('temp_')) {
                // Update pending data
                Object.keys(filteredData).forEach(fieldName => {
                    this._updateLocalStorageEntry(rowId, fieldName, filteredData[fieldName]);
                });
            } else {
                // Track as edited
                Object.keys(filteredData).forEach(fieldName => {
                    this._trackEditedRow(rowId, fieldName, filteredData[fieldName]);
                });
            }
            return;
        }

        // DIRECT MODE: Send to server
        if (!endpointUrl) {
            console.warn('Update endpoint not configured!');
            return;
        }

        // Get row index for meta
        const rowNode = this.table.table().node().querySelector(`tr[data-id="${rowId}"]`);
        const rowIndex = rowNode ? this.table.row(rowNode).index() : null;

        // Consistent payload structure
        const payload = {
            operation: 'update',
            data: {
                id: rowId,
                fields: filteredData
            },
            meta: {
                timestamp: Date.now(),
                hasErrors: false,
                validationErrors: [],
                rowCount: 1,
                fieldCount: Object.keys(filteredData).length,
                rowIndex: rowIndex
            }
        };

        if (typeof updateConf.preSave === 'function') {
            let result = updateConf.preSave(payload);
            if (result === false) return;
        }

        fetch(endpointUrl, {
            method: 'POST',
            headers: {
                ...this.getExtraHeaders(),
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        })
            .then(async res => {
                if (!res.ok) {
                    const txt = await res.text();
                    throw new Error('Server error: ' + txt);
                }
                return res.json();
            })
            .then(data => {
                if (typeof updateConf.postSave === 'function') updateConf.postSave(data);
            })
            .catch(err => {
                console.error('[DataTable] Update batch failed:', err);
            });
    }

    /**
     * Update multiple rows in a single API call (true batch update)
     * 
     * Used for paste operations to minimize network overhead.
     * Payload: { rows: [{ id: rowId, fields: {...} }, ...] }
     * 
     * @param {Array} rowsData - Array of { id: rowId, fields: {...} }
     * @returns {void}
     * 
     * @example
     * this.updateMultipleRows([
     *     { id: 1, fields: { name: 'John', email: 'john@test.com' } },
     *     { id: 2, fields: { name: 'Jane', email: 'jane@test.com' } }
     * ]);
     */
    updateMultipleRows(rowsData) {
        if (!rowsData || rowsData.length === 0) {
            return;
        }

        let updateConf = this.endpoints.update || {};
        let endpointUrl = updateConf.endpoint || '';

        // BATCH MODE: Just log, data already tracked
        if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
            return;
        }

        // DIRECT MODE: Send all rows to server in one call
        if (!endpointUrl) {
            console.warn('Update endpoint not configured!');
            return;
        }

        // Filter fields by allowedFields for each row
        const filteredRowsData = rowsData.map(row => {
            let filteredFields = row.fields;
            if (this.allowedFields && Array.isArray(this.allowedFields)) {
                filteredFields = {};
                this.allowedFields.forEach(fieldName => {
                    if (row.fields[fieldName] !== undefined) {
                        filteredFields[fieldName] = row.fields[fieldName];
                    }
                });
            }
            return { id: row.id, fields: filteredFields };
        }).filter(row => Object.keys(row.fields).length > 0);

        if (filteredRowsData.length === 0) {
            return;
        }

        // Calculate total field count
        const totalFieldCount = filteredRowsData.reduce((sum, row) => sum + Object.keys(row.fields).length, 0);

        // Consistent payload structure
        const payload = {
            operation: 'update_batch',
            data: {
                rows: filteredRowsData
            },
            meta: {
                timestamp: Date.now(),
                hasErrors: false,
                validationErrors: [],
                rowCount: filteredRowsData.length,
                fieldCount: totalFieldCount
            }
        };


        if (typeof updateConf.preSave === 'function') {
            let result = updateConf.preSave(payload);
            if (result === false) return;
        }

        fetch(endpointUrl, {
            method: 'POST',
            headers: {
                ...this.getExtraHeaders(),
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        })
            .then(async res => {
                if (!res.ok) {
                    const txt = await res.text();
                    throw new Error('Server error: ' + txt);
                }
                return res.json();
            })
            .then(data => {
                if (typeof updateConf.postSave === 'function') updateConf.postSave(data);
            })
            .catch(err => {
                console.error('[DataTable] Batch update failed:', err);
            });
    }

    /**
     * Save new row to server via save endpoint
     * 
     * Collects all cell data from row, filters by allowedFields,
     * runs preSave callback, sends to server, then runs postSave.
     * 
     * @param {HTMLTableRowElement} rowNode - DOM element TR to save
     * @returns {void}
     * 
     * @fires preSave - Called before save with row data
     * @fires postSave - Called after save succeeds with server response
     * 
     * @example
     * const row = document.querySelector('tr[data-id="new"]');
     * this.saveNewRow(row);
     */
    saveNewRow(rowNode) {
        let saveConf = this.endpoints.save || {};
        let endpointUrl = saveConf.endpoint || '';

        // Collect data from all cells in row
        const rowData = {};
        Array.from(rowNode.children).forEach((td, idx) => {
            const th = this.table.column(idx).header();
            const fieldName = th.getAttribute('data-name');
            if (fieldName) {
                rowData[fieldName] = this._getCellValue(td, idx);
            }
        });

        // Filter allowed fields
        let filteredRowData = rowData;
        if (this.allowedFields && Array.isArray(this.allowedFields)) {
            filteredRowData = {};
            this.allowedFields.forEach(f => {
                if (rowData[f] !== undefined) filteredRowData[f] = rowData[f];
            });
        }

        // Get row index for meta
        const rowIndex = this.table.row(rowNode).index();
        const tempIdForPayload = rowNode.getAttribute('data-id') || 'temp_' + Date.now();

        // Build consistent payload for preSave
        const preSavePayload = {
            operation: 'insert',
            data: {
                row: filteredRowData
            },
            meta: {
                timestamp: Date.now(),
                hasErrors: false,
                validationErrors: [],
                rowCount: 1,
                tempId: tempIdForPayload,
                fieldCount: Object.keys(filteredRowData).length,
                rowIndex: rowIndex
            }
        };

        if (typeof saveConf.preSave === 'function') {
            let result = saveConf.preSave(preSavePayload);
            if (result === false) return; // Cancel save process
        }

        // ========== BATCH MODE: Save to localStorage ==========
        if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
            // Check if this row already has temp_id (already saved before)
            const existingTempId = rowNode.getAttribute('data-id');

            if (existingTempId && existingTempId.startsWith('temp_')) {
                // Row already exists in localStorage, update all fields
                const entryIndex = this._pendingData.findIndex(entry => entry._rowTempId === existingTempId);
                if (entryIndex !== -1) {
                    // Update all fields
                    Object.keys(filteredRowData).forEach(key => {
                        this._pendingData[entryIndex][key] = filteredRowData[key];
                    });
                    this._pendingData[entryIndex]._timestamp = Date.now();

                    try {
                        localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
                    } catch (e) {
                        console.error('Error updating localStorage:', e);
                    }
                }
            } else {
                // New row - generate tempId and save
                const tempId = 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);

                // Save to localStorage with tempId
                this._saveToLocalStorage({ ...filteredRowData, _rowTempId: tempId });

                // Update row with temp ID
                rowNode.removeAttribute('data-new-row');
                rowNode.setAttribute('data-id', tempId);
                rowNode.setAttribute('data-pending', 'true');
                this.emptyRowExists = false;

            }

            this.renumberRows();

            if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                setTimeout(() => this.addEmptyRow(), 0);
            }
            return;
        }

        // ========== DIRECT MODE: Save directly to server ==========
        if (!endpointUrl) {
            console.warn('Save endpoint not configured!');
            return;
        }

        // Get temp ID for tracking
        const tempId = rowNode.getAttribute('data-id') || 'temp_' + Date.now();

        // Consistent payload structure (same as preSavePayload but for server)
        const payload = {
            operation: 'insert',
            data: {
                row: filteredRowData
            },
            meta: {
                timestamp: Date.now(),
                hasErrors: false,
                validationErrors: [],
                rowCount: 1,
                tempId: tempId,
                fieldCount: Object.keys(filteredRowData).length,
                rowIndex: rowIndex
            }
        };

        fetch(endpointUrl, {
            method: 'POST',
            headers: {
                ...this.getExtraHeaders(),
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        })
            .then(async res => {
                if (!res.ok) {
                    const txt = await res.text();
                    throw new Error('Server error: ' + txt);
                }
                return res.json();
            })
            .then(data => {
                if (typeof saveConf.postSave === 'function') saveConf.postSave(data);

                // Extract response data for server-field auto update
                const responseData = data.data || data;
                const newId = responseData.id;

                if (newId) {
                    rowNode.removeAttribute('data-new-row');
                    rowNode.setAttribute('data-id', newId);
                    this.emptyRowExists = false;

                    // SERVER-FIELD AUTO UPDATE: Update readonly cells from server response
                    this._updateRowFromServerResponse(rowNode, responseData);

                    this.renumberRows();

                    // DON'T FORCE call addEmptyRow
                    // Let the last draw event add it (automatically controlled)
                    // But if needed can use the following patch:
                    if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                        setTimeout(() => this.addEmptyRow(), 0);
                    }
                } else {
                    throw new Error('No ID returned from server');
                }
            })
            .catch(err => {
                console.error('[DataTable] Failed to save row:', err);
            });
    }

    /**
     * Save multiple new rows to server in a single API call (Direct Mode)
     * Used for paste operations on multiple new rows
     * @param {Array<HTMLElement>} rowNodes - Array of TR elements to save
     */
    saveMultipleRows(rowNodes) {
        if (!rowNodes || rowNodes.length === 0) {
            return;
        }

        // If only 1 row, use single save
        if (rowNodes.length === 1) {
            return this.saveNewRow(rowNodes[0]);
        }

        let saveConf = this.endpoints.save || {};
        let endpointUrl = saveConf.endpoint || '';

        if (!endpointUrl) {
            console.warn('Save endpoint not configured!');
            return;
        }

        // Collect all row data
        const rowsData = rowNodes.map(rowNode => {
            const rowData = {};
            Array.from(rowNode.children).forEach((td, idx) => {
                const header = this.table.column(idx).header();
                const fieldName = header ? header.getAttribute('data-name') : null;
                if (fieldName) {
                    rowData[fieldName] = this._getCellValue(td, idx);
                }
            });

            // Filter by allowedFields
            let filteredRowData = rowData;
            if (this.allowedFields && Array.isArray(this.allowedFields)) {
                filteredRowData = {};
                this.allowedFields.forEach(f => {
                    if (rowData[f] !== undefined) filteredRowData[f] = rowData[f];
                });
            }

            // Get temp ID for tracking
            const tempId = rowNode.getAttribute('data-id') || 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
            rowNode.setAttribute('data-temp-id', tempId);

            return {
                _tempId: tempId,
                ...filteredRowData
            };
        });

        // Calculate total field count
        const totalFieldCount = rowsData.reduce((sum, row) => {
            const fieldKeys = Object.keys(row).filter(k => k !== '_tempId');
            return sum + fieldKeys.length;
        }, 0);

        // Consistent payload structure
        const payload = {
            operation: 'insert_batch',
            data: {
                rows: rowsData
            },
            meta: {
                timestamp: Date.now(),
                hasErrors: false,
                validationErrors: [],
                rowCount: rowsData.length,
                fieldCount: totalFieldCount
            }
        };

        // preSave callback with consistent payload
        if (typeof saveConf.preSave === 'function') {
            if (saveConf.preSave(payload) === false) {
                return;
            }
        }


        fetch(endpointUrl, {
            method: 'POST',
            headers: {
                ...this.getExtraHeaders(),
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        })
            .then(async res => {
                if (!res.ok) {
                    const txt = await res.text();
                    throw new Error('Server error: ' + txt);
                }
                return res.json();
            })
            .then(data => {
                if (typeof saveConf.postSave === 'function') saveConf.postSave(data);

                // Extract IDs from response
                // Expected: { data: { inserted: [{ tempId, id, ...serverFields }, ...] } }
                const insertedRows = data.data?.inserted || data.inserted || [];

                // Build map of tempId to response data for server-field auto update
                const tempIdToResponseData = {};
                insertedRows.forEach(item => {
                    if (item.tempId) {
                        tempIdToResponseData[item.tempId] = item;
                    }
                });

                // Update row nodes with new IDs and server fields
                rowNodes.forEach(rowNode => {
                    const tempId = rowNode.getAttribute('data-temp-id');
                    const responseData = tempIdToResponseData[tempId];

                    if (responseData && responseData.id) {
                        rowNode.removeAttribute('data-new-row');
                        rowNode.removeAttribute('data-temp-id');
                        rowNode.setAttribute('data-id', responseData.id);

                        // SERVER-FIELD AUTO UPDATE: Update readonly cells from server response
                        this._updateRowFromServerResponse(rowNode, responseData);

                    }
                });

                this.emptyRowExists = false;
                this.renumberRows();

                // Check if should add empty row
                if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                    setTimeout(() => this.addEmptyRow(), 0);
                }

            })
            .catch(err => {
                console.error('[DataTable] Batch insert failed:', err);
            });
    }

    /**
     * Add empty row at bottom
     */
    addEmptyRow(refRowIndex = 0) {
        // Block if allowAddEmptyRow is false (fixed row count mode)
        if (!this.allowAddEmptyRow) {
            return;
        }

        // Block if insert is in progress
        if (this._isInsertingRow) {
            console.log('Blocked addEmptyRow: insert in progress');
            return;
        }

        // ANTISPAM LOOP PATCH
        if (this.isAddingEmptyRow) return;
        this.isAddingEmptyRow = true;

        // Ensure only on last page & only one empty row in data, not counting DOM
        if (!this.isOnLastPage()) {
            this.isAddingEmptyRow = false;
            return;
        }
        if (this.hasEmptyRow()) {
            this.isAddingEmptyRow = false;
            return;
        }

        // --- ADD ROW ---
        let columnCount = this.table.columns().count();
        let newRowData = Array(columnCount).fill('');
        if (this.enableColumnNumber) {
            const currentDataCount = this.table
                .rows((idx, data, node) => !node || node.getAttribute('data-new-row') !== 'true')
                .count();
            newRowData[0] = currentDataCount + 1;
        }
        this.table.row.add(newRowData).draw(false); // DON'T immediately change page!

        setTimeout(() => {
            let newRowIndex = this.table.rows().count() - 1;
            let newRowNode = this.table.row(newRowIndex).node();
            if (newRowNode) {
                newRowNode.setAttribute('data-new-row', 'true');
                newRowNode.setAttribute('data-id', 'new');
                let firstCell = newRowNode.cells[0];
                if (firstCell) firstCell.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                setTimeout(() => this.table.cell(newRowIndex, 0).focus(), 0);
            }
            this.emptyRowExists = true;
            this.isAddingEmptyRow = false; // <------ FLAG RESET
        }, 0);
    }

    /**
     * XSS prevention on cell output
     */
    escapeHtml(str) {
        if (typeof str !== 'string') return str;
        return str.replace(/[&<>"']/g, function (m) {
            return ({
                '&': '&amp;',
                '<': '&lt;',
                '>': '&gt;',
                '"': '&quot;',
                "'": '&#39;'
            })[m];
        });
    }

    renumberRows() {
        // Renumber all rows based on visual order in DOM
        // For paging, start from current page offset
        const tbody = this.table.table().body();
        const rows = tbody.querySelectorAll('tr');

        // Get starting row number based on page offset
        let pageInfo = this.table.page.info();
        let startNumber = pageInfo.start + 1; // page.info().start is 0-indexed

        rows.forEach(rowNode => {
            const cellNo = rowNode.cells[0];
            if (cellNo) {
                cellNo.textContent = startNumber++;
            }
        });
    }

    isOnLastPage() {
        let info = this.table.page.info();
        // Handle empty table (pages = 0)
        if (info.pages === 0) return true;
        return this.table.page() === (info.pages - 1);
    }

    hasEmptyRow() {
        return this.table
            .rows((idx, data, node) => node && node.getAttribute('data-new-row') === 'true')
            .count() > 0;
    }

    isShouldAddEmptyRow() {
        const info = this.table.page.info();
        return info.pages <= 1 || this.table.page() === (info.pages - 1);
    }

    getExtraHeaders() {
        return this.extraHeaders || {};
    }

    // ==================== CONTEXT MENU METHODS ====================

    /**
     * Get clipboard row count from clipboard text
     * @returns {Promise<number>} Number of rows in clipboard
     */
    async _getClipboardRowCount() {
        console.log('[Clipboard] Reading clipboard...');
        try {
            const text = await navigator.clipboard.readText();
            console.log('[Clipboard] Raw text:', text.substring(0, 100) + (text.length > 100 ? '...' : ''));
            if (!text || text.trim() === '') {
                console.log('[Clipboard] Empty clipboard');
                return 0;
            }
            const lines = text.split('\n').filter(line => line.trim() !== '');
            console.log('[Clipboard] Detected', lines.length, 'rows');
            return lines.length;
        } catch (err) {
            console.error('[Clipboard] Failed to read:', err);
            return 0;
        }
    }

    /**
     * Build and show context menu with async clipboard detection
     */
    async _buildContextMenu(e, td) {
        // Remove old menu if exists
        const oldMenu = document.getElementById('custom-contextmenu');
        if (oldMenu) oldMenu.remove();

        const rowNode = td.closest('tr');
        const cellIndex = Array.from(rowNode.children).indexOf(td);

        // Check if this is No. column (first column if enableColumnNumber is active)
        const isNoColumn = this.enableColumnNumber && cellIndex === 0;

        // Check if there's multi-cell selection
        const hasMultiSelection = this._selectedCells && this._selectedCells.length > 1;

        console.log('[ContextMenu] Mode detection:', { isNoColumn, hasMultiSelection, selectedRowsLength: this._selectedRows?.length });

        // Row highlight ONLY if right-click on No. column
        if (isNoColumn && !hasMultiSelection) {
            this._highlightRow(rowNode);
        }

        // Get row info
        const rowId = rowNode ? rowNode.getAttribute('data-id') : null;
        const isPendingRow = rowNode && (
            rowNode.getAttribute('data-pending') === 'true' ||
            rowNode.getAttribute('data-new-row') === 'true' ||
            (rowId && (rowId.startsWith('temp_') || rowId === 'new'))
        );

        // Read clipboard to get actual row count (works for Excel copy too)
        // Always check clipboard, not just when _hasClipboardData is true
        console.log('[ContextMenu] Reading clipboard to detect content...');
        const clipboardRowCount = await this._getClipboardRowCount();
        const hasClipboardData = clipboardRowCount > 0;
        console.log('[ContextMenu] Clipboard status:', { hasClipboardData, clipboardRowCount });

        // Create menu div element
        let menu = document.createElement('div');
        menu.id = 'custom-contextmenu';
        menu.className = 'dt-context-menu';
        menu.style.position = 'absolute';
        menu.style.zIndex = 9999;
        menu.style.left = `${e.pageX}px`;
        menu.style.top = `${e.pageY}px`;
        menu.style.background = '#fff';
        menu.style.border = '1px solid #e2e8f0';
        menu.style.padding = '4px 0';
        menu.style.boxShadow = '0 4px 16px rgba(0,0,0,0.12)';
        menu.style.minWidth = '160px';
        menu.style.borderRadius = '8px';
        menu.style.overflow = 'hidden';

        // Build menu options based on mode
        const opts = [];

        // ============ MULTI-ROW SELECTION MODE ============
        if (this._selectedRows.length > 1) {
            const rowCount = this._selectedRows.length;
            console.log('[ContextMenu] MULTI-ROW MODE:', rowCount, 'rows selected');

            opts.push({
                label: this.lang.copyRows.replace('{n}', rowCount),
                action: () => this._copyMultipleRows()
            });

            // Only show Paste Rows if clipboard has data
            if (hasClipboardData) {
                console.log('[ContextMenu] MULTI-ROW MODE - Clipboard has', clipboardRowCount, 'rows');
                opts.push({
                    label: this.lang.pasteRows.replace('{n}', clipboardRowCount),
                    action: () => this._pasteMultipleRows()
                });
            }

            // Clear rows (reset content, keep row) - no danger styling
            opts.push({
                label: this.lang.clearRows.replace('{n}', rowCount),
                action: () => this._clearMultipleRows()
            });

            opts.push({ label: '---' });

            // Delete multiple rows - only if allowDeleteRow is true
            if (this.allowDeleteRow) {
                opts.push({
                    label: this.lang.deleteRows.replace('{n}', rowCount),
                    action: () => this._deleteMultipleRows(),
                    danger: true
                });
            }
        }
        // ============ MULTI-CELL SELECTION MODE ============
        else if (hasMultiSelection) {
            opts.push({
                label: 'Copy',
                action: () => this._copyCellsToClipboard()
            });

            // Show Paste if clipboard has data
            if (hasClipboardData) {
                opts.push({
                    label: 'Paste',
                    action: () => this._pasteToSelectedCells()
                });
            }

            opts.push({ label: '---' });

            opts.push({
                label: 'Delete',
                action: () => this._deleteSelectedCells(),
                danger: true
            });
        }
        // ============ NO. COLUMN MODE (full row menu) ============
        else if (isNoColumn) {
            // Copy Row
            opts.push({
                label: this.lang.copyRow,
                action: () => this._copyRowToClipboard(rowNode)
            });

            // Clear Row (clear data without deleting row)
            opts.push({
                label: this.lang.clearRow,
                action: () => this._clearRow(rowNode)
            });

            // Paste Row - show for any row when clipboard has data
            if (hasClipboardData) {
                opts.push({
                    label: this.lang.pasteRows.replace('{n}', clipboardRowCount),
                    action: () => this._pasteRowFromClipboard(rowNode)
                });
            }

            // Insert Row Above/Below - only for emptyTable mode, unsaved rows, and when adding rows is allowed
            if (this.emptyTable.enabled && isPendingRow && this.allowAddEmptyRow) {
                const rowHasContent = this._rowHasContent(rowNode);
                const rowIdx = this.table.row(rowNode).index();
                const nextRowNode = this.table.row(rowIdx + 1).node();
                const nextRowIsPending = nextRowNode ? (
                    nextRowNode.getAttribute('data-pending') === 'true' ||
                    nextRowNode.getAttribute('data-id') === 'new' ||
                    nextRowNode.getAttribute('data-new-row') === 'true'
                ) : false;

                // Separator before insert options
                opts.push({ label: '---' });

                // Insert Row Above: row must have data and not saved yet
                if (rowHasContent) {
                    opts.push({
                        label: this.lang.insertAbove,
                        action: () => this._insertRowAt(rowNode, 'above')
                    });
                }

                // Insert Row Below: row must have data AND row below must have data (not empty)
                const nextRowHasContent = nextRowNode ? this._rowHasContent(nextRowNode) : false;
                if (rowHasContent && nextRowHasContent) {
                    opts.push({
                        label: this.lang.insertBelow,
                        action: () => this._insertRowAt(rowNode, 'below')
                    });
                }
            }

            // Separator before readonly
            opts.push({ label: '---' });

            // Readonly Row toggle
            const isRowReadonly = rowNode.getAttribute('data-readonly-row') === 'true';
            opts.push({
                label: isRowReadonly ? this.lang.readonlyRow + ' ✓' : this.lang.readonlyRow,
                action: () => this._toggleRowReadonly(rowNode)
            });

            // Separator before danger
            opts.push({ label: '---' });

            // Delete Row - only if allowDeleteRow is true
            if (this.allowDeleteRow) {
                opts.push({
                    label: this.lang.deleteRow,
                    action: () => this._deleteRow(rowNode, rowId),
                    danger: true
                });
            }
        }
        // ============ SINGLE CELL MODE (cell menu only) ============
        else {
            // Get cell reference for copy/delete
            const cell = this.table.cell(td);

            // Remove existing focus and highlight/focus the clicked cell
            const tableNode = this.table.table().node();

            // Clear ALL focus-related classes from all cells
            tableNode.querySelectorAll('td.focus, td.dt-cell-selected').forEach(c => {
                c.classList.remove('focus', 'dt-cell-selected');
            });
            this._clearCellSelectionHighlight();
            this._removeRowHighlight();
            this._removeFillHandle();

            // Blur KeyTable internal focus
            if (this.table.cell.blur) {
                this.table.cell.blur();
            }

            // Add focus to clicked cell
            td.classList.add('focus');
            td.classList.add('dt-cell-range-selected');
            td.classList.add('dt-range-top', 'dt-range-bottom', 'dt-range-left', 'dt-range-right');

            // Update internal state
            const cellIdx = cell.index();
            this.currentCell = cell;
            this._selectionStart = { row: cellIdx.row, col: cellIdx.column };
            this._selectedCells = [{ row: cellIdx.row, col: cellIdx.column }];
            this._updateFillHandle(cellIdx.row, cellIdx.column);

            // Check if cell is readonly
            const isCellReadonly = td.getAttribute('data-readonly') === 'true';
            const isRowReadonlyCell = rowNode.getAttribute('data-readonly-row') === 'true';

            opts.push({
                label: 'Copy',
                action: () => this._copySingleCell(td)
            });

            // Paste - only if cell is not readonly AND clipboard has data
            if (!isCellReadonly && !isRowReadonlyCell && hasClipboardData) {
                console.log('[ContextMenu] SINGLE CELL MODE - Clipboard has', clipboardRowCount, 'rows');
                if (clipboardRowCount > 1) {
                    // Show "Paste {n} Rows" option for multi-row clipboard data
                    console.log('[ContextMenu] Showing "Paste', clipboardRowCount, 'Rows" option');
                    opts.push({
                        label: this.lang.pasteRows.replace('{n}', clipboardRowCount),
                        action: () => this._pasteMultipleRowsStartingFromCell(td, cellIdx)
                    });
                } else {
                    // Show regular "Paste" for single cell/row
                    opts.push({
                        label: this.lang.paste,
                        action: () => this._pasteSingleCell(td, cellIdx)
                    });
                }
            }

            // Readonly Cell toggle
            opts.push({
                label: isCellReadonly ? 'Readonly ✓' : 'Readonly',
                action: () => this._toggleCellReadonly(td)
            });

            // Separator before danger
            opts.push({ label: '---' });

            // Delete - only if not readonly (always show but disable if readonly)
            opts.push({
                label: 'Clear',
                action: () => this._clearCellContent(cell),
                danger: true,
                disabled: isCellReadonly || isRowReadonlyCell
            });
        }

        // Render menu items
        let addSeparatorToNext = false;
        opts.forEach(opt => {
            if (opt.label === '---') {
                // Mark next item to have separator border
                addSeparatorToNext = true;
            } else {
                let item = document.createElement('div');
                item.className = 'dt-context-menu-item';
                if (opt.danger) item.className += ' dt-context-menu-danger';
                if (opt.disabled) item.className += ' dt-context-menu-disabled';
                item.innerText = opt.label;
                item.style.padding = '8px 16px';
                item.style.cursor = opt.disabled ? 'not-allowed' : 'pointer';
                item.style.fontSize = '13px';

                if (addSeparatorToNext) {
                    item.style.borderTop = '1px solid var(--dt-border, #eee)';
                    item.style.marginTop = '4px';
                    item.style.paddingTop = '12px';
                    addSeparatorToNext = false;
                }

                if (opt.danger) {
                    item.style.color = opt.disabled ? '#ccc' : '#dc3545';
                }
                if (opt.disabled) {
                    item.style.color = '#aaa';
                }

                if (!opt.disabled) {
                    item.onmouseover = () => item.style.background = opt.danger ? '#fff5f5' : 'var(--dt-primary-light, #f5f5f5)';
                    item.onmouseout = () => item.style.background = '';
                    item.onclick = () => {
                        opt.action();
                        menu.remove();
                        if (!hasMultiSelection) {
                            this._removeRowHighlight();
                        }
                    };
                }
                menu.appendChild(item);
            }
        });

        document.body.appendChild(menu);

        // Remove menu if click outside
        // Close menu handler
        const closeMenu = () => {
            if (menu.parentNode) menu.remove();
            if (!hasMultiSelection) {
                this._removeRowHighlight();
            }
            document.removeEventListener('mousedown', onMouseDown);
            document.removeEventListener('keydown', onKeyDown);
        };

        // Click outside handler
        const onMouseDown = (event) => {
            if (!menu.contains(event.target)) {
                closeMenu();
            }
        };

        // Escape key handler
        const onKeyDown = (event) => {
            if (event.key === 'Escape') {
                event.preventDefault();
                closeMenu();
            }
        };

        document.addEventListener('mousedown', onMouseDown);
        document.addEventListener('keydown', onKeyDown);
    }

    /**
     * Delete selected cells content (clear content, not delete rows)
     */
    _deleteSelectedCells() {
        if (!this._selectedCells || this._selectedCells.length === 0) return;

        // Track which rows were affected
        const affectedRows = new Set();

        this._selectedCells.forEach(cellCoord => {
            const cell = this.table.cell(cellCoord.row, cellCoord.col);
            if (cell) {
                const cellNode = cell.node();
                const rowNode = cellNode.closest('tr');
                if (rowNode) {
                    affectedRows.add(rowNode);
                }
                cell.data('').draw(false);
            }
        });

        // Check if any affected pending row is now empty
        affectedRows.forEach(rowNode => {
            this._checkAndResetEmptyPendingRow(rowNode);
        });

    }

    /**
     * Toggle readonly state for entire row
     */
    _toggleRowReadonly(rowNode) {
        if (!rowNode) return;

        const isReadonly = rowNode.getAttribute('data-readonly-row') === 'true';

        if (isReadonly) {
            rowNode.removeAttribute('data-readonly-row');
            // Remove readonly from all cells in row except No. column
            const cells = rowNode.querySelectorAll('td');
            const startIdx = this.enableColumnNumber ? 1 : 0;
            for (let i = startIdx; i < cells.length; i++) {
                cells[i].removeAttribute('data-readonly');
            }
        } else {
            rowNode.setAttribute('data-readonly-row', 'true');
            // Set readonly for all cells in row except No. column
            const cells = rowNode.querySelectorAll('td');
            const startIdx = this.enableColumnNumber ? 1 : 0;
            for (let i = startIdx; i < cells.length; i++) {
                cells[i].setAttribute('data-readonly', 'true');
            }
        }
    }

    /**
     * Toggle readonly state for single cell
     */
    _toggleCellReadonly(cellNode) {
        if (!cellNode) return;

        const isReadonly = cellNode.getAttribute('data-readonly') === 'true';

        if (isReadonly) {
            cellNode.removeAttribute('data-readonly');
        } else {
            cellNode.setAttribute('data-readonly', 'true');
        }
    }

    /**
     * Check if a pending row is now empty and reset its status
     * Removes from localStorage and resets to new row
     */
    _checkAndResetEmptyPendingRow(rowNode) {
        if (!rowNode) return;

        const rowId = rowNode.getAttribute('data-id');
        const hasPendingAttr = rowNode.getAttribute('data-pending') === 'true';
        const isInPendingData = this._pendingData.some(row => row._rowTempId === rowId);
        const isPending = hasPendingAttr || isInPendingData;

        // Only check pending rows
        if (!isPending) return;

        // Check if row is now empty
        if (!this._rowHasContent(rowNode)) {

            // Remove from localStorage
            if (rowId && rowId.startsWith('temp_')) {
                this._removeFromLocalStorage(rowId);
            }

            // Reset row attributes
            rowNode.setAttribute('data-id', 'new');
            rowNode.setAttribute('data-new-row', 'true');
            rowNode.removeAttribute('data-pending');

            // Update badge count
            this._updateBatchCount();
        }
    }

    /**
     * Remove entry from localStorage by temp ID
     */
    _removeFromLocalStorage(tempId) {
        const entryIndex = this._pendingData.findIndex(entry => entry._rowTempId === tempId);
        if (entryIndex !== -1) {
            this._pendingData.splice(entryIndex, 1);
            try {
                localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
            } catch (e) {
                console.error('Error removing from localStorage:', e);
            }
        }
    }

    // ==================== CELL SELECTION METHODS ====================

    /**
     * Update visual selection for multi-cell
     * Highlights all cells between _selectionStart and _selectionEnd
     */
    _updateCellSelection() {
        // Clear existing selection highlight
        this._clearCellSelectionHighlight();

        if (!this._selectionStart) return;

        const start = this._selectionStart;
        const end = this._selectionEnd || this._selectionStart;

        // Calculate range
        const minRow = Math.min(start.row, end.row);
        const maxRow = Math.max(start.row, end.row);
        const minCol = Math.min(start.col, end.col);
        const maxCol = Math.max(start.col, end.col);

        // Highlight all cells in range
        this._selectedCells = [];
        for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
                const cellNode = this.table.cell(r, c).node();
                if (cellNode) {
                    cellNode.classList.add('dt-cell-range-selected');
                    this._selectedCells.push({ row: r, col: c });

                    // Add edge border classes for Google Sheets style
                    if (r === minRow) cellNode.classList.add('dt-range-top');
                    if (r === maxRow) cellNode.classList.add('dt-range-bottom');
                    if (c === minCol) cellNode.classList.add('dt-range-left');
                    if (c === maxCol) cellNode.classList.add('dt-range-right');
                }
            }
        }

        // Update fill handle position
        this._updateFillHandle(maxRow, maxCol);

    }

    // ==================== FILL HANDLE METHODS ====================

    /**
     * Create and update fill handle (blue dot at bottom-right)
     */
    _updateFillHandle(row, col) {
        // Remove existing fill handle
        this._removeFillHandle();

        const cellNode = this.table.cell(row, col).node();
        if (!cellNode) return;

        // Ensure cell has position relative for absolute positioning
        cellNode.style.position = 'relative';

        // Create fill handle element
        const handle = document.createElement('div');
        handle.className = 'dt-fill-handle';

        // Store reference
        this._fillHandleElement = handle;
        this._fillHandleCell = cellNode;

        // Append to cell (position absolute in CSS will handle placement)
        cellNode.appendChild(handle);

        // Add mousedown event for fill drag
        this._addEvent(handle, 'mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this._startFillDrag(e);
        });
    }

    /**
     * Remove fill handle from DOM
     */
    _removeFillHandle() {
        if (this._fillHandleElement) {
            this._fillHandleElement.remove();
            this._fillHandleElement = null;
        }
    }

    /**
     * Start fill drag operation
     */
    _startFillDrag(e) {
        if (!this._selectedCells || this._selectedCells.length === 0) return;

        this._isFillDragging = true;
        this._fillStartCells = [...this._selectedCells];

        // Get original selection bounds
        const rows = this._fillStartCells.map(c => c.row);
        const cols = this._fillStartCells.map(c => c.col);
        this._fillOriginalBounds = {
            minRow: Math.min(...rows),
            maxRow: Math.max(...rows),
            minCol: Math.min(...cols),
            maxCol: Math.max(...cols)
        };

        // Collect source data
        this._fillSourceData = [];
        for (let r = this._fillOriginalBounds.minRow; r <= this._fillOriginalBounds.maxRow; r++) {
            const rowData = [];
            for (let c = this._fillOriginalBounds.minCol; c <= this._fillOriginalBounds.maxCol; c++) {
                const cellNode = this.table.cell(r, c).node();
                // Use centralized helper to get clean value (handles select, checkbox, formatted cells)
                rowData.push(cellNode ? this._getCellRawValue(cellNode, c) : '');
            }
            this._fillSourceData.push(rowData);
        }

        // Add document-level listeners for fill drag
        document.addEventListener('mousemove', this._onFillDragMove);
        document.addEventListener('mouseup', this._onFillDragEnd);
    }

    /**
     * Handle fill drag move - update preview
     */
    _handleFillDragMove(e) {
        if (!this._isFillDragging) return;

        const tableNode = this.table.table().node();
        const td = document.elementFromPoint(e.clientX, e.clientY);
        if (!td || !td.closest || !tableNode.contains(td)) return;

        const cell = td.closest('td');
        if (!cell) return;

        const rowNode = cell.closest('tr');
        if (!rowNode) return;

        const rowIdx = this.table.row(rowNode).index();
        const colIdx = Array.from(rowNode.children).indexOf(cell);

        // Clear previous fill preview
        const preview = tableNode.querySelectorAll('.dt-cell-fill-preview');
        preview.forEach(c => c.classList.remove('dt-cell-fill-preview'));

        this._fillTargetEnd = { row: rowIdx, col: colIdx };

        // Show fill preview - only for cells outside original selection
        const bounds = this._fillOriginalBounds;

        // Determine fill direction (vertical or horizontal)
        const isVertical = colIdx >= bounds.minCol && colIdx <= bounds.maxCol;
        const isHorizontal = rowIdx >= bounds.minRow && rowIdx <= bounds.maxRow;

        if (isVertical && rowIdx > bounds.maxRow) {
            // Fill down
            for (let r = bounds.maxRow + 1; r <= rowIdx; r++) {
                for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
                    const cellNode = this.table.cell(r, c).node();
                    if (cellNode) cellNode.classList.add('dt-cell-fill-preview');
                }
            }
        } else if (isVertical && rowIdx < bounds.minRow) {
            // Fill up
            for (let r = rowIdx; r < bounds.minRow; r++) {
                for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
                    const cellNode = this.table.cell(r, c).node();
                    if (cellNode) cellNode.classList.add('dt-cell-fill-preview');
                }
            }
        } else if (isHorizontal && colIdx > bounds.maxCol) {
            // Fill right
            for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
                for (let c = bounds.maxCol + 1; c <= colIdx; c++) {
                    const cellNode = this.table.cell(r, c).node();
                    if (cellNode) cellNode.classList.add('dt-cell-fill-preview');
                }
            }
        } else if (isHorizontal && colIdx < bounds.minCol) {
            // Fill left
            for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
                for (let c = colIdx; c < bounds.minCol; c++) {
                    const cellNode = this.table.cell(r, c).node();
                    if (cellNode) cellNode.classList.add('dt-cell-fill-preview');
                }
            }
        }
    }

    /**
     * Handle fill drag end - apply fill
     */
    _handleFillDragEnd(e) {
        if (!this._isFillDragging) return;

        // Remove listeners
        document.removeEventListener('mousemove', this._onFillDragMove);
        document.removeEventListener('mouseup', this._onFillDragEnd);

        // Apply fill if we have a target
        if (this._fillTargetEnd && this._fillOriginalBounds) {
            this._applyFill();
        }

        // Cleanup
        this._isFillDragging = false;
        this._fillTargetEnd = null;

        // Clear fill preview
        const tableNode = this.table.table().node();
        const preview = tableNode.querySelectorAll('.dt-cell-fill-preview');
        preview.forEach(c => c.classList.remove('dt-cell-fill-preview'));
    }

    /**
     * Apply fill operation - copy source data to target cells
     * Uses batch API call for efficiency (1 call for all affected rows)
     */
    _applyFill() {
        const bounds = this._fillOriginalBounds;
        const target = this._fillTargetEnd;
        const sourceData = this._fillSourceData;

        if (!bounds || !target || !sourceData.length) return;

        const sourceRowCount = sourceData.length;
        const sourceColCount = sourceData[0].length;

        // Collect all cells to update (for batch API)
        const cellsToUpdate = [];

        // Determine fill direction and range
        const isVertical = target.col >= bounds.minCol && target.col <= bounds.maxCol;
        const isHorizontal = target.row >= bounds.minRow && target.row <= bounds.maxRow;

        if (isVertical && target.row > bounds.maxRow) {
            // Fill down
            for (let r = bounds.maxRow + 1; r <= target.row; r++) {
                const sourceRowIdx = (r - bounds.maxRow - 1) % sourceRowCount;
                for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
                    const sourceColIdx = c - bounds.minCol;
                    const value = sourceData[sourceRowIdx][sourceColIdx];
                    cellsToUpdate.push({ row: r, col: c, value });
                }
            }
        } else if (isVertical && target.row < bounds.minRow) {
            // Fill up
            for (let r = target.row; r < bounds.minRow; r++) {
                const sourceRowIdx = (bounds.minRow - r - 1) % sourceRowCount;
                for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
                    const sourceColIdx = c - bounds.minCol;
                    const value = sourceData[sourceRowCount - 1 - sourceRowIdx][sourceColIdx];
                    cellsToUpdate.push({ row: r, col: c, value });
                }
            }
        } else if (isHorizontal && target.col > bounds.maxCol) {
            // Fill right
            for (let c = bounds.maxCol + 1; c <= target.col; c++) {
                const sourceColIdx = (c - bounds.maxCol - 1) % sourceColCount;
                for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
                    const sourceRowIdx = r - bounds.minRow;
                    const value = sourceData[sourceRowIdx][sourceColIdx];
                    cellsToUpdate.push({ row: r, col: c, value });
                }
            }
        } else if (isHorizontal && target.col < bounds.minCol) {
            // Fill left
            for (let c = target.col; c < bounds.minCol; c++) {
                const sourceColIdx = (bounds.minCol - c - 1) % sourceColCount;
                for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
                    const sourceRowIdx = r - bounds.minRow;
                    const value = sourceData[sourceRowIdx][sourceColCount - 1 - sourceColIdx];
                    cellsToUpdate.push({ row: r, col: c, value });
                }
            }
        }

        // Apply values to DOM and collect data for API
        const rowsData = new Map(); // rowId -> { id, fields: {} }

        cellsToUpdate.forEach(({ row, col, value }) => {
            const cellNode = this.table.cell(row, col).node();
            if (!cellNode) return;

            // Update DOM using centralized helper (handles select, checkbox, formatted cells properly)
            this._setCellValueWithFormatting(cellNode, value, col, true);

            // Get row info
            const rowNode = cellNode.closest('tr');
            const rowId = rowNode ? rowNode.getAttribute('data-id') : null;
            const th = this.table.column(col).header();
            const fieldName = th ? th.getAttribute('data-name') : null;

            if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
                // BATCH MODE: Mark as pending AND track change
                if (rowNode) {
                    rowNode.setAttribute('data-pending', 'true');
                }
                // Track the cell change for batch save
                this._handleCellChange(cellNode, col, value);
            } else {
                // DIRECT MODE: Collect for batch API call
                if (rowId && rowId !== 'new' && !rowId.startsWith('temp_') && fieldName) {
                    if (!rowsData.has(rowId)) {
                        rowsData.set(rowId, { id: rowId, fields: {} });
                    }
                    rowsData.get(rowId).fields[fieldName] = value;
                }
            }
        });

        // DIRECT MODE: Send batch API call
        if (!(this.emptyTable.enabled && this.emptyTable.saveMode === 'batch')) {
            const rowsArray = Array.from(rowsData.values());
            if (rowsArray.length === 1) {
                // Single row - use updateRowBatch
                this.updateRowBatch(rowsArray[0].id, rowsArray[0].fields);
            } else if (rowsArray.length > 1) {
                // Multiple rows - use updateMultipleRows (1 API call!)
                this.updateMultipleRows(rowsArray);
            }
        }

        // Recalculate formulas for all affected rows
        const affectedRows = new Set();
        cellsToUpdate.forEach(({ row }) => {
            const rowNode = this.table.row(row).node();
            if (rowNode) affectedRows.add(rowNode);
        });
        affectedRows.forEach(rowNode => this._recalculateRowFormulas(rowNode));

        // Update footer totals after formula recalculation
        this._updateFooterTotals();

        // Re-render select columns to add arrow styling
        setTimeout(() => this._renderSelectColumns(), 10);

        // Check if any affected rows are new rows that are now complete
        if (!(this.emptyTable.enabled && this.emptyTable.saveMode === 'batch')) {
            affectedRows.forEach(rowNode => {
                const rowId = rowNode.getAttribute('data-id');
                if (rowId === 'new' || rowNode.getAttribute('data-new-row') === 'true') {
                    // Check if row is complete after fill
                    setTimeout(() => {
                        if (this.isRowRequiredFieldsFilled(rowNode)) {
                            Promise.resolve(this.saveNewRow(rowNode))
                                .finally(() => {
                                    if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                                        this.addEmptyRow();
                                    }
                                });
                        }
                    }, 50);
                }
            });
        }

    }

    /**
     * Set cell value helper for fill operation (DOM only, no API call)
     * API calls are now handled by _applyFill in batch
     * @deprecated Use _applyFill which handles batch API calls
     */
    _setCellValue(row, col, value) {
        const cellNode = this.table.cell(row, col).node();
        if (cellNode) {
            cellNode.textContent = value;
            this.table.cell(row, col).data(value);

            // Get row info
            const rowNode = cellNode.closest('tr');
            const rowId = rowNode ? rowNode.getAttribute('data-id') : null;
            const th = this.table.column(col).header();
            const fieldName = th ? th.getAttribute('data-name') : null;

            // Mark as pending if batch mode
            if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
                if (rowNode) {
                    rowNode.setAttribute('data-pending', 'true');
                }
            } else {
                // DIRECT MODE: Update cell via API
                if (rowId && rowId !== 'new' && !rowId.startsWith('temp_') && fieldName) {
                    this.updateCell(fieldName, rowId, value);
                }
            }
        }
    }

    /**
     * Helper to extract value from cell, handling special column types
     * 
     * Behavior based on column type:
     * - checkbox: return 'true' or 'false' from input.checked
     * - select: return text from .dt-select-content span (without arrow)
     * - text/default: return textContent.trim()
     * 
     * @param {HTMLTableCellElement} cellNode - TD Element
     * @param {number} colIdx - Column index
     * @returns {string} Extracted cell value
     * 
     * @example
     * const value = this._getCellValue(td, 3);
     * // For checkbox: 'true' or 'false'
     * // For select: 'Jakarta' (not 'Jakarta▼')
     */
    _getCellValue(cellNode, colIdx) {
        if (!cellNode) return '';

        const th = this.table.column(colIdx).header();
        const dataType = th ? th.getAttribute('data-type') : null;

        // Checkbox column: get checked state
        if (dataType === 'checkbox') {
            const checkbox = cellNode.querySelector('input[type="checkbox"]');
            if (checkbox) {
                return checkbox.checked ? 'true' : 'false';
            }
            // Fallback to data attribute
            if (cellNode.hasAttribute('data-checkbox-value')) {
                return cellNode.getAttribute('data-checkbox-value');
            }
            // Fallback to text content
            const text = cellNode.textContent.trim().toLowerCase();
            return (text === 'true' || text === '1') ? 'true' : 'false';
        }

        // Select column: get text from dt-select-content span (without arrow)
        if (dataType === 'select') {
            const selectContent = cellNode.querySelector('.dt-select-content');
            if (selectContent) {
                return selectContent.textContent.trim();
            }
            // Fallback: strip arrow character from text content
            return cellNode.textContent.trim().replace(/[▼▲]/g, '').trim();
        }

        // Default: get text content
        return cellNode.textContent.trim();
    }

    /**
     * Clear cell selection highlight from DOM
     */
    _clearCellSelectionHighlight() {
        const tableNode = this.table.table().node();

        // Clear range selection classes
        const selected = tableNode.querySelectorAll('.dt-cell-range-selected');
        selected.forEach(cell => {
            cell.classList.remove('dt-cell-range-selected');
            cell.classList.remove('dt-range-top', 'dt-range-bottom', 'dt-range-left', 'dt-range-right');
        });

        // Also clear focus class from all cells (manually added by contextmenu)
        tableNode.querySelectorAll('td.focus').forEach(cell => {
            cell.classList.remove('focus');
        });

        // Clear fill preview
        const preview = tableNode.querySelectorAll('.dt-cell-fill-preview');
        preview.forEach(cell => cell.classList.remove('dt-cell-fill-preview'));
    }

    /**
     * Clear cell selection completely
     */
    _clearCellSelection() {
        this._clearCellSelectionHighlight();
        this._removeFillHandle();
        this._selectedCells = [];
        this._selectionEnd = null;
    }

    // ==================== CLIPBOARD METHODS (Copy/Paste) ====================

    /**
     * Copy multiple selected cells to clipboard
     * Format: tab-separated per row, newline between rows
     */
    _copyCellsToClipboard() {
        if (!this._selectedCells || this._selectedCells.length === 0) return false;

        // Group cells by row
        const rows = {};
        this._selectedCells.forEach(cell => {
            if (!rows[cell.row]) rows[cell.row] = [];
            rows[cell.row].push(cell.col);
        });

        // Sort columns within each row
        Object.keys(rows).forEach(r => {
            rows[r].sort((a, b) => a - b);
        });

        // Build text
        const lines = [];
        const sortedRows = Object.keys(rows).map(Number).sort((a, b) => a - b);

        sortedRows.forEach(r => {
            const values = rows[r].map(c => {
                const cellNode = this.table.cell(r, c).node();
                return cellNode ? this._getCellValue(cellNode, c) : '';
            });
            lines.push(values.join('\t'));
        });

        const text = lines.join('\n');
        navigator.clipboard.writeText(text).then(() => {
            this._hasClipboardData = true;
        }).catch(err => {
            console.error('[DataTable] Failed to copy to clipboard:', err);
        });

        return true;
    }

    /**
     * Highlight row (for No click or right-click)
     */
    _highlightRow(rowNode) {
        this._removeRowHighlight();
        // Also remove cell highlight when row is highlighted
        this._removeHighlight(null);
        // Remove focus styling from KeyTable
        const allFocused = this.table.table().node().querySelectorAll('.focus');
        allFocused.forEach(cell => cell.classList.remove('focus'));
        // Clear cell selection and fill handle
        this._clearCellSelection();
        if (rowNode) {
            rowNode.classList.add('row-highlight');
            this._highlightedRow = rowNode;
        }
    }

    /**
     * Remove row highlight
     */
    _removeRowHighlight() {
        if (this._highlightedRow) {
            this._highlightedRow.classList.remove('row-highlight');
            this._highlightedRow = null;
        }
        const highlighted = document.querySelectorAll('.row-highlight');
        highlighted.forEach(row => row.classList.remove('row-highlight'));
    }

    /**
     * Copy row data to clipboard (tab-separated for Excel)
     */
    _copyRowToClipboard(rowNode) {
        if (!rowNode) return;
        const cells = rowNode.querySelectorAll('td');
        const values = [];
        const startIdx = this.enableColumnNumber ? 1 : 0;
        for (let i = startIdx; i < cells.length; i++) {
            const cell = cells[i];
            // Check for checkbox value
            const checkbox = cell.querySelector('input[type="checkbox"]');
            if (checkbox) {
                values.push(checkbox.checked ? 'true' : 'false');
            } else if (cell.getAttribute('data-checkbox-value')) {
                values.push(cell.getAttribute('data-checkbox-value'));
            } else {
                // Check for select/dropdown content (exclude arrow)
                const selectContent = cell.querySelector('.dt-select-content');
                if (selectContent) {
                    values.push(selectContent.textContent.trim());
                } else {
                    values.push(cell.textContent.trim());
                }
            }
        }
        const text = values.join('\t');
        navigator.clipboard.writeText(text).then(() => {
            this._hasClipboardData = true;
        }).catch(err => {
            console.error('Failed to copy:', err);
        });
    }

    /**
     * Copy multiple selected rows to clipboard
     */
    _copyMultipleRows() {
        if (this._selectedRows.length === 0) return;

        const lines = [];
        const startColIdx = this.enableColumnNumber ? 1 : 0;

        // Sort rows by index
        const sortedRows = [...this._selectedRows].sort((a, b) => a - b);

        sortedRows.forEach(rowIdx => {
            const rowNode = this.table.row(rowIdx).node();
            if (!rowNode) return;

            const cells = rowNode.querySelectorAll('td');
            const values = [];

            for (let i = startColIdx; i < cells.length; i++) {
                const cell = cells[i];
                const checkbox = cell.querySelector('input[type="checkbox"]');
                if (checkbox) {
                    values.push(checkbox.checked ? 'true' : 'false');
                } else if (cell.getAttribute('data-checkbox-value')) {
                    values.push(cell.getAttribute('data-checkbox-value'));
                } else {
                    // Check for select/dropdown content (exclude arrow)
                    const selectContent = cell.querySelector('.dt-select-content');
                    if (selectContent) {
                        values.push(selectContent.textContent.trim());
                    } else {
                        values.push(cell.textContent.trim());
                    }
                }
            }
            lines.push(values.join('\t'));
        });

        const text = lines.join('\n');
        navigator.clipboard.writeText(text).then(() => {
            this._hasClipboardData = true;
            this._clipboardRowCount = lines.length;
            console.log('[CopyMultipleRows] Copied', lines.length, 'rows to clipboard');
            // Keep row selection intact after copy (user may want to delete or do other operations)
        }).catch(err => {
            console.error('Failed to copy:', err);
        });
    }

    /**
     * Clear multiple selected rows
     * Delegates to _clearRow() for each row to properly handle
     * localStorage tracking in batch mode and DOM clearing
     */
    _clearMultipleRows() {
        if (this._selectedRows.length === 0) return;

        // Sort rows descending to avoid index shift issues
        const sortedRows = [...this._selectedRows].sort((a, b) => b - a);

        sortedRows.forEach(rowIdx => {
            const rowNode = this.table.row(rowIdx).node();
            if (!rowNode) return;

            // Delegate to _clearRow which handles:
            // - Pending rows: remove from _pendingData + localStorage
            // - Batch mode existing rows: track in _deletedData + localStorage
            // - Direct mode: call delete endpoint
            // - DOM clearing via _clearRowDom
            this._clearRow(rowNode);
        });

        // Clear selection state
        this._clearRowSelection();
        this._updateBatchCount();
    }

    /**
     * Delete multiple selected rows
     */
    _deleteMultipleRows() {
        if (this._selectedRows.length === 0) return;

        // Sort rows descending to avoid index shift issues when removing
        const sortedRows = [...this._selectedRows].sort((a, b) => b - a);

        sortedRows.forEach(rowIdx => {
            const rowNode = this.table.row(rowIdx).node();
            if (!rowNode) return;

            const rowId = rowNode.getAttribute('data-id');
            this._deleteRow(rowNode, rowId);
        });

        // Clear selection state
        this._clearRowSelection();
    }

    /**
     * Paste multiple rows from clipboard
     */
    async _pasteMultipleRows() {
        if (this._selectedRows.length === 0) return;

        try {
            const text = await navigator.clipboard.readText();
            if (!text || text.trim() === '') {
                return;
            }

            // Parse clipboard lines
            const lines = text.split('\n').filter(line => line.trim() !== '');
            const sortedRows = [...this._selectedRows].sort((a, b) => a - b);
            const startColIdx = this.enableColumnNumber ? 1 : 0;

            // Collect all row updates for batch processing (Direct Mode)
            const existingRowsToUpdate = [];
            const newRowsToSave = [];

            // Paste each line - paste to consecutive rows starting from first selected row
            const firstSelectedRowIdx = sortedRows.length > 0 ? sortedRows[0] : this.table.rows().count();


            // Calculate how many rows we need total and how many need to be created
            const existingRowCount = this.table.rows().count();
            const targetEndRowIdx = firstSelectedRowIdx + lines.length - 1;
            let rowsToCreate = Math.max(0, targetEndRowIdx - existingRowCount + 1);

            // If allowAddEmptyRow is false, don't create new rows - trim lines instead
            if (!this.allowAddEmptyRow && rowsToCreate > 0) {
                const maxLines = existingRowCount - firstSelectedRowIdx;
                lines.length = Math.min(lines.length, maxLines);
                rowsToCreate = 0;
            }

            // Create all needed rows first (before pasting)
            if (rowsToCreate > 0) {

                for (let r = 0; r < rowsToCreate; r++) {
                    const columnCount = this.table.columns().count();
                    const newRowData = Array(columnCount).fill('');
                    if (this.enableColumnNumber) {
                        const currentDataCount = this.table.rows().count();
                        newRowData[0] = currentDataCount + 1;
                    }

                    this.table.row.add(newRowData);
                }

                // Draw once after all rows are added
                this.table.draw(false);

                // Set attributes on new rows
                for (let r = 0; r < rowsToCreate; r++) {
                    const newRowIdx = existingRowCount + r;
                    const newRowNode = this.table.row(newRowIdx).node();
                    if (newRowNode) {
                        newRowNode.setAttribute('data-new-row', 'true');
                        newRowNode.setAttribute('data-id', 'new');
                    }
                }
            }

            // Now paste to all consecutive rows starting from firstSelectedRowIdx
            for (let i = 0; i < lines.length; i++) {
                // Simple consecutive row calculation
                const rowIdx = firstSelectedRowIdx + i;
                const rowNode = this.table.row(rowIdx).node();

                if (!rowNode) continue;

                const values = lines[i].split('\t');
                const colCount = this.table.columns().count();

                // Build row data object for batch mode
                const rowData = {};

                // Paste values into cells and collect data
                let valueIdx = 0;
                for (let colIdx = startColIdx; colIdx < colCount && valueIdx < values.length; colIdx++) {
                    const cell = this.table.cell(rowIdx, colIdx);
                    const th = this.table.column(colIdx).header();
                    const fieldName = th ? th.getAttribute('data-name') : null;
                    const dataType = th ? th.getAttribute('data-type') : null;

                    // Skip readonly columns
                    if (dataType === 'readonly') {
                        valueIdx++;
                        continue;
                    }

                    if (cell) {
                        let value = values[valueIdx].trim();

                        const cellNode = cell.node();

                        // Skip readonly cells
                        if (cellNode) {
                            const isCellReadonly = cellNode.getAttribute('data-readonly') === 'true';
                            const isRowReadonly = rowNode.getAttribute('data-readonly-row') === 'true';

                            if (isCellReadonly || isRowReadonly) {
                                valueIdx++;
                                continue;
                            }
                        }

                        // Strip arrow from select column values
                        if (dataType === 'select') {
                            value = value.replace(/[▼▲]/g, '').trim();
                        }

                        // Use direct DOM update instead of cell.data() to avoid triggering events
                        if (cellNode) {
                            this._setCellValueWithFormatting(cellNode, value, colIdx, true);
                        }

                        // Collect field data (for both batch and direct modes)
                        if (fieldName) {
                            rowData[fieldName] = value;
                        }
                    }
                    valueIdx++;
                }

                // Handle batch mode - update _pendingData properly
                if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
                    let rowId = rowNode.getAttribute('data-id');

                    if (rowId === 'new') {
                        // New row - generate temp_id
                        const tempId = 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
                        rowNode.setAttribute('data-id', tempId);
                        rowNode.setAttribute('data-pending', 'true');

                        // Add to pending data
                        rowData._rowTempId = tempId;
                        rowData._timestamp = Date.now();
                        this._pendingData.push(rowData);

                    } else if (rowId && rowId.startsWith('temp_')) {
                        // Existing temp row - UPDATE instead of adding duplicate
                        rowNode.setAttribute('data-pending', 'true');

                        const entryIndex = this._pendingData.findIndex(entry => entry._rowTempId === rowId);

                        if (entryIndex !== -1) {
                            // Update existing entry - REPLACE all field data, keep _rowTempId
                            // Clear old fields and set new ones
                            Object.keys(this._pendingData[entryIndex]).forEach(key => {
                                if (key !== '_rowTempId' && key !== '_timestamp') {
                                    delete this._pendingData[entryIndex][key];
                                }
                            });
                            Object.keys(rowData).forEach(key => {
                                this._pendingData[entryIndex][key] = rowData[key];
                            });
                            this._pendingData[entryIndex]._timestamp = Date.now();
                        } else {
                            // Entry not found, add as new
                            rowData._rowTempId = rowId;
                            rowData._timestamp = Date.now();
                            this._pendingData.push(rowData);
                        }
                    } else if (rowId && rowId !== '' && !rowId.startsWith('temp_')) {
                        // Existing saved row with real ID - mark as edited
                        rowNode.setAttribute('data-edited', 'true');

                        // Update _editedData
                        if (!this._editedData[rowId]) {
                            this._editedData[rowId] = { id: rowId, fields: {}, timestamp: Date.now() };
                        }
                        Object.keys(rowData).forEach(key => {
                            this._editedData[rowId].fields[key] = rowData[key];
                        });
                        this._editedData[rowId].timestamp = Date.now();
                    }
                } else {
                    // DIRECT MODE - collect for batch update
                    const rowId = rowNode.getAttribute('data-id');

                    if (rowId === 'new' && this.isRowRequiredFieldsFilled(rowNode)) {
                        // New row - queue for save
                        newRowsToSave.push(rowNode);
                    } else if (rowId && rowId !== 'new' && !rowId.startsWith('temp_')) {
                        // Existing row from DB - collect for batch update
                        existingRowsToUpdate.push({ id: rowId, fields: rowData });
                    }
                }
            }

            // DIRECT MODE: Process collected rows
            if (!(this.emptyTable.enabled && this.emptyTable.saveMode === 'batch')) {
                // Save new rows using batch insert (1 API call for all)
                if (newRowsToSave.length > 0) {
                    this.saveMultipleRows(newRowsToSave);
                }

                // Update existing rows in single batch call
                if (existingRowsToUpdate.length > 0) {
                    this.updateMultipleRows(existingRowsToUpdate);
                }
            }

            // Save to localStorage AFTER all rows processed (not inside loop) - BATCH MODE ONLY
            if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
                // Only save pending data if there's actually data
                if (this._pendingData.length > 0) {
                    localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
                }
                // Only save edited data if there's actually data
                if (Object.keys(this._editedData).length > 0) {
                    localStorage.setItem(this.emptyTable.storageKey + '_edited', JSON.stringify(this._editedData));
                }

                // Check each pasted row for completion and mark appropriately
                let anyRowComplete = false;
                for (let i = 0; i < lines.length; i++) {
                    const rowIdx = firstSelectedRowIdx + i;
                    const rowNode = this.table.row(rowIdx).node();
                    if (rowNode && this.isRowRequiredFieldsFilled(rowNode)) {
                        rowNode.removeAttribute('data-new-row');
                        anyRowComplete = true;
                    }
                }

                // Add empty row if any row was completed and no empty rows remain
                if (anyRowComplete) {
                    setTimeout(() => {
                        if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                            this.addEmptyRow();
                        } else {
                        }
                    }, 100);
                }
            }

            // Collect all temp_ids before draw for re-application
            const tempIdMap = new Map();
            this._pendingData.forEach(entry => {
                if (entry._rowTempId) {
                    tempIdMap.set(entry._rowTempId, entry);
                }
            });

            // Redraw
            this.table.draw(false);

            // Re-apply data-id attributes after draw (draw may reset DOM)
            const self = this;
            this.table.rows().every(function (rowIdx) {
                const rowNode = this.node();
                if (rowNode) {
                    const currentId = rowNode.getAttribute('data-id');
                    // Only set pending if temp row is actually in pending data
                    if (currentId && currentId.startsWith('temp_') && tempIdMap.has(currentId)) {
                        rowNode.setAttribute('data-pending', 'true');
                    }
                }
            });

            this._renderCheckboxColumns();
            this._renderSelectColumns();
            this._clearRowSelection();
            this._updateBatchCount();

            // Clear clipboard after paste
            await navigator.clipboard.writeText('');
            this._hasClipboardData = false;
            this._clipboardRowCount = 0;

        } catch (err) {
            console.error('Failed to paste:', err);
            alert('Unable to access clipboard. Please allow clipboard permissions.');
        }
    }

    /**
     * Copy single cell to clipboard
     */
    _copySingleCell(td, colIdx = null) {
        if (!td) return;

        // Get column index if not provided
        if (colIdx === null) {
            const cell = this.table.cell(td);
            colIdx = cell ? cell.index().column : 0;
        }

        // Use helper to get proper value (handles checkbox, select, etc.)
        const text = this._getCellValue(td, colIdx);

        navigator.clipboard.writeText(text).then(() => {
            this._hasClipboardData = true;
            this._clipboardRowCount = 1;
        }).catch(err => {
            console.error('Failed to copy:', err);
        });
    }

    /**
     * Paste clipboard content starting from a single cell
     * If clipboard has multi-column data (tab-separated), paste across columns
     * If clipboard has multi-row data (newline-separated), paste across rows
     */
    async _pasteSingleCell(td, cellIdx) {
        if (!td) return;

        try {
            const text = await navigator.clipboard.readText();
            if (!text || text.trim() === '') {
                return;
            }

            const rowNode = td.closest('tr');
            const rowIdx = this.table.row(rowNode).index();
            const startColIndex = cellIdx.column;
            const colCount = this.table.columns().count();
            const rowCount = this.table.rows().count();

            // Parse clipboard: newlines = rows, tabs = columns
            const lines = text.split('\n').filter(line => line !== '');

            // Collect all changes for batch handling
            const changedCells = [];

            for (let lineIdx = 0; lineIdx < lines.length; lineIdx++) {
                const currentRowIdx = rowIdx + lineIdx;
                if (currentRowIdx >= rowCount) break;

                const currentRowNode = this.table.row(currentRowIdx).node();
                if (!currentRowNode) continue;

                const values = lines[lineIdx].split('\t');

                for (let valueIdx = 0; valueIdx < values.length; valueIdx++) {
                    const currentColIdx = startColIndex + valueIdx;
                    if (currentColIdx >= colCount) break;

                    const value = values[valueIdx].trim();
                    const cellNode = this.table.cell(currentRowIdx, currentColIdx).node();

                    if (cellNode) {
                        // Skip readonly cells
                        const isCellReadonly = cellNode.getAttribute('data-readonly') === 'true';
                        const isRowReadonly = currentRowNode.getAttribute('data-readonly-row') === 'true';
                        const th = this.table.column(currentColIdx).header();
                        const fieldType = th ? th.getAttribute('data-type') : null;
                        const isColumnReadonly = fieldType === 'readonly';

                        if (isCellReadonly || isRowReadonly || isColumnReadonly) {
                            continue;
                        }

                        // Use centralized helper to set cell value with formatting
                        this._setCellValueWithFormatting(cellNode, value, currentColIdx, true);

                        // Get clean value after formatting
                        const cleanedValue = this._getCellRawValue(cellNode, currentColIdx);

                        // Collect change for batch processing
                        changedCells.push({
                            rowNode: currentRowNode,
                            colIndex: currentColIdx,
                            value: cleanedValue
                        });
                    }
                }
            }

            // Process all changes (batch or direct mode)
            if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
                // BATCH MODE: Track each cell change
                changedCells.forEach(({ rowNode, colIndex, value }) => {
                    this._handleCellChange(rowNode.querySelector(`td:nth-child(${colIndex + 1})`), colIndex, value);
                });
            } else {
                // DIRECT MODE: Collect by row for batch update
                const rowUpdates = new Map();

                changedCells.forEach(({ rowNode, colIndex, value }) => {
                    const rowId = rowNode.getAttribute('data-id');
                    if (!rowId || rowId === 'new') return;

                    const th = this.table.column(colIndex).header();
                    const fieldName = th ? th.getAttribute('data-name') : null;
                    if (!fieldName) return;

                    if (!rowUpdates.has(rowId)) {
                        rowUpdates.set(rowId, { id: rowId, fields: {} });
                    }
                    rowUpdates.get(rowId).fields[fieldName] = value;
                });

                // Send updates for existing rows
                if (rowUpdates.size === 1) {
                    // Single row - use updateRowBatch
                    const [rowId, data] = [...rowUpdates.entries()][0];
                    this.updateRowBatch(rowId, data.fields);
                } else if (rowUpdates.size > 1) {
                    // Multiple rows - use updateMultipleRows
                    this.updateMultipleRows([...rowUpdates.values()]);
                }

                // DIRECT MODE: Handle NEW ROWS after paste
                // Collect unique new rows that were pasted to
                const newRows = new Set();
                changedCells.forEach(({ rowNode }) => {
                    const rowId = rowNode.getAttribute('data-id');
                    if (rowId === 'new' || rowNode.getAttribute('data-new-row') === 'true') {
                        newRows.add(rowNode);
                    }
                });

                // Check each new row - if complete, save it
                newRows.forEach(rowNode => {
                    if (this.isRowRequiredFieldsFilled(rowNode)) {
                        this.saveNewRow(rowNode);
                    }
                });
            }

            // Recalculate formulas for all affected rows
            const affectedRows = new Set();
            changedCells.forEach(({ rowNode }) => affectedRows.add(rowNode));
            affectedRows.forEach(rowNode => this._recalculateRowFormulas(rowNode));
            // Update footer totals after paste
            this._updateFooterTotals();

            this._updateBatchCount();

            // Re-render checkbox and select columns to restore arrow indicators
            this._renderCheckboxColumns();
            this._renderSelectColumns();

            // Clear clipboard after paste
            await navigator.clipboard.writeText('');
            this._hasClipboardData = false;
            this._clipboardRowCount = 0;

        } catch (err) {
            console.error('Failed to paste:', err);
            alert('Unable to access clipboard. Please allow clipboard permissions.');
        }
    }

    /**
     * Paste clipboard content into multiple selected cells
     * Supports grid paste (tab = column, newline = row)
     * Uses centralized helpers for consistent handling
     */
    async _pasteToSelectedCells() {
        if (!this._selectedCells || this._selectedCells.length === 0) return;

        try {
            const text = await navigator.clipboard.readText();
            if (!text || text.trim() === '') {
                return;
            }

            // Parse clipboard: newlines = rows, tabs = columns
            const lines = text.split('\n').filter(line => line !== '');
            const values = lines.map(line => line.split('\t').map(v => v.trim()));

            // Get bounds of selection
            const minRow = Math.min(...this._selectedCells.map(c => c.row));
            const maxRow = Math.max(...this._selectedCells.map(c => c.row));
            const minCol = Math.min(...this._selectedCells.map(c => c.col));
            const maxCol = Math.max(...this._selectedCells.map(c => c.col));

            // Track affected rows for batch processing
            const affectedRows = new Set();

            // Paste values to selection area using centralized helper
            let valueRowIdx = 0;
            for (let r = minRow; r <= maxRow && valueRowIdx < values.length; r++) {
                affectedRows.add(r);

                let valueColIdx = 0;
                for (let c = minCol; c <= maxCol && valueColIdx < values[valueRowIdx].length; c++) {
                    const value = values[valueRowIdx][valueColIdx];
                    const cellNode = this.table.cell(r, c).node();

                    if (cellNode) {
                        // Skip readonly cells
                        const rowNode = cellNode.closest('tr');
                        const isCellReadonly = cellNode.getAttribute('data-readonly') === 'true';
                        const isRowReadonly = rowNode && rowNode.getAttribute('data-readonly-row') === 'true';
                        const th = this.table.column(c).header();
                        const fieldType = th ? th.getAttribute('data-type') : null;
                        const isColumnReadonly = fieldType === 'readonly' || fieldType === 'formula';

                        if (isCellReadonly || isRowReadonly || isColumnReadonly) {
                            valueColIdx++;
                            continue;
                        }

                        // Use centralized helper to set cell value with formatting
                        this._setCellValueWithFormatting(cellNode, value, c, true);
                    }

                    valueColIdx++;
                }
                valueRowIdx++;
            }

            // Handle batch mode / direct mode for affected rows
            // Collect rows for batch API call (Direct Mode)
            const rowsToUpdate = [];

            affectedRows.forEach(rowIdx => {
                const rowNode = this.table.row(rowIdx).node();
                if (!rowNode) return;

                const rowId = rowNode.getAttribute('data-id');
                const rowData = this._getRowData(rowNode);

                if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
                    // BATCH MODE
                    const isNewRow = rowId === 'new' || !rowId;
                    const isTempRow = rowId && rowId.startsWith('temp_');
                    const isDbRow = rowId && rowId !== 'new' && !rowId.startsWith('temp_');

                    if (isNewRow || isTempRow) {
                        // New/pending row
                        let tempId = rowId;
                        if (isNewRow) {
                            tempId = 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
                            rowNode.setAttribute('data-id', tempId);
                            rowNode.removeAttribute('data-new-row');
                        }
                        rowNode.setAttribute('data-pending', 'true');

                        // Save all fields to pending data
                        Object.keys(rowData).forEach(fieldName => {
                            this._updateLocalStorageEntry(tempId, fieldName, rowData[fieldName], false);
                        });

                        // Check if complete and save
                        if (this.isRowRequiredFieldsFilled(rowNode)) {
                            try {
                                localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
                            } catch (e) {
                                console.error('Error saving to localStorage:', e);
                            }
                        }
                    } else if (isDbRow) {
                        // Existing DB row - track edits
                        rowNode.setAttribute('data-edited', 'true');
                        Object.keys(rowData).forEach(fieldName => {
                            this._trackEditedField(rowId, fieldName, rowData[fieldName]);
                        });
                    }
                } else {
                    // DIRECT MODE: Collect for batch API call
                    if (rowId && rowId !== 'new' && !rowId.startsWith('temp_')) {
                        rowsToUpdate.push({ id: rowId, fields: rowData });
                    }
                }
            });

            // DIRECT MODE: Single batch API call for all affected rows
            if (rowsToUpdate.length === 1) {
                // Only 1 row - use updateRowBatch for single row
                this.updateRowBatch(rowsToUpdate[0].id, rowsToUpdate[0].fields);
            } else if (rowsToUpdate.length > 1) {
                // Multiple rows - use updateMultipleRows for batch
                this.updateMultipleRows(rowsToUpdate);
            }

            this._updateBatchCount();

            // Update footer totals after paste
            this._updateFooterTotals();

            // Redraw and re-render
            this.table.draw(false);
            this._renderCheckboxColumns();
            this._renderSelectColumns();

            // Clear clipboard after paste
            await navigator.clipboard.writeText('');
            this._hasClipboardData = false;
            this._clipboardRowCount = 0;

        } catch (err) {
            console.error('Failed to paste:', err);
            alert('Unable to access clipboard. Please allow clipboard permissions.');
        }
    }

    /**
     * Paste multiple rows starting from a specific cell
     * Called from context menu when clicking on a single cell with multi-row clipboard data
     * @param {HTMLElement} td - Starting cell element
     * @param {Object} cellIdx - Cell index object with row and column
     */
    async _pasteMultipleRowsStartingFromCell(td, cellIdx) {
        if (!td) return;

        try {
            const text = await navigator.clipboard.readText();
            if (!text || text.trim() === '') {
                return;
            }

            const rowNode = td.closest('tr');
            const startRowIdx = cellIdx.row;
            const startColIdx = cellIdx.column;
            const colCount = this.table.columns().count();
            const rowCount = this.table.rows().count();

            // Parse clipboard: newlines = rows, tabs = columns
            const lines = text.split('\n').filter(line => line.trim() !== '');

            // Collect all changes for batch handling
            const changedCells = [];

            for (let lineIdx = 0; lineIdx < lines.length; lineIdx++) {
                const currentRowIdx = startRowIdx + lineIdx;
                if (currentRowIdx >= rowCount) break;

                const currentRowNode = this.table.row(currentRowIdx).node();
                const values = lines[lineIdx].split('\t');

                for (let colOffset = 0; colOffset < values.length; colOffset++) {
                    const currentColIdx = startColIdx + colOffset;
                    if (currentColIdx >= colCount) break;

                    const value = values[colOffset].trim();
                    const targetCell = currentRowNode.cells[currentColIdx];

                    if (targetCell) {
                        changedCells.push({
                            rowNode: currentRowNode,
                            rowIdx: currentRowIdx,
                            cell: targetCell,
                            colIdx: currentColIdx,
                            value: value
                        });
                    }
                }
            }

            // Apply all changes
            for (const { rowNode: rNode, rowIdx, cell, colIdx, value } of changedCells) {
                const rowId = rNode.getAttribute('data-id');
                const header = this.table.column(colIdx).header();
                const fieldName = header?.getAttribute('data-name');
                const fieldType = header?.getAttribute('data-type');

                // Skip readonly cells
                if (cell.getAttribute('data-readonly') === 'true' ||
                    rNode.getAttribute('data-readonly-row') === 'true') {
                    continue;
                }

                // Validate if needed
                const isValid = this._validateField(value, fieldType);
                if (!isValid) {
                    this._handleValidationError(cell, value);
                    continue;
                }

                // Set cell value with formatting
                this._setCellValueWithFormatting(cell, value, colIdx, false);

                // Handle batch vs direct mode
                if (this.emptyTable.enabled && 'batch' === this.emptyTable.saveMode) {
                    if (rowId && rowId.startsWith('temp_')) {
                        this._updateLocalStorageEntry(rowId, fieldName, value, true);
                    }
                } else {
                    // Direct mode - save immediately
                    if (rowId && rowId !== 'new') {
                        this.updateCell(fieldName, rowId, value);
                    }
                }
            }

            // Update footer totals after paste
            this._updateFooterTotals();

            // Redraw and re-render
            this.table.draw(false);
            this._renderCheckboxColumns();
            this._renderSelectColumns();

            // Clear clipboard after paste
            await navigator.clipboard.writeText('');
            this._hasClipboardData = false;
            this._clipboardRowCount = 0;

        } catch (err) {
            console.error('Failed to paste:', err);
            alert('Unable to access clipboard. Please allow clipboard permissions.');
        }
    }

    /**
     * Paste row data from clipboard (supports tab-separated from Excel)
     * Uses centralized helpers for consistent handling
     * @param {HTMLElement} rowNode - target row node to paste into
     * @param {number} startColOverride - optional starting column index
     */
    async _pasteRowFromClipboard(rowNode, startColOverride = null) {
        if (!rowNode) return;

        try {
            const text = await navigator.clipboard.readText();
            if (!text || text.trim() === '') {
                return;
            }

            // Split by newlines first (for multi-row paste)
            const lines = text.split('\n').filter(line => line.trim() !== '');
            const startRowIdx = this.table.row(rowNode).index();

            // Determine starting column: use override, focused cell, or default
            let startColIdx;
            if (startColOverride !== null) {
                startColIdx = startColOverride;
            } else {
                // Check if there's a focused cell to get its column
                const tableNode = this.table.table().node();
                const focusedCell = tableNode.querySelector('td.focus');
                if (focusedCell) {
                    const focusedRow = focusedCell.closest('tr');
                    startColIdx = Array.from(focusedRow.children).indexOf(focusedCell);
                } else {
                    startColIdx = this.enableColumnNumber ? 1 : 0;
                }
            }

            // First, determine how many new rows we need to create
            const existingRowCount = this.table.rows().count();
            const neededRows = startRowIdx + lines.length;
            let rowsToCreate = Math.max(0, neededRows - existingRowCount);

            // If allowAddEmptyRow is false, don't create new rows - trim lines instead
            if (!this.allowAddEmptyRow && rowsToCreate > 0) {
                const maxLines = existingRowCount - startRowIdx;
                lines.length = Math.min(lines.length, maxLines);
                rowsToCreate = 0;
            }

            // Create all needed rows first (before pasting)
            for (let r = 0; r < rowsToCreate; r++) {
                const columnCount = this.table.columns().count();
                const newRowData = Array(columnCount).fill('');
                if (this.enableColumnNumber) {
                    const currentDataCount = this.table.rows().count();
                    newRowData[0] = currentDataCount + 1;
                }

                this.table.row.add(newRowData);
            }

            // Draw once after all rows are added
            if (rowsToCreate > 0) {
                this.table.draw(false);

                // Set attributes on new rows
                for (let r = 0; r < rowsToCreate; r++) {
                    const newRowIdx = existingRowCount + r;
                    const newRowNode = this.table.row(newRowIdx).node();
                    if (newRowNode) {
                        newRowNode.setAttribute('data-new-row', 'true');
                        newRowNode.setAttribute('data-id', 'new');
                    }
                }
            }

            // Now paste to all rows (existing + newly created)
            for (let lineIdx = 0; lineIdx < lines.length; lineIdx++) {
                const currentRowIdx = startRowIdx + lineIdx;
                const currentRowNode = this.table.row(currentRowIdx).node();

                if (!currentRowNode) {
                    continue;
                }

                // Parse tab-separated values for this line
                const values = lines[lineIdx].split('\t');
                const cells = currentRowNode.querySelectorAll('td');

                // Paste values into cells using centralized helper
                let valueIdx = 0;
                for (let colIdx = startColIdx; colIdx < cells.length && valueIdx < values.length; colIdx++) {
                    const cell = cells[colIdx];
                    const value = values[valueIdx].trim();

                    // Check if column is readonly
                    const th = this.table.column(colIdx).header();
                    const dataType = th ? th.getAttribute('data-type') : null;

                    if (dataType === 'readonly') {
                        valueIdx++;
                        continue;
                    }

                    // Check if cell is readonly
                    const isCellReadonly = cell.getAttribute('data-readonly') === 'true';
                    const isRowReadonly = currentRowNode.getAttribute('data-readonly-row') === 'true';

                    if (isCellReadonly || isRowReadonly) {
                        valueIdx++;
                        continue;
                    }

                    // Use centralized helper to set cell value with formatting
                    this._setCellValueWithFormatting(cell, value, colIdx, true);

                    valueIdx++;
                }

                // Get row data using centralized helper
                const rowData = this._getRowData(currentRowNode);
                const rowId = currentRowNode.getAttribute('data-id');

                // Handle batch mode / direct mode
                if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
                    // BATCH MODE
                    const isNewRow = rowId === 'new' || !rowId;

                    if (isNewRow || (rowId && rowId.startsWith('temp_'))) {
                        let tempId = rowId;
                        if (isNewRow) {
                            tempId = 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
                            currentRowNode.setAttribute('data-id', tempId);
                        }
                        currentRowNode.setAttribute('data-pending', 'true');

                        // Save all fields to pending data
                        Object.keys(rowData).forEach(fieldName => {
                            this._updateLocalStorageEntry(tempId, fieldName, rowData[fieldName], false);
                        });

                        // Check if complete and save to localStorage
                        const isComplete = this.isRowRequiredFieldsFilled(currentRowNode);

                        if (isComplete) {
                            // Remove data-new-row since row is now complete
                            currentRowNode.removeAttribute('data-new-row');

                            try {
                                localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
                            } catch (e) {
                                console.error('Error saving to localStorage:', e);
                            }

                            // Add empty row if needed (after paste completes)
                            // Use setTimeout to batch all row additions
                            setTimeout(() => {
                                if (this.isOnLastPage() && !this.isAddingEmptyRow && !this.hasEmptyRow()) {
                                    this.addEmptyRow();
                                } else {
                                }
                            }, 50);
                        }
                    } else if (rowId && !rowId.startsWith('temp_')) {
                        // Existing DB row - track edits ONLY in _edited storage
                        currentRowNode.setAttribute('data-edited', 'true');
                        Object.keys(rowData).forEach(fieldName => {
                            this._trackEditedField(rowId, fieldName, rowData[fieldName]);
                        });
                    }
                } else {
                    // DIRECT MODE
                    if (rowId === 'new' && this.isRowRequiredFieldsFilled(currentRowNode)) {
                        this.saveNewRow(currentRowNode);
                    } else if (rowId && rowId !== 'new' && !rowId.startsWith('temp_')) {
                        this.updateRowBatch(rowId, rowData);
                    }
                }
            }

            this._updateBatchCount();

            // Clear clipboard after paste
            await navigator.clipboard.writeText('');
            this._hasClipboardData = false;
            this._clipboardRowCount = 0;

            // Re-render checkbox and select columns
            setTimeout(() => {
                this._renderCheckboxColumns();
                this._renderSelectColumns();
            }, 10);

        } catch (err) {
            console.error('Failed to paste:', err);
            alert('Unable to access clipboard. Please allow clipboard permissions.');
        }
    }

    /**
     * Insert row at position (above or below) using data-based approach
     * Approach: Collect all row data, insert new empty row at correct position, rebuild table
     */
    _insertRowAt(rowNode, position) {
        if (!rowNode) return;

        // Block if allowAddEmptyRow is false (fixed row count mode)
        if (!this.allowAddEmptyRow) {
            return;
        }

        // Set flag to prevent auto-add row during insert
        this._isInsertingRow = true;

        const targetRowIdx = this.table.row(rowNode).index();
        const colCount = this.table.columns().count();

        // Collect all row data along with their attributes
        // Use raw values from DOM (data-raw-value) instead of DataTables .data()
        // to avoid capturing rendered artifacts like ▼ icons and checkbox HTML
        const allRowsData = [];
        const self = this;
        this.table.rows().every(function (rowIdx) {
            const node = this.node();
            const rawData = [];
            if (node) {
                const cells = node.querySelectorAll('td');
                cells.forEach((td, colIdx) => {
                    rawData.push(self._getCellRawValue(td, colIdx));
                });
            }
            allRowsData.push({
                data: rawData,
                id: node ? node.getAttribute('data-id') : null,
                pending: node ? node.getAttribute('data-pending') : null,
                newRow: node ? node.getAttribute('data-new-row') : null
            });
        });


        // Create data for new row
        const newRowData = {
            data: Array(colCount).fill(''),
            id: 'temp_' + Date.now(),
            pending: 'true',
            newRow: 'true'
        };

        // Insert new row at the correct position
        const insertIdx = position === 'above' ? targetRowIdx : targetRowIdx + 1;
        allRowsData.splice(insertIdx, 0, newRowData);


        // Clear raw value cache before rebuilding - cache keys use row indices
        // which will be stale after inserting a new row shifts all indices
        this._rawValueCache.clear();

        // Clear table
        this.table.clear();

        // Re-add all rows with new order
        allRowsData.forEach((rowInfo, idx) => {
            // Update column number if enabled
            if (this.enableColumnNumber) {
                rowInfo.data[0] = idx + 1;
            }
            const addedRow = this.table.row.add(rowInfo.data);
        });

        // Draw table
        this.table.draw(false);

        // Re-render checkbox and select columns after rebuild
        this._renderCheckboxColumns();
        this._renderSelectColumns();
        this._applyColumnTypes();


        // Re-apply attributes to rows
        setTimeout(() => {
            allRowsData.forEach((rowInfo, idx) => {
                const node = this.table.row(idx).node();
                if (node && rowInfo.id) {
                    node.setAttribute('data-id', rowInfo.id);
                    if (rowInfo.pending) {
                        node.setAttribute('data-pending', rowInfo.pending);
                    }
                    if (rowInfo.newRow) {
                        node.setAttribute('data-new-row', rowInfo.newRow);
                    }
                }
            });

            // Renumber rows
            if (this.enableColumnNumber) {
                this.renumberRows();
            }

            // Focus on the newly inserted row (first editable cell)
            // Use manual focus class to avoid triggering auto-add row events
            const newRowNode = this.table.row(insertIdx).node();
            if (newRowNode) {
                const cellIdx = this.enableColumnNumber ? 1 : 0;
                const firstEditableCell = newRowNode.querySelectorAll('td')[cellIdx];
                if (firstEditableCell) {
                    // Clear ALL selection highlights first
                    const tableNode = this.table.table().node();
                    tableNode.querySelectorAll('td.focus').forEach(c => c.classList.remove('focus'));
                    tableNode.querySelectorAll('td.dt-cell-range-selected').forEach(c => {
                        c.classList.remove('dt-cell-range-selected', 'dt-range-top', 'dt-range-bottom', 'dt-range-left', 'dt-range-right');
                    });
                    tableNode.querySelectorAll('tr.row-highlight').forEach(r => r.classList.remove('row-highlight'));

                    // Remove fill handle
                    this._removeFillHandle();

                    // Add focus to new cell manually
                    firstEditableCell.classList.add('focus');

                    // Update internal state
                    const cell = this.table.cell(firstEditableCell);
                    this.currentCell = cell;
                    this._selectionStart = { row: insertIdx, col: cellIdx };
                    this._selectedCells = [{ row: insertIdx, col: cellIdx }];

                    // Update fill handle
                    this._updateFillHandle(insertIdx, cellIdx);
                }
            }

            // Clear flag after insert is complete
            this._isInsertingRow = false;
        }, 100);

    }

    /**
     * Clear row data without removing the row
     * Behavior mirrors _deleteRow (calls endpoint/localStorage) but keeps the row
     */
    _clearRow(rowNode) {
        if (!rowNode) return;

        const rowId = rowNode.getAttribute('data-id');
        const deleteConf = this.endpoints.delete || {};
        const endpointUrl = deleteConf.endpoint || '';
        const hasPendingAttr = rowNode.getAttribute('data-pending') === 'true';
        const isInPendingData = this._pendingData.some(row => row._rowTempId === rowId);
        const isPendingRow = hasPendingAttr || isInPendingData;
        const isExistingRow = rowId && rowId !== 'new' && !isPendingRow;

        // preDelete hook
        if (typeof deleteConf.preDelete === 'function') {
            const shouldContinue = deleteConf.preDelete(rowId);
            if (shouldContinue === false) return;
        }

        // ========== PENDING ROW (not yet saved) ==========
        if (isPendingRow) {
            if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
                this._pendingData = this._pendingData.filter(entry => entry._rowTempId !== rowId);
                localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
            }
            this._clearRowDom(rowNode);
            return;
        }

        // ========== BATCH MODE: Track deleted existing row ==========
        if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch' && isExistingRow) {
            if (!this._deletedData.includes(rowId)) {
                this._deletedData.push(rowId);
                try {
                    localStorage.setItem(
                        this.emptyTable.storageKey + '_deleted',
                        JSON.stringify(this._deletedData)
                    );
                } catch (e) {
                    console.error('Error saving deleted data to localStorage:', e);
                }
            }

            // Remove from _editedData
            if (this._editedData[rowId]) {
                delete this._editedData[rowId];
                localStorage.setItem(
                    this.emptyTable.storageKey + '_edited',
                    JSON.stringify(this._editedData)
                );
            }

            this._clearRowDom(rowNode);
            return;
        }

        // ========== DIRECT MODE: Delete on server, then clear row ==========
        if (endpointUrl && rowId) {
            const payload = {
                operation: 'delete',
                data: { id: rowId },
                meta: { timestamp: Date.now() }
            };

            fetch(endpointUrl, {
                method: 'POST',
                headers: { ...this.getExtraHeaders(), 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            })
                .then(response => response.json())
                .then(result => {
                    this._clearRowDom(rowNode);
                    if (typeof deleteConf.postDelete === 'function') {
                        deleteConf.postDelete(result, rowId);
                    }
                })
                .catch(err => console.error('Clear row (delete) failed:', err));
        } else {
            this._clearRowDom(rowNode);
        }
    }

    /**
     * Helper: Clear row DOM content (empty cells, reset attributes)
     */
    _clearRowDom(rowNode) {
        const rowIdx = this.table.row(rowNode).index();
        const colCount = this.table.columns().count();
        const startCol = this.enableColumnNumber ? 1 : 0;

        // Clear DataTables internal data
        for (let colIdx = startCol; colIdx < colCount; colIdx++) {
            const cell = this.table.cell(rowIdx, colIdx);
            if (cell) cell.data('');
        }

        // Reset row attributes to fresh empty state
        rowNode.setAttribute('data-id', 'new');
        rowNode.setAttribute('data-new-row', 'true');
        rowNode.removeAttribute('data-pending');
        rowNode.removeAttribute('data-edited');

        // Redraw
        this.table.draw(false);
        this._renderCheckboxColumns();
        this._renderSelectColumns();
        this._updateBatchCount();

        // AFTER draw/render, force-clear DOM for non-checkbox/select columns
        const freshRowNode = this.table.row(rowIdx).node();
        if (freshRowNode) {
            const cells = freshRowNode.querySelectorAll('td');
            cells.forEach((td, idx) => {
                if (this.enableColumnNumber && idx === 0) return;
                const th = this.table.column(idx).header();
                const colType = th ? th.getAttribute('data-type') : '';
                if (colType === 'checkbox' || colType === 'select') return;
                td.innerHTML = '';
                td.removeAttribute('data-raw-value');
                td.removeAttribute('data-original-value');
                td.classList.remove('dt-error', 'dt-cell-edited', 'dt-cell-pending');
            });
        }
    }

    /**
     * Delete row with endpoint callback
     */
    _deleteRow(rowNode, rowId) {
        if (!rowNode) return;
        const deleteConf = this.endpoints.delete || {};
        const endpointUrl = deleteConf.endpoint || '';
        const hasPendingAttr = rowNode.getAttribute('data-pending') === 'true';
        const isInPendingData = this._pendingData.some(row => row._rowTempId === rowId);
        const isPendingRow = hasPendingAttr || isInPendingData;
        const isExistingRow = rowId && rowId !== 'new' && !isPendingRow;

        if (typeof deleteConf.preDelete === 'function') {
            const shouldContinue = deleteConf.preDelete(rowId);
            if (shouldContinue === false) return;
        }

        // ========== PENDING ROW (not yet saved) ==========
        if (isPendingRow) {
            if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch') {
                this._pendingData = this._pendingData.filter(entry => entry._rowTempId !== rowId);
                localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
                this._updateBatchCount();
            }
            this.table.row(rowNode).remove().draw(false);
            if (this.enableColumnNumber) this.renumberRows();
            return;
        }

        // ========== BATCH MODE: Track deleted existing row ==========
        if (this.emptyTable.enabled && this.emptyTable.saveMode === 'batch' && isExistingRow) {
            // Add to _deletedData if not already present
            if (!this._deletedData.includes(rowId)) {
                this._deletedData.push(rowId);
                try {
                    localStorage.setItem(
                        this.emptyTable.storageKey + '_deleted',
                        JSON.stringify(this._deletedData)
                    );
                } catch (e) {
                    console.error('Error saving deleted data to localStorage:', e);
                }
            }

            // Remove from _editedData if exists (no need to edit row that will be deleted)
            if (this._editedData[rowId]) {
                delete this._editedData[rowId];
                localStorage.setItem(
                    this.emptyTable.storageKey + '_edited',
                    JSON.stringify(this._editedData)
                );
            }

            this.table.row(rowNode).remove().draw(false);
            if (this.enableColumnNumber) this.renumberRows();
            this._updateBatchCount();
            return;
        }

        // ========== DIRECT MODE: Delete directly to server ==========
        if (endpointUrl && rowId) {
            // Consistent payload structure
            const payload = {
                operation: 'delete',
                data: {
                    id: rowId
                },
                meta: {
                    timestamp: Date.now()
                }
            };

            fetch(endpointUrl, {
                method: 'POST',
                headers: { ...this.getExtraHeaders(), 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            })
                .then(response => response.json())
                .then(result => {
                    this.table.row(rowNode).remove().draw(false);
                    if (this.enableColumnNumber) this.renumberRows();
                    if (typeof deleteConf.postDelete === 'function') {
                        deleteConf.postDelete(result, rowId);
                    }
                })
                .catch(err => console.error('Delete failed:', err));
        } else {
            this.table.row(rowNode).remove().draw(false);
            if (this.enableColumnNumber) this.renumberRows();
        }
    }

    onAddMenu(td) {
        this.addEmptyRow(this.table.row(td).index() + 1);
    }

    onEditMenu(td) {
        this.editRow(this.table.row(td).index());
    }

    onDeleteMenu(td) {
        const rowNode = td.closest('tr');
        const rowId = rowNode ? rowNode.getAttribute('data-id') : null;
        this._deleteRow(rowNode, rowId);
    }

    /**
     * Static method: Inject No. column BEFORE DataTable is initialized
     * Call: GridSheet.injectColumnNumberBefore('#tableId')
     * @param {string|HTMLElement} tableSelector - selector or table element
     */
    static injectColumnNumberBefore(tableSelector) {
        const tableNode = typeof tableSelector === 'string'
            ? document.querySelector(tableSelector)
            : tableSelector;

        if (!tableNode) {
            console.warn('Table not found for injectColumnNumberBefore');
            return;
        }

        // Inject header <th data-no="no">No.</th>
        const thead = tableNode.querySelector('thead tr');
        if (thead && !thead.querySelector('th[data-no="no"]')) {
            const thNo = document.createElement('th');
            thNo.setAttribute('data-no', 'no');
            thNo.textContent = '';
            thead.insertBefore(thNo, thead.firstChild);
        }

        // Inject cell <td> in each tbody row with sequential number
        const tbodyRows = tableNode.querySelectorAll('tbody tr');
        tbodyRows.forEach((row, index) => {
            const tdNo = document.createElement('td');
            tdNo.textContent = index + 1;
            row.insertBefore(tdNo, row.firstChild);
        });

        // Inject tfoot if exists
        const tfoot = tableNode.querySelector('tfoot tr');
        if (tfoot) {
            const tfNo = document.createElement('th');
            tfNo.textContent = '';
            tfoot.insertBefore(tfNo, tfoot.firstChild);
        }
    }

    // ==================== GOOGLE SHEETS BEHAVIOR HELPERS ====================

    /**
     * Add visual highlight to selected cell (selection mode)
     */
    _highlightCell(cellNode) {
        if (cellNode) {
            cellNode.classList.add('dt-cell-selected');
        }
    }

    /**
     * Remove highlight from cell
     */
    _removeHighlight(cellNode) {
        if (cellNode) {
            cellNode.classList.remove('dt-cell-selected');
        }
        // Also remove from all other cells in table
        const allSelected = this.table.table().node().querySelectorAll('.dt-cell-selected');
        allSelected.forEach(cell => cell.classList.remove('dt-cell-selected'));
    }

    /**
     * Clear cell content (for Delete/Backspace)
     * Changes:
     * - Clear content visually (allow user to delete)
     * - Mark cell as dirty (pending save)
     * - DON'T save immediately - validation and save happens on blur/navigation
     * - For existing rows: directly update (backward compat), for new rows: just clear
     */
    _clearCellContent(cell) {
        const cellNode = cell.node();
        const rowNode = cellNode.closest('tr');
        const rowId = rowNode ? rowNode.getAttribute('data-id') : null;
        const colIndex = cell.index().column;
        const colHeader = this.table.column(colIndex).header();
        const fieldName = colHeader.getAttribute('data-name');
        const dataType = colHeader ? colHeader.getAttribute('data-type') : null;

        // Check if this column allows empty values
        const allowEmpty = colHeader && colHeader.getAttribute('data-empty') === 'true';

        // Checkbox columns always have a value (true/false), so Delete toggles to false
        if (dataType === 'checkbox') {
            this._toggleCheckbox(cell);
            return;
        }

        // Readonly columns cannot be cleared
        if (dataType === 'readonly') {
            console.log('Cannot clear readonly column');
            return;
        }

        // Store old value for potential restore
        const oldValue = this._getCellRawValue(cellNode, colIndex);

        // Clear visual immediately (allow user to see delete effect)
        cellNode.textContent = '';
        cellNode.removeAttribute('data-raw-value');

        // Update DataTables data
        cell.data('').draw(false);

        // If column doesn't allow empty, show warning indicator and don't save
        if (!allowEmpty) {
            cellNode.classList.add('dt-error');
            console.log('Column', fieldName, 'does not allow empty - restoring value');
            // Restore original value
            cellNode.textContent = oldValue;
            cellNode.setAttribute('data-raw-value', oldValue);
            cell.data(oldValue).draw(false);
            return;
        }

        // For existing rows: Save immediately (allowEmpty is true)
        if (rowId && rowId !== 'new') {
            this.updateCell(fieldName, rowId, '');
        }

        // Check if pending row is now empty and reset if needed
        this._checkAndResetEmptyPendingRow(rowNode);
    }

    /**
     * Toggle checkbox value (for spacebar handling)
     */
    _toggleCheckbox(cell) {
        const cellNode = cell.node();
        const rowNode = cellNode.closest('tr');
        const colIndex = cell.index().column;
        const th = this.table.column(colIndex).header();

        // Check if column is readonly
        if (th && th.getAttribute('data-type') === 'readonly') {
            console.log('Cannot toggle checkbox - column is readonly');
            return;
        }

        // Check if row is readonly
        if (rowNode && rowNode.getAttribute('data-readonly-row') === 'true') {
            console.log('Cannot toggle checkbox - row is readonly');
            return;
        }

        const checkbox = cellNode.querySelector('input[type="checkbox"]');

        if (checkbox) {
            // Toggle the checkbox
            checkbox.checked = !checkbox.checked;

            // Trigger the change event to handle saving
            checkbox.dispatchEvent(new Event('change', { bubbles: true }));
        }
    }

    /**
     * Check if key is a printable character (letters, numbers, symbols)
     */
    _isPrintableKey(keyCode) {
        // 48-57 = 0-9
        // 65-90 = A-Z
        // 96-111 = Numpad
        // 186-222 = Symbols (;=,-./`[]\')
        // 32 = Space
        return (
            (keyCode >= 48 && keyCode <= 57) ||   // 0-9
            (keyCode >= 65 && keyCode <= 90) ||   // A-Z
            (keyCode >= 96 && keyCode <= 111) ||  // Numpad
            (keyCode >= 186 && keyCode <= 222) || // Symbols
            keyCode === 32                         // Space
        );
    }

    /**
     * Check if row has content (at least one cell filled besides No column)
     * @param {HTMLElement} rowNode - DOM node of the row
     * @returns {boolean} true if row has content
     */
    _rowHasContent(rowNode) {
        if (!rowNode) return false;

        const cells = rowNode.querySelectorAll('td');
        const startIdx = this.enableColumnNumber ? 1 : 0; // Skip No column if present

        for (let i = startIdx; i < cells.length; i++) {
            // Skip checkbox cells - by either checkbox element or data attribute
            const checkbox = cells[i].querySelector('input[type="checkbox"]');
            const isCheckboxCell = cells[i].hasAttribute('data-checkbox-value');
            if (checkbox || isCheckboxCell) continue;

            // Also skip if this is a checkbox column (check header)
            const th = this.table.column(i).header();
            if (th && th.getAttribute('data-type') === 'checkbox') continue;

            const cellContent = cells[i].textContent.trim();
            // Skip if there's input/select (currently editing)
            const hasEditor = cells[i].querySelector('input, select');
            if (hasEditor) {
                const editorValue = hasEditor.value.trim();
                if (editorValue !== '') return true;
            } else if (cellContent !== '') {
                return true;
            }
        }
        return false;
    }

    /**
     * Add new empty row (for emptyTable mode when navigating on last row)
     */
    _addNewEmptyRow() {
        // Block if allowAddEmptyRow is false (fixed row count mode)
        if (!this.allowAddEmptyRow) {
            return;
        }

        // Block auto-add if insert is in progress
        if (this._isInsertingRow) {
            console.log('Blocked auto-add: insert in progress');
            return;
        }

        const columnCount = this.table.columns().count();
        const newRowData = Array(columnCount).fill('');

        // If enableColumnNumber, set number in first column
        if (this.enableColumnNumber) {
            const currentRowCount = this.table.rows().count();
            newRowData[0] = currentRowCount + 1;
        }

        // Add row to DataTables
        this.table.row.add(newRowData).draw(false);

        // Set attribute on new row
        setTimeout(() => {
            const newRowIndex = this.table.rows().count() - 1;
            const newRowNode = this.table.row(newRowIndex).node();
            if (newRowNode) {
                newRowNode.setAttribute('data-new-row', 'true');
                newRowNode.setAttribute('data-id', 'new');
            }
        }, 10);

    }

    /**
     * Add empty row at bottom (public method for auto-add after save)
     * @param {number} refRowIndex - Reference row index (unused, for compatibility)
     */
    addEmptyRow(refRowIndex = 0) {
        // Block if allowAddEmptyRow is false (fixed row count mode)
        if (!this.allowAddEmptyRow) {
            return;
        }

        // Block if insert is in progress
        if (this._isInsertingRow) {
            console.log('Blocked addEmptyRow: insert in progress');
            return;
        }

        // ANTISPAM LOOP PATCH
        if (this.isAddingEmptyRow) return;
        this.isAddingEmptyRow = true;

        // Ensure only on last page & only one empty row in data, not counting DOM
        if (!this.isOnLastPage()) {
            this.isAddingEmptyRow = false;
            return;
        }
        if (this.hasEmptyRow()) {
            this.isAddingEmptyRow = false;
            return;
        }

        // --- ADD ROW ---
        let columnCount = this.table.columns().count();
        let newRowData = Array(columnCount).fill('');
        if (this.enableColumnNumber) {
            const currentDataCount = this.table
                .rows((idx, data, node) => !node || node.getAttribute('data-new-row') !== 'true')
                .count();
            newRowData[0] = currentDataCount + 1;
        }
        this.table.row.add(newRowData).draw(false); // DON'T immediately change page!

        setTimeout(() => {
            let newRowIndex = this.table.rows().count() - 1;
            let newRowNode = this.table.row(newRowIndex).node();
            if (newRowNode) {
                newRowNode.setAttribute('data-new-row', 'true');
                newRowNode.setAttribute('data-id', 'new');
                let firstCell = newRowNode.cells[0];
                if (firstCell) firstCell.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                setTimeout(() => this.table.cell(newRowIndex, 0).focus(), 0);
            }
            this.emptyRowExists = true;
            this.isAddingEmptyRow = false; // <------ FLAG RESET
        }, 0);
    }

    /**
     * Check if currently on last page of DataTables
     * @returns {boolean} True if on last page
     */
    isOnLastPage() {
        let info = this.table.page.info();
        // Handle empty table (pages = 0)
        if (info.pages === 0) return true;
        return this.table.page() === (info.pages - 1);
    }

    /**
     * Check if table already has an empty row (data-new-row="true")
     * @returns {boolean} True if empty row exists
     */
    hasEmptyRow() {
        return this.table
            .rows((idx, data, node) => node && node.getAttribute('data-new-row') === 'true')
            .count() > 0;
    }

    // ==================== BATCH MODE METHODS ====================

    /**
     * Add Save All button below table
     */
    _addSaveButton() {
        const wrapper = this.table.table().container();
        const buttonId = 'dt-batch-save-' + (this.emptyTable.storageKey || 'default');

        // Check if button already exists (prevent duplicate)
        if (document.getElementById(buttonId)) {
            return;
        }

        // Skip rendering entirely if both buttons are hidden
        if (this.emptyTable.showSaveButton === false && !this.emptyTable.showClearButton) {
            return;
        }

        // Create container for button
        const btnContainer = document.createElement('div');
        btnContainer.id = buttonId;
        btnContainer.className = 'dt-batch-save-container';
        btnContainer.style.cssText = 'display: flex; justify-content: flex-end; align-items: center; gap: 8px;';

        // Clear button (optional)
        if (this.emptyTable.showClearButton) {
            const clearBtn = document.createElement('button');
            clearBtn.className = 'ui button dt-batch-clear-btn';
            clearBtn.textContent = this.lang.reset || 'Reset';
            this._addEvent(clearBtn, 'click', () => {
                this._showClearConfirmDialog();
            });
            btnContainer.appendChild(clearBtn);
            this._clearButton = clearBtn;
        }

        // Save All button (conditionally rendered)
        if (this.emptyTable.showSaveButton !== false) {
            const saveBtn = document.createElement('button');
            saveBtn.className = 'ui primary button dt-batch-save-btn';
            saveBtn.textContent = this.lang.saveAll || this.emptyTable.saveButtonText;
            this._addEvent(saveBtn, 'click', () => this._saveBatchToServer());

            btnContainer.appendChild(saveBtn);
            this._saveButton = saveBtn;
        }

        // Insert AFTER DataTables wrapper
        wrapper.parentNode.insertBefore(btnContainer, wrapper.nextSibling);

        this._saveButtonContainer = btnContainer;
    }

    /**
     * Update badge count for pending data (new + edited + deleted)
     */
    _updateBatchCount() {
        if (this._countBadge) {
            const newCount = this._pendingData.length;
            const editedCount = Object.keys(this._editedData).length;
            const deletedCount = this._deletedData.length;
            const totalCount = newCount + editedCount + deletedCount;
            this._countBadge.textContent = totalCount;
            this._countBadge.style.display = totalCount > 0 ? 'inline-block' : 'none';
        }
    }

    /**
     * Load data from localStorage (pending + edited + deleted)
     * Show dialog if there's saved data
     */
    _loadFromLocalStorage() {
        try {
            // Load pending (new) rows
            const stored = localStorage.getItem(this.emptyTable.storageKey);
            const editedStored = localStorage.getItem(this.emptyTable.storageKey + '_edited');
            const deletedStored = localStorage.getItem(this.emptyTable.storageKey + '_deleted');

            let pendingData = [];
            let editedData = {};
            let deletedData = [];

            if (stored) {
                pendingData = JSON.parse(stored);
            }
            if (editedStored) {
                editedData = JSON.parse(editedStored);
            }
            if (deletedStored) {
                deletedData = JSON.parse(deletedStored);
            }

            const hasNewRows = pendingData.length > 0;
            const hasEditedRows = Object.keys(editedData).length > 0;
            const hasDeletedRows = deletedData.length > 0;

            // If there's saved data, restore or show dialog
            if (hasNewRows || hasEditedRows || hasDeletedRows) {
                if (this.emptyTable.autoRestore) {
                    // Auto-restore: apply data immediately without modal
                    this._pendingData = pendingData;
                    this._editedData = editedData;
                    this._deletedData = deletedData;
                    this._applyRecoveredData();
                } else {
                    // Manual: show recovery dialog
                    this._showRecoveryDialog(pendingData, editedData, deletedData, hasNewRows, hasEditedRows, hasDeletedRows);
                }
            } else {
                this._pendingData = [];
                this._editedData = {};
                this._deletedData = [];
            }
        } catch (e) {
            console.error('Error loading from localStorage:', e);
            this._pendingData = [];
            this._editedData = {};
            this._deletedData = [];
        }
    }

    /**
     * Show recovery dialog for saved data
     */
    _showRecoveryDialog(pendingData, editedData, deletedData, hasNewRows, hasEditedRows, hasDeletedRows) {
        // Show modal dialog
        const modal = document.createElement('div');
        modal.id = 'dt-recovery-modal';
        modal.style.cssText = '';
        modal.className = 'dt-modal-overlay';

        const dialog = document.createElement('div');
        dialog.style.cssText = '';
        dialog.className = 'dt-modal-content';

        const title = document.createElement('div');
        title.style.cssText = '';
        title.className = 'dt-modal-title';
        title.textContent = this.lang.restoreTitle;

        // Content with list format
        const content = document.createElement('div');
        content.style.cssText = '';
        content.className = 'dt-modal-body';

        const subtitle = document.createElement('div');
        subtitle.style.cssText = 'font-weight: 500; margin-bottom: 8px;';
        subtitle.textContent = this.lang.restoreMessage;
        content.appendChild(subtitle);

        const list = document.createElement('ul');
        list.style.cssText = '';

        if (hasNewRows) {
            const li = document.createElement('li');
            li.textContent = `${pendingData.length} new row${pendingData.length > 1 ? 's' : ''}`;
            list.appendChild(li);
        }
        if (hasEditedRows) {
            const li = document.createElement('li');
            li.textContent = `${Object.keys(editedData).length} edited row${Object.keys(editedData).length > 1 ? 's' : ''}`;
            list.appendChild(li);
        }
        if (hasDeletedRows) {
            const li = document.createElement('li');
            li.textContent = `${deletedData.length} deleted row${deletedData.length > 1 ? 's' : ''}`;
            list.appendChild(li);
        }
        content.appendChild(list);

        const btnContainer = document.createElement('div');
        btnContainer.style.cssText = '';
        btnContainer.className = 'dt-modal-footer';

        const discardBtn = document.createElement('button');
        discardBtn.textContent = this.lang.discard;
        discardBtn.style.cssText = '';
        discardBtn.className = 'dt-modal-btn secondary';
        discardBtn.onclick = () => {
            this._clearLocalStorage();
            modal.remove();
        };

        const applyBtn = document.createElement('button');
        applyBtn.textContent = this.lang.restore;
        applyBtn.style.cssText = '';
        applyBtn.className = 'dt-modal-btn primary';
        applyBtn.onclick = () => {
            this._pendingData = pendingData;
            this._editedData = editedData;
            this._deletedData = deletedData;
            this._applyRecoveredData();
            modal.remove();
        };

        btnContainer.appendChild(discardBtn);
        btnContainer.appendChild(applyBtn);

        dialog.appendChild(title);
        dialog.appendChild(content);
        dialog.appendChild(btnContainer);
        modal.appendChild(dialog);

        document.body.appendChild(modal);
    }

    _applyRecoveredData() {

        // Apply NEW ROWS - fill into existing empty rows
        if (this._pendingData.length > 0) {
            const colCount = this.table.columns().count();
            const startColIdx = this.enableColumnNumber ? 1 : 0; // Skip No. column

            // Find all empty rows in table
            const allRows = this.table.rows().nodes().toArray();
            const emptyRows = allRows.filter(rowNode => {
                // Empty row = row with data-id="new" or all cells empty
                const dataId = rowNode.getAttribute('data-id');
                if (dataId === 'new') return true;

                // Check if all cells (except No.) are empty
                const cells = rowNode.querySelectorAll('td');
                for (let i = startColIdx; i < cells.length; i++) {
                    if (cells[i].textContent.trim() !== '') return false;
                }
                return true;
            });


            // Fill data into empty rows
            this._pendingData.forEach((rowData, idx) => {
                const tempId = rowData._rowTempId;
                let rowNode;

                if (idx < emptyRows.length) {
                    // Use existing empty row
                    rowNode = emptyRows[idx];
                } else {
                    // Not enough empty rows, add new row (only if allowed)
                    if (!this.allowAddEmptyRow) {
                        return; // Skip this recovered row
                    }
                    const newRowData = [];
                    for (let i = 0; i < colCount; i++) {
                        newRowData.push('');
                    }
                    const newRow = this.table.row.add(newRowData).draw(false);
                    rowNode = newRow.node();
                }

                if (rowNode) {
                    // Fill data into each cell
                    for (let i = 0; i < colCount; i++) {
                        const th = this.table.column(i).header();
                        const fieldName = th.getAttribute('data-name');

                        if (fieldName && rowData[fieldName] !== undefined) {
                            const cell = this.table.cell(rowNode, i);
                            if (cell) {
                                cell.data(rowData[fieldName]);
                            }
                        }
                    }

                    // Set attributes
                    rowNode.setAttribute('data-id', tempId);
                    rowNode.setAttribute('data-pending', 'true');
                    rowNode.removeAttribute('data-new-row');
                }
            });

            // Calculate how many empty rows should remain
            // Formula: initialRows - number of new rows recovered
            const initialRows = this.emptyTable.initialRows || 5;
            const usedCount = this._pendingData.length;
            const remainingEmptyNeeded = Math.max(1, initialRows - usedCount); // Minimal 1

            const extraEmptyRows = emptyRows.slice(usedCount);
            const emptyToRemove = extraEmptyRows.length - remainingEmptyNeeded;

            // Remove excess empty rows
            if (emptyToRemove > 0) {
                extraEmptyRows.slice(remainingEmptyNeeded).forEach(rowNode => {
                    this.table.row(rowNode).remove();
                });
            }

            // Redraw table
            this.table.draw(false);

            // Renumber
            if (this.enableColumnNumber) {
                this.renumberRows();
            }
        }

        // Apply EDITED ROWS - update cells in table (cell-level format)
        if (Object.keys(this._editedData).length > 0) {
            Object.entries(this._editedData).forEach(([rowId, cellMap]) => {
                // Find row with matching data-id
                const tableNode = this.table.table().node();
                const rowNode = tableNode.querySelector(`tr[data-id="${rowId}"]`);

                if (rowNode) {
                    // Update each edited field
                    Object.entries(cellMap).forEach(([fieldName, value]) => {
                        // Skip internal keys
                        if (fieldName === '_timestamp') return;

                        // Find column index based on field name
                        const headers = this.table.columns().header().toArray();
                        const colIdx = headers.findIndex(th => th.getAttribute('data-name') === fieldName);

                        if (colIdx !== -1) {
                            const cell = this.table.cell(rowNode, colIdx);
                            if (cell) {
                                cell.data(value).draw(false);

                                // Highlight only the edited cell (not entire row)
                                const cellNode = cell.node();
                                if (cellNode) {
                                    cellNode.classList.add('dt-cell-edited');
                                }
                            }
                        }
                    });

                    // Mark row as edited
                    rowNode.setAttribute('data-edited', 'true');
                }
            });
        }

        // Apply DELETED ROWS - remove row from table
        if (this._deletedData.length > 0) {
            const tableNode = this.table.table().node();

            this._deletedData.forEach(rowId => {
                // Find row with matching data-id
                const rowNode = tableNode.querySelector(`tr[data-id="${rowId}"]`);

                if (rowNode) {
                    this.table.row(rowNode).remove();
                }
            });

            // Redraw and renumber
            this.table.draw(false);
            if (this.enableColumnNumber) {
                this.renumberRows();
            }
        }

        this._updateBatchCount();

        // Re-render checkbox columns after recovery
        setTimeout(() => this._renderCheckboxColumns(), 50);

    }

    /**
     * Save data to localStorage
     * @param {object} rowData - row data to save (must include _rowTempId)
     */
    _saveToLocalStorage(rowData) {
        // Add timestamp for tracking
        rowData._timestamp = Date.now();

        // Use _rowTempId from caller, or generate new if not present
        if (!rowData._rowTempId) {
            rowData._rowTempId = 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
        }

        this._pendingData.push(rowData);

        try {
            localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
            this._updateBatchCount();
        } catch (e) {
            console.error('[DataTable] Failed to save to localStorage:', e);
        }
    }

    /**
     * Track edited field for existing DB row (batch mode - cell-level)
     * Stores data per-cell for granular tracking
     * @param {string} rowId - ID of the row in database
     * @param {string} fieldName - name of field being edited
     * @param {*} value - new value
     */
    _trackEditedField(rowId, fieldName, value) {
        if (!rowId || !fieldName) return;

        // Initialize entry if not exists
        if (!this._editedData[rowId]) {
            this._editedData[rowId] = {
                _timestamp: Date.now()
            };
        }

        // Store per-cell value
        this._editedData[rowId][fieldName] = value;
        this._editedData[rowId]._timestamp = Date.now();

        // Mark the specific cell with edited class
        const tableNode = this.table.table().node();
        const rowNode = tableNode.querySelector(`tr[data-id="${rowId}"]`);
        if (rowNode) {
            const headers = this.table.columns().header().toArray();
            const colIdx = headers.findIndex(th => th.getAttribute('data-name') === fieldName);
            if (colIdx !== -1) {
                const cellNode = rowNode.querySelectorAll('td')[colIdx];
                if (cellNode) {
                    cellNode.classList.add('dt-cell-edited');
                }
            }
        }

        // Save to localStorage
        try {
            localStorage.setItem(this.emptyTable.storageKey + '_edited', JSON.stringify(this._editedData));
        } catch (e) {
            console.error('[DataTable] Failed to save edited data:', e);
        }
    }

    /**
     * Save field to pending data for new/temp row (batch mode)
     * @param {HTMLElement} rowNode - row element
     * @param {string} fieldName - field name
     * @param {*} value - value
     */
    _saveFieldToPending(rowNode, fieldName, value) {
        if (!rowNode || !fieldName) return;

        let tempId = rowNode.getAttribute('data-id');

        // Generate temp_id if not present
        if (tempId === 'new') {
            tempId = 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
            rowNode.setAttribute('data-id', tempId);
            rowNode.removeAttribute('data-new-row');
        }

        // Find existing entry or create new
        let existingIdx = this._pendingData.findIndex(e => e._rowTempId === tempId);
        if (existingIdx >= 0) {
            // Update existing entry
            this._pendingData[existingIdx][fieldName] = value;
            this._pendingData[existingIdx]._timestamp = Date.now();
        } else {
            // Create new entry
            const rowData = { _rowTempId: tempId, _timestamp: Date.now() };
            rowData[fieldName] = value;
            this._pendingData.push(rowData);
        }

        // Save to localStorage
        try {
            localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
        } catch (e) {
            console.error('[DataTable] Failed to save pending data:', e);
        }
    }

    /**
     * Save complete row to pending data (when all required fields are filled)
     * @param {HTMLElement} rowNode - row element
     * @param {string} tempId - temp ID for the row
     */
    _saveCompleteRowToPending(rowNode, tempId) {
        if (!rowNode || !tempId) return;

        // Collect all field values from the row
        const rowData = { _rowTempId: tempId, _timestamp: Date.now() };
        const headers = this.table.columns().header().toArray();

        Array.from(rowNode.children).forEach((cell, idx) => {
            if (this.enableColumnNumber && idx === 0) return; // Skip No. column
            const th = headers[idx];
            if (th) {
                const fname = th.getAttribute('data-name');
                if (fname) {
                    rowData[fname] = this._getCellValue(cell, idx);
                }
            }
        });

        // Check if entry already exists
        const existingIdx = this._pendingData.findIndex(e => e._rowTempId === tempId);
        if (existingIdx === -1) {
            this._pendingData.push(rowData);
        } else {
            // Update existing entry with all fields
            Object.assign(this._pendingData[existingIdx], rowData);
            this._pendingData[existingIdx]._timestamp = Date.now();
        }

        // Save to localStorage
        try {
            localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
        } catch (e) {
            console.error('[DataTable] Failed to save pending data:', e);
        }
    }

    /**
     * Show clear confirmation dialog (modal)
     */
    _showClearConfirmDialog() {
        // Show modal dialog
        const modal = document.createElement('div');
        modal.id = 'dt-clear-modal';
        modal.style.cssText = '';
        modal.className = 'dt-modal-overlay';

        const dialog = document.createElement('div');
        dialog.style.cssText = '';
        dialog.className = 'dt-modal-content';

        const title = document.createElement('div');
        title.style.cssText = '';
        title.className = 'dt-modal-title';
        title.textContent = this.lang.clearTitle;

        const content = document.createElement('div');
        content.style.cssText = '';
        content.className = 'dt-modal-body';
        content.textContent = this.lang.clearMessage;

        const btnContainer = document.createElement('div');
        btnContainer.style.cssText = '';
        btnContainer.className = 'dt-modal-footer';

        const cancelBtn = document.createElement('button');
        cancelBtn.textContent = this.lang.cancel;
        cancelBtn.style.cssText = '';
        cancelBtn.className = 'dt-modal-btn secondary';
        cancelBtn.onclick = () => modal.remove();

        const clearBtn = document.createElement('button');
        clearBtn.textContent = this.lang.clear;
        clearBtn.style.cssText = '';
        clearBtn.className = 'dt-modal-btn danger';
        clearBtn.onclick = () => {
            modal.remove();
            this._clearAndResetTable();
        };

        btnContainer.appendChild(cancelBtn);
        btnContainer.appendChild(clearBtn);

        dialog.appendChild(title);
        dialog.appendChild(content);
        dialog.appendChild(btnContainer);
        modal.appendChild(dialog);

        document.body.appendChild(modal);
    }

    /**
     * Clear all data and reset table to initial state
     * - Remove pending/new rows
     * - Revert edited rows to original (reload page)
     * - Restore deleted rows (reload page)
     */
    _clearAndResetTable() {
        // Clear localStorage
        this._clearLocalStorage();

        // Reload page to get original data from DB
        // This is the simplest way to revert all changes
        window.location.reload();
    }

    /**
     * Update entry in localStorage based on rowTempId
     * @param {string} rowTempId - temp ID of row to update
     * @param {string} fieldName - name of field to update
     * @param {string} value - new value
     */
    _updateLocalStorageEntry(rowTempId, fieldName, value) {
        // Find entry with matching rowTempId
        const entryIndex = this._pendingData.findIndex(entry => entry._rowTempId === rowTempId);

        if (entryIndex !== -1) {
            // Update field in existing entry
            this._pendingData[entryIndex][fieldName] = value;
            this._pendingData[entryIndex]._timestamp = Date.now();
        } else {
            // Entry not found - create new entry
            const newEntry = {
                _rowTempId: rowTempId,
                _timestamp: Date.now()
            };
            newEntry[fieldName] = value;
            this._pendingData.push(newEntry);
        }

        // Always save to localStorage (per-cell tracking)
        try {
            localStorage.setItem(this.emptyTable.storageKey, JSON.stringify(this._pendingData));
        } catch (e) {
            console.error('Error updating localStorage:', e);
        }
    }

    /**
     * Track existing row being edited (for batch save - cell-level)
     * Delegates to _trackEditedField for consistent cell-level tracking
     * @param {string} rowId - Row ID from database
     * @param {string} fieldName - name of field being edited
     * @param {string} value - new value
     */
    _trackEditedRow(rowId, fieldName, value) {
        this._trackEditedField(rowId, fieldName, value);
        this._updateBatchCount();
    }

    /**
     * Clear localStorage and data (pending + edited + deleted)
     */
    _clearLocalStorage() {
        this._pendingData = [];
        this._editedData = {};
        this._deletedData = [];
        try {
            localStorage.removeItem(this.emptyTable.storageKey);
            localStorage.removeItem(this.emptyTable.storageKey + '_edited');
            localStorage.removeItem(this.emptyTable.storageKey + '_deleted');
            this._updateBatchCount();
        } catch (e) {
            console.error('Error clearing localStorage:', e);
        }
    }
    /**
     * Validate all pending data (new rows + edited rows)
     * Also scans DOM for cells with dt-error class (invalid values not saved)
     * Returns array of validation errors with details
     * @returns {Array} Array of error objects: [{rowIndex, field, value, type, tempId, isNew}]
     */
    _validateAllPendingData() {
        const errors = [];
        const tableNode = this.table.table().node();
        const headers = tableNode.querySelectorAll('thead th');

        // Build column type map and field index from headers
        // Support both data-field and data-name for backwards compatibility
        const columnTypes = {};
        const fieldByIndex = {};
        headers.forEach((th, idx) => {
            // Try data-field first, then data-name for backwards compatibility
            const fieldName = th.getAttribute('data-field') || th.getAttribute('data-name');
            const dataType = th.getAttribute('data-type');
            if (fieldName && dataType) {
                columnTypes[fieldName] = dataType;
            }
            if (fieldName) {
                fieldByIndex[idx] = fieldName;
            }
        });

        // Validate NEW ROWS (pending data)
        this._pendingData.forEach((row, index) => {
            Object.entries(row).forEach(([field, value]) => {
                // Skip internal fields
                if (field.startsWith('_')) return;

                if (value && columnTypes[field]) {
                    const isValid = this._validateField(value, columnTypes[field]);
                    if (!isValid) {
                        errors.push({
                            rowIndex: index,
                            field: field,
                            value: value,
                            type: columnTypes[field],
                            tempId: row._rowTempId,
                            isNew: true
                        });
                    }
                }
            });
        });

        // Validate EDITED ROWS (cell-level format)
        Object.entries(this._editedData).forEach(([rowId, cellMap]) => {
            Object.entries(cellMap).forEach(([field, value]) => {
                // Skip internal keys
                if (field.startsWith('_')) return;

                if (value && columnTypes[field]) {
                    const isValid = this._validateField(value, columnTypes[field]);
                    if (!isValid) {
                        errors.push({
                            rowIndex: null,  // Not applicable for edited rows
                            field: field,
                            value: value,
                            type: columnTypes[field],
                            rowId: rowId,
                            isNew: false
                        });
                    }
                }
            });
        });

        // ===== SCAN DOM for cells with dt-error class =====
        // These are cells with invalid values that weren't saved to _pendingData
        const errorCells = tableNode.querySelectorAll('tbody td.dt-error');
        errorCells.forEach(cell => {
            const row = cell.closest('tr');
            const cellIndex = cell.cellIndex;
            const fieldName = fieldByIndex[cellIndex];
            const dataType = columnTypes[fieldName];
            const tempId = row ? row.getAttribute('data-temp-id') : null;
            const rowId = row ? row.getAttribute('data-id') : null;

            // Get the displayed value (which is invalid)
            const rawValue = cell.getAttribute('data-raw-value') || cell.textContent.trim();

            // Check if this error is already in the list (avoid duplicates)
            const isDuplicate = errors.some(e =>
                (e.tempId && e.tempId === tempId && e.field === fieldName) ||
                (e.rowId && e.rowId === rowId && e.field === fieldName)
            );

            if (!isDuplicate && fieldName && dataType) {
                errors.push({
                    rowIndex: row ? row.rowIndex : null,
                    field: fieldName,
                    value: rawValue,
                    type: dataType,
                    tempId: tempId,
                    rowId: rowId,
                    isNew: !!tempId && !rowId,
                    fromDom: true  // Flag to indicate this was found in DOM, not pending data
                });
            }
        });

        return errors;
    }

    /**
     * Scroll to and highlight the first error cell
     * @param {Object} error - Error object from _validateAllPendingData
     */
    _scrollToErrorCell(error) {
        const tableNode = this.table.table().node();
        let targetRow = null;
        let targetCell = null;

        if (error.isNew && error.tempId) {
            // Find row by temp ID for new rows
            targetRow = tableNode.querySelector(`tbody tr[data-temp-id="${error.tempId}"]`);
        } else if (error.rowId) {
            // Find row by data-id for existing rows
            targetRow = tableNode.querySelector(`tbody tr[data-id="${error.rowId}"]`);
        }

        if (targetRow) {
            // Find the specific cell by data-field attribute or column index
            const headers = tableNode.querySelectorAll('thead th');
            let colIndex = -1;
            headers.forEach((th, idx) => {
                if (th.getAttribute('data-field') === error.field) {
                    colIndex = idx;
                }
            });

            if (colIndex >= 0) {
                targetCell = targetRow.querySelectorAll('td')[colIndex];
            }
        }

        if (targetCell) {
            // Add error highlight
            targetCell.classList.add('dt-error');

            // Scroll into view
            targetCell.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });

            // Focus the cell using KeyTable if available
            try {
                const dtCell = this.table.cell(targetCell);
                if (dtCell && dtCell.index()) {
                    this.table.cell(dtCell.index()).focus();
                }
            } catch (e) {
                // If KeyTable focus fails, just do visual scroll
            }

            // Remove error class after animation
            setTimeout(() => {
                targetCell.classList.remove('dt-error');
            }, 3000);
        }
    }

    /**
     * Save all pending data to server (newRows + editedRows + deletedRows)
     */
    async _saveBatchToServer() {
        const hasNewRows = this._pendingData.length > 0;
        const hasEditedRows = Object.keys(this._editedData).length > 0;
        const hasDeletedRows = this._deletedData.length > 0;

        if (!hasNewRows && !hasEditedRows && !hasDeletedRows) {
            console.log('No data to save');
            return;
        }

        // Use batch endpoint, fallback to save endpoint for backward compatibility
        const batchConf = this.endpoints.batch || this.endpoints.save || {};
        const endpointUrl = batchConf.endpoint || '';

        if (!endpointUrl) {
            console.warn('Batch endpoint not configured!');
            return;
        }

        // ===== VALIDATION STEP =====
        // Validate all pending data before saving
        const validationErrors = this._validateAllPendingData();

        // ===== Prepare NEW ROWS data =====
        const insertRows = this._pendingData.map(row => {
            const filtered = { _tempId: row._rowTempId };
            if (this.allowedFields && Array.isArray(this.allowedFields)) {
                this.allowedFields.forEach(f => {
                    if (row[f] !== undefined) filtered[f] = row[f];
                });
            } else {
                // Remove internal fields
                Object.keys(row).forEach(k => {
                    if (!k.startsWith('_')) filtered[k] = row[k];
                });
            }
            return filtered;
        });

        // ===== Prepare EDITED ROWS data =====
        // Cell-level data is grouped back to row-level for API compatibility
        // Format: [ { id: "123", fields: { name: "new name", email: "new@email" } }, ... ]
        const updateRows = Object.entries(this._editedData).map(([rowId, cellMap]) => {
            const fields = {};
            Object.entries(cellMap).forEach(([fieldName, value]) => {
                // Skip internal keys like _timestamp
                if (fieldName.startsWith('_')) return;
                fields[fieldName] = value;
            });
            return { id: rowId, fields: fields };
        });

        // ===== Prepare DELETED ROWS data =====
        // Format: [ "id1", "id2", ... ]
        const deleteIds = [...this._deletedData];

        // Calculate total field count
        const insertFieldCount = insertRows.reduce((sum, row) => {
            return sum + Object.keys(row).filter(k => k !== '_tempId').length;
        }, 0);
        const updateFieldCount = updateRows.reduce((sum, row) => {
            return sum + Object.keys(row.fields).length;
        }, 0);
        const totalFieldCount = insertFieldCount + updateFieldCount;

        // Consistent payload structure
        const payload = {
            operation: 'batch',
            data: {
                insert: insertRows,
                update: updateRows,
                delete: deleteIds
            },
            meta: {
                timestamp: Date.now(),
                hasErrors: validationErrors.length > 0,
                validationErrors: validationErrors,
                rowCount: insertRows.length + updateRows.length + deleteIds.length,
                insertCount: insertRows.length,
                updateCount: updateRows.length,
                deleteCount: deleteIds.length,
                fieldCount: totalFieldCount
            }
        };

        // Call preSave hook if defined (now with consistent payload)
        if (typeof batchConf.preSave === 'function') {
            const shouldProceed = batchConf.preSave(payload);

            // If preSave returns false, abort save and scroll to first error
            if (shouldProceed === false) {
                console.log('Batch save aborted by preSave hook');

                // Auto-scroll to first error cell
                if (validationErrors.length > 0) {
                    this._scrollToErrorCell(validationErrors[0]);
                }
                return;
            }
        }

        // Disable button during process
        if (this._saveButton) {
            this._saveButton.disabled = true;
            this._saveButton.textContent = 'Saving...';
        }

        try {

            // Send to server as batch
            const response = await fetch(endpointUrl, {
                method: 'POST',
                headers: {
                    ...this.getExtraHeaders(),
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                throw new Error('Server error: ' + await response.text());
            }

            const result = await response.json();
            console.log('Batch save result:', result);

            // Callback postSave if exists
            if (typeof batchConf.postSave === 'function') {
                batchConf.postSave(result);
            }

            // SERVER-FIELD AUTO UPDATE: Update inserted rows with server-generated fields
            if (result.data?.inserted && Array.isArray(result.data.inserted)) {
                result.data.inserted.forEach(insertedItem => {
                    // Find row by tempId
                    const tempId = insertedItem.tempId;
                    if (tempId) {
                        const rowNode = this.table.table().node().querySelector(`tr[data-id="${tempId}"]`);
                        if (rowNode) {
                            // Update row ID and server fields
                            this._updateRowFromServerResponse(rowNode, insertedItem);
                        }
                    }
                });
            }

            // Cleanup: Remove pending state from ALL rows that were just saved
            // This handles cases where server doesn't return id mapping
            const tableNode = this.table.table().node();
            const pendingRows = tableNode.querySelectorAll('tr[data-pending="true"]');
            pendingRows.forEach(row => {
                row.removeAttribute('data-pending');
                row.removeAttribute('data-new-row');
                row.classList.remove('dt-row-pending');
            });

            // Also clean up any remaining temp_ ID rows that were in our pending data
            // but server didn't return explicit id mapping
            insertRows.forEach(insertedRow => {
                if (insertedRow._tempId) {
                    const tempRow = tableNode.querySelector(`tr[data-id="${insertedRow._tempId}"]`);
                    if (tempRow) {
                        tempRow.removeAttribute('data-pending');
                        tempRow.removeAttribute('data-new-row');
                        tempRow.classList.remove('dt-row-pending');
                    }
                }
            });

            // Clear localStorage and pending data
            this._clearLocalStorage();
            console.log(`Successfully saved ${insertRows.length} new rows, ${updateRows.length} edited rows, ${deleteIds.length} deleted rows!`);

            // Refresh table if needed
            // this.table.ajax.reload();

        } catch (error) {
            console.error('Batch save failed:', error);
        } finally {
            // Re-enable button
            if (this._saveButton) {
                this._saveButton.disabled = false;
                this._saveButton.textContent = this.emptyTable.saveButtonText;
            }
        }
    }



    /**
     * Get pending data count
     */
    getPendingCount() {
        return this._pendingData.length;
    }

    /**
     * Get all pending data
     */
    getPendingData() {
        return [...this._pendingData];
    }

    /**
     * Public method to trigger batch save
     * Allows external code to trigger the same save logic as the built-in Save All button
     * Useful when emptyTable.showSaveButton is false and user handles save via custom UI
     * @returns {Promise} Resolves when save is complete
     * 
     * @example
     * // With showSaveButton: false
     * document.getElementById('my-save-btn').addEventListener('click', () => {
     *     gridsheet.saveBatch();
     * });
     */
    saveBatch() {
        return this._saveBatchToServer();
    }

    // ========================================
    // REGION: Event Management System
    // ========================================

    /**
     * Add an event listener with tracking for cleanup
     * @param {Element} element - DOM element to attach listener to
     * @param {string} eventType - Event type (e.g., 'click', 'keydown')
     * @param {Function} handler - Event handler function
     * @param {Object|boolean} options - Event listener options (optional)
     * @returns {Function} The bound handler function (for removal if needed)
     */
    _addEvent(element, eventType, handler, options = false) {
        if (!element) {
            console.warn('_addEvent: element is null or undefined');
            return null;
        }

        // Bind handler to this context if it's a method reference
        const boundHandler = typeof handler === 'function' ? handler.bind(this) : handler;

        element.addEventListener(eventType, boundHandler, options);

        // Store reference for cleanup
        this._eventListeners.push({
            element,
            eventType,
            handler: boundHandler,
            options
        });

        return boundHandler;
    }

    /**
     * Remove a specific event listener
     * @param {Element} element - DOM element
     * @param {string} eventType - Event type
     * @param {Function} handler - The bound handler function returned by _addEvent
     */
    _removeEvent(element, eventType, handler) {
        if (!element || !handler) return;

        element.removeEventListener(eventType, handler);

        // Remove from registry
        this._eventListeners = this._eventListeners.filter(
            e => !(e.element === element && e.eventType === eventType && e.handler === handler)
        );
    }

    /**
     * Remove all registered event listeners (for cleanup/destroy)
     */
    _removeAllEvents() {
        this._eventListeners.forEach(({ element, eventType, handler, options }) => {
            try {
                element.removeEventListener(eventType, handler, options);
            } catch (err) {
                console.warn('Error removing event listener:', err);
            }
        });
        this._eventListeners = [];
    }

    /**
     * Destroy GridSheet instance and cleanup resources
     * 
     * Removes all event listeners, DOM elements (fill handle, context menu),
     * and internal references. Call this method before destroying DataTable.
     * 
     * @returns {void}
     * 
     * @example
     * // Cleanup before destroying DataTable
     * interactive.destroy();
     * table.destroy();
     */
    destroy() {
        // Remove all tracked event listeners
        this._removeAllEvents();

        // Remove fill handle
        if (this._fillHandleElement) {
            this._fillHandleElement.remove();
            this._fillHandleElement = null;
        }

        // Remove context menu
        const contextMenu = document.getElementById('dt-context-menu');
        if (contextMenu) {
            contextMenu.remove();
        }

        // Remove save button if exists
        if (this._saveButton) {
            const container = this._saveButton.closest('.dt-batch-controls');
            if (container) container.remove();
        }

        // Clear selection state
        this._selectedRows = [];
        this._selectedCells = [];
        this._selectionStart = null;
        this._selectionEnd = null;

        // Clear pending data
        this._pendingData = [];
        this._editedData = {};
        this._deletedData = [];

        // Reset flags
        this.isEditMode = false;
        this._hasClipboardData = false;
        this._clipboardRowCount = 0;

        // Remove row highlights
        this._removeRowHighlight();

        // Clear cell focus styling
        const tableNode = this.table.table().node();
        tableNode.querySelectorAll('.focus, .dt-cell-editing').forEach(cell => {
            cell.classList.remove('focus', 'dt-cell-editing');
        });

        console.log('GridSheet destroyed successfully');
    }
}
