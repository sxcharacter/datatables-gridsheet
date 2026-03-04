# GridSheet Examples & Demos

Collection of working examples demonstrating GridSheet features and use cases.

---

## Quick Start

1. **Start local server** (XAMPP, WAMP, or similar)
2. **Open browser** to `http://localhost/datatables-gridsheet/examples/`
3. **Select a demo** from the list below

> **Note:** These examples use **PHP** (`endpoints/`) for the backend API logic. You must run them on a PHP-enabled web server (like Apache/Nginx via XAMPP) for saving, updating, and deleting to work.

---

## Main Demos

### Core Functionality

| Demo | File | Description | Save Mode |
|------|------|-------------|-----------|
| **Real-time Sync** | `basic.html` | All column types with immediate save | Direct |
| **Batch Save** | `batch.html` | Collect changes, save all at once | Batch |

---

### Feature Demos

Located in `demos/` subfolder:

| Demo | File | Description |
|------|------|-------------|
| **Themes** | `demos/themes.html` | Light/Dark mode toggle |
| **Formula & Footer** | `demos/formula-footer.html` | SUM, AVG, COUNT calculations |
| **Readonly Columns** | `demos/readonly-columns.html` | Non-editable cells |
| **Custom Validation** | `demos/custom-validation.html` | preSave callbacks |

---

### Feature Documentation Examples

Located in `features/` subfolder:

| Category | Examples |
|----------|----------|
| **Configuration** | `features/configuration.html` |
| **Column Types** | `features/column-types.html` |
| **Formats** | `features/formats.html` |
| **Validation** | `features/validation.html` |
| **Endpoints** | `features/endpoints.html` |
| **Formula** | `features/formula.html` |
| **Theming** | `features/theming.html` |
| **Language** | `features/language.html` |
| **Copy & Paste** | `features/copy-paste.html` |
| **Context Menu** | `features/context-menu.html` |
| **Keyboard** | `features/keyboard.html` |

---

## Detailed Examples

### 1. Real-time Sync Mode (`basic.html`)

**Features:**
- All column types (text, email, select, date, time, checkbox, currency, percentage, textarea)
- Direct save mode - every change saved immediately to server
- Full keyboard navigation
- Context menu with row operations

**Best for:** Frequently changing data requiring high consistency

```javascript
const gridsheet = new GridSheet({
    tableSelector: '#basicExample',
    emptyTable: {
        enabled: true,
        initialRows: 2,
        saveMode: 'direct'  // Save immediately
    },
    columns: {
        0: { name: 'name', type: 'text' },
        1: { name: 'email', type: 'email' },
        // ... more columns
    },
    endpoints: {
        save: { endpoint: 'endpoints/save.php' },
        update: { endpoint: 'endpoints/update.php' },
        delete: { endpoint: 'endpoints/delete.php' }
    }
});
```

---

### 2. Batch Save Mode (`batch.html`, `batch-test.html`)

**Features:**
- Changes stored in localStorage first
- "Save All" button to submit all changes at once
- Validation before batch save
- Confirmation dialog with change summary

**Best for:** Bulk data entry, unstable connections

```javascript
const gridsheet = new GridSheet({
    tableSelector: '#batchExample',
    emptyTable: {
        enabled: true,
        initialRows: 6,
        saveMode: 'batch',  // Save to localStorage first
        storageKey: 'demo_batch_full',
        showSaveButton: true,
        showClearButton: true
    },
    endpoints: {
        batch: {
            endpoint: 'endpoints/batch.php',
            preSave: function(payload) {
                // Validate all changes
                if (payload.meta.hasErrors) {
                    alert('Fix validation errors first');
                    return false;
                }
                return confirm(`Save ${payload.meta.insertCount} new, ${payload.meta.updateCount} edited rows?`);
            },
            postSave: function(result) {
                if (result.status === 'ok') {
                    alert('Saved successfully!');
                }
            }
        }
    }
});
```

---

### 3. Empty Table Mode (`empty.html`)

**Features:**
- Table starts completely empty
- Configurable initial rows (default: 10)
- Direct save mode for immediate persistence
- Auto-add new row when reaching last page

**Best for:** New data entry forms

```javascript
const gridsheet = new GridSheet({
    tableSelector: '#emptyExample',
    emptyTable: {
        enabled: true,
        initialRows: 10,  // Start with 10 empty rows
        saveMode: 'direct'
    },
    columns: {
        0: { name: 'name', type: 'text' },
        // ... more columns
    }
});
```

---

## Column Types Reference

All demos showcase these column types:

| Type | Description | Example |
|------|-------------|---------|
| `text` | Standard text input | Name, Address |
| `textarea` | Multi-line text | Notes, Description |
| `email` | Email with validation | user@example.com |
| `number` | Numeric input | Age, Quantity |
| `currency` | Currency format | Rp 50.000 |
| `percentage` | Percentage format | 85% |
| `date` | Date picker | 15/05/1992 |
| `time` | Time picker | 09:00 |
| `select` | Searchable dropdown | Office location |
| `checkbox` | Toggle true/false | Active status |
| `readonly` | Display only | Created date |

---

## Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Enter edit mode | `Enter` |
| Cancel edit | `Escape` |
| Save & move next | `Tab` |
| Toggle checkbox | `Space` |
| Clear cell | `Delete` |
| Copy selection | `Ctrl+C` |
| Paste | `Ctrl+V` |
| Context menu | `Right-click` |

---

## Context Menu Features

### Multi-Row Selection
- Copy {n} Rows
- **Paste {n} Rows** (from Excel or GridSheet)
- Clear {n} Rows
- Delete {n} Rows

### Single Cell
- Copy
- **Paste {n} Rows** (when clipboard has multi-row data)
- Readonly toggle
- Clear

### No. Column (Row Selector)
- Copy Row
- Clear Row
- **Paste {n} Rows** 
- Insert Row Above/Below
- Readonly Row toggle
- Delete Row

---

## Testing Copy/Paste from Excel

1. **Open Excel** or Google Sheets
2. **Enter test data:**
   ```
   John    john@email.com    New York
   Jane    jane@email.com    Miami
   ```
3. **Select cells** and copy (`Ctrl+C`)
4. **Open demo** (e.g., `batch-test.html`)
5. **Right-click** on any cell
6. **Select "Paste 2 Rows"** from context menu 

**Expected result:** Data pasted starting from clicked cell

**Debug:** Open browser console (F12) to see clipboard detection logs

---

## File Structure

```
examples/
├── basic.html           # Real-time sync demo
├── batch.html           # Batch save demo
├── batch-test.html      # Batch test (empty, 10 rows)
├── batch-self.html      # Alternative batch demo
├── empty.html           # Empty table demo
├── demo-page.css        # Common demo styles
├── demo-index.css       # Index page styles
├── demos/               # Feature-specific demos
│   ├── themes.html
│   ├── formula-footer.html
│   ├── readonly-columns.html
│   └── custom-validation.html
├── features/            # Documentation examples
│   ├── configuration.html
│   ├── column-types.html
│   ├── formats.html
│   ├── validation.html
│   ├── endpoints.html
│   ├── formula.html
│   ├── theming.html
│   ├── language.html
│   ├── copy-paste.html
│   ├── context-menu.html
│   └── keyboard.html
└── endpoints/           # Backend examples (PHP)
    ├── save.php
    ├── update.php
    ├── delete.php
    ├── batch.php
    └── office.php
```

---

## Backend Integration

All demos use PHP endpoints in `endpoints/` folder:

| Endpoint | Purpose | Example File |
|----------|---------|--------------|
| `save.php` | Insert new row | `examples/endpoints/save.php` |
| `update.php` | Update existing row | `examples/endpoints/update.php` |
| `delete.php` | Delete row | `examples/endpoints/delete.php` |
| `batch.php` | Batch operations | `examples/endpoints/batch.php` |
| `office.php` | Select options | `examples/endpoints/office.php` |

---

## Troubleshooting

### Issue: Demo not loading

**Solution:**
1. Ensure XAMPP/WAMP is running
2. Check Apache is started
3. Verify URL: `http://localhost/datatables-gridsheet/examples/`

---

### Issue: Copy/Paste not working

**Solution:**
1. Hard refresh browser: `Ctrl+Shift+R`
2. Check browser console (F12) for errors
3. Ensure using HTTPS or localhost (Clipboard API requirement)
4. Try keyboard shortcut `Ctrl+V` as alternative

---

### Issue: Save button not appearing (Batch Mode)

**Solution:**
1. Check `emptyTable.showSaveButton: true` in config
2. Make some changes to the table
3. Check localStorage in DevTools (F12 → Application → Local Storage)

---

**Last Updated:** March 2026  
**Version:** 1.0.0
