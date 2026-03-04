/**
 * GridSheet Custom Language
 * ================================================
 * Override default language labels here.
 * This file will NOT be overwritten during GridSheet update.
 * 
 * Usage:
 * <script src="datatables-gridsheet/js/gridsheet.js"></script>
 * <script src="datatables-gridsheet/custom/gridsheet.lang.js"></script>
 * 
 * const gridsheet = new GridSheet({
 *     language: GridSheetLang.id,  // Indonesian
 *     // ... other options
 * });
 */

const GridSheetLang = {

    // ====== ENGLISH (Default) ======
    en: {
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
        readonlyRow: 'Readonly Row',

        // Buttons
        saveAll: 'Save',
        cancel: 'Cancel',
        discard: 'Discard',
        restore: 'Restore',
        clear: 'Delete',
        reset: 'Reset',

        // Dropdown
        noResult: 'No results found',

        // Restore Modal
        restoreTitle: 'Restore unsaved changes?',
        restoreMessage: 'You have pending changes:',
        restoreNewRow: 'new row',
        clearTitle: 'Clear all changes?',
        clearMessage: 'This will discard all pending changes and reload the page.'
    },

    // ====== INDONESIAN ======
    id: {
        // Context Menu
        copyRow: 'Salin Baris',
        pasteRow: 'Tempel Baris',
        pasteRows: 'Tempel {n} Baris',  // NEW: For multi-row paste
        deleteRow: 'Hapus Baris',
        clearRow: 'Bersihkan Baris',
        copyRows: 'Salin {n} Baris',
        clearRows: 'Bersihkan {n} Baris',
        deleteRows: 'Hapus {n} Baris',
        insertAbove: 'Sisip di Atas',
        insertBelow: 'Sisip di Bawah',
        readonlyRow: 'Baris Hanya-Baca',

        // Buttons
        saveAll: 'Simpan',
        cancel: 'Batal',
        discard: 'Buang',
        restore: 'Pulihkan',
        clear: 'Hapus',
        reset: 'Atur Ulang',

        // Dropdown
        noResult: 'Tidak ada hasil',

        // Restore Modal
        restoreTitle: 'Pulihkan perubahan?',
        restoreMessage: 'Anda memiliki perubahan:',
        restoreNewRow: 'baris baru',

        // Clear Modal
        clearTitle: 'Hapus semua perubahan?',
        clearMessage: 'Ini akan menghapus semua perubahan dan memuat ulang halaman.'
    },

    // ====== TEMPLATE - Copy for other languages ======
    /*
    xx: {
        // Context Menu
        copyRow: '',
        pasteRow: '',
        pasteRows: 'Paste {n} Rows',
        deleteRow: '',
        clearRow: '',
        copyRows: '',
        clearRows: '',
        deleteRows: '',
        insertAbove: '',
        insertBelow: '',
        readonlyRow: '',
        copyCells: '',
        pasteCells: '',
        copy: '',
        paste: '',  // NEW: For single cell paste
        readonly: '',
        clear: '',
        reset: '',

        // Dropdown
        noResult: '',

        // Buttons
        saveAll: '',
        cancel: '',
        discard: '',
        restore: '',

        // Restore Modal
        restoreTitle: '',
        restoreMessage: '',
        restoreNewRow: '',

        // Clear/Reset Modal
        clearTitle: '',
        clearMessage: ''
    }
    */
};
