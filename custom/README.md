# GridSheet Customization

This folder contains files to customize GridSheet **without modifying core files**.  
Files in this folder will **not be overwritten** during GridSheet update.

## Files

| File | Description |
|------|-------------|
| `theme-classic.css` | **Classic Green** - Google Sheets style |
| `theme-ocean.css` | **Ocean Blue** - Professional blue tones |
| `theme-emerald.css` | **Emerald Teal** - Fresh natural feel |
| `theme-clean.css` | **Clean White** - Minimal light header |
| `theme-slate.css` | **Slate** - Neutral professional tones |
| `gridsheet.lang.js` | Custom language/translation |

---

## How to Use Themes

### Simple: Just 2 CSS Files!

```html
<link rel="stylesheet" href="datatables-gridsheet/css/gridsheet.css">
<link rel="stylesheet" href="datatables-gridsheet/custom/theme-ocean.css">
```

Each theme file includes:
- Color variables
- Modern table design (rounded corners, shadows)
- Zebra striping
- Dark mode support
- Responsive design
- Smooth animations

---

## Theme Preview

| Theme | Primary Color | Best For |
|-------|---------------|----------|
| **Classic** | `#34a853` | Default, Google Sheets users |
| **Ocean** | `#0ea5e9` | Corporate, Professional |
| **Emerald** | `#10b981` | Finance, Success metrics |
| **Clean** | `#94a3b8` | Minimal, Light UI |
| **Slate** | `#64748b` | Neutral, Dark header |

---

## Switching Themes

Just change the theme file name:

```html
<!-- Classic Green -->
<link rel="stylesheet" href="datatables-gridsheet/custom/theme-classic.css">

<!-- Ocean Blue -->
<link rel="stylesheet" href="datatables-gridsheet/custom/theme-ocean.css">

<!-- Emerald Teal -->
<link rel="stylesheet" href="datatables-gridsheet/custom/theme-emerald.css">

<!-- Clean White -->
<link rel="stylesheet" href="datatables-gridsheet/custom/theme-clean.css">

<!-- Slate -->
<link rel="stylesheet" href="datatables-gridsheet/custom/theme-slate.css">
```

---

## Dark Mode

All themes include automatic dark mode support:
- **Auto-detect**: Follows system preference (`prefers-color-scheme: dark`)
- **Manual toggle**: Add `data-theme="dark"` to `<html>` tag

---

## Custom Language

Include after `gridsheet.js`:

```html
<script src="datatables-gridsheet/js/gridsheet.js"></script>
<script src="datatables-gridsheet/custom/gridsheet.lang.js"></script>
```

Use in options:

```javascript
const gridsheet = new GridSheet({
    language: GridSheetLang.id,  // Indonesian
    // ... other options
});
```

### Available Languages

| Code | Language |
|------|----------|
| `GridSheetLang.en` | English (default) |
| `GridSheetLang.id` | Indonesian |

Add other languages in `gridsheet.lang.js` using the provided template.
