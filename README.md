# Sales Funnel Search Web Part

A SharePoint Framework (SPFx) web part that provides a rich search experience for both **documents** and **list items** within a SharePoint site or scoped to a specific list/library.

---

## ✨ Features

### Search UI
- **Search bar** — full-text search input with keyboard `Enter` support
- **3 checkboxes** — dynamically filter which content to search:
  - ☑ Documents
  - ☑ List Items
  - ☑ Both
- **Search button** — triggers the search
- **Results area** — shows results with title, metadata, and a snippet

### Tabbed Results (`Both` mode)
When both content types are enabled, results appear in three tabs:
| Tab | Content |
|---|---|
| **All** | Documents + List Items combined |
| **Documents** | Document results only |
| **List Items** | List item results only |

### Property Pane Configuration

| Setting | Description |
|---|---|
| **Display Mode** | Choose `Documents only`, `List Items only`, or `Both` |
| **Search Scope** | `Entire Site` or `Specific List/Library URL` |
| **Site URL** | (when scope = Site) Searches all lists & libraries in the site |
| **List/Library URL** | (when scope = URL) Searches only the specified target |

---

## 🗂 Project Structure

```
salesFunnelSearch/
├── .github/
│   └── workflows/
│       └── build.yml                         ← CI/CD pipeline
├── config/
│   ├── config.json
│   ├── package-solution.json
│   ├── serve.json
│   └── write-manifests.json
├── src/
│   └── webparts/
│       └── salesFunnelSearch/
│           ├── components/
│           │   ├── ISalesFunnelSearchProps.ts  ← Component props interface
│           │   ├── SalesFunnelSearch.module.scss      ← Styles
│           │   ├── SalesFunnelSearch.module.scss.ts   ← SCSS typings
│           │   └── SalesFunnelSearch.tsx       ← Main React component
│           ├── models/
│           │   └── ISalesFunnelItem.ts         ← Data models & types
│           ├── services/
│           │   └── SalesFunnelSearchService.ts ← SharePoint REST API calls
│           ├── SalesFunnelSearchWebPart.manifest.json
│           └── SalesFunnelSearchWebPart.ts     ← Web part entry point
├── gulpfile.js
├── package.json
└── tsconfig.json
```

---

## 🚀 Getting Started

### Prerequisites
- Node.js 18.x
- SharePoint Framework toolchain (`@microsoft/generator-sharepoint`)

### Install & Run

```bash
# 1. Clone the repository
git clone https://github.com/your-org/salesFunnelSearch.git
cd salesFunnelSearch

# 2. Install dependencies
npm install

# 3. Trust dev certificate (first time only)
gulp trust-dev-cert

# 4. Update serve.json with your SharePoint site workbench URL
#    config/serve.json → "initialPage"

# 5. Serve locally
gulp serve
```

### Build for Production

```bash
gulp bundle --ship
gulp package-solution --ship
```

The `.sppkg` package will be at `sharepoint/solution/sales-funnel-search.sppkg`.

### Deploy to SharePoint

1. Upload `sales-funnel-search.sppkg` to your **App Catalog**.
2. Deploy it tenant-wide or to specific sites.
3. Add the web part to a modern SharePoint page.
4. Open the **property pane** (edit the page → click the pencil on the web part) to configure:
   - Display mode
   - Search scope
   - Site or List/Library URL

---

## 🔧 Configuration Examples

### Search an entire site
| Property | Value |
|---|---|
| Search scope | `Entire Site` |
| Site URL | `https://contoso.sharepoint.com/sites/Sales` |

### Search a single document library
| Property | Value |
|---|---|
| Search scope | `Specific List or Library URL` |
| List/Library URL | `https://contoso.sharepoint.com/sites/Sales/Shared Documents` |

### Search a single list
| Property | Value |
|---|---|
| Search scope | `Specific List or Library URL` |
| List/Library URL | `https://contoso.sharepoint.com/sites/Sales/Lists/Opportunities` |

---

## 🛠 Architecture

### `SalesFunnelSearchWebPart.ts`
The SPFx web part entry point. Renders the React component and defines the property pane with two configuration groups:
- **Display Options** — controls which result types are shown and whether tabs appear
- **Search Scope** — controls whether to search a full site or a specific list/library URL

### `SalesFunnelSearch.tsx`
The main React component. Manages:
- Search input and checkbox state
- Calling `SalesFunnelSearchService`
- Rendering loading, empty, and result states
- Tab switching when `displayMode === 'both'`

### `SalesFunnelSearchService.ts`
Calls the SharePoint Search REST API (`/_api/search/query`) with:
- Full-text query
- Path filter scoped to the configured site or URL
- `contentclass:STS_ListItem` filter for list items
- Documents source ID for document results

### `ISalesFunnelItem.ts`
Shared data models: `ISalesFunnelItem`, `DisplayMode`, `SearchScope`, `ISearchResult`.

---

## 🤝 Contributing

1. Fork the repo
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes
4. Push and open a Pull Request

---

## 📄 License

MIT
