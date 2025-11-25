# ArcGIS-Pro-Custom-Geoprocessing-Grouping-Excel-Export
A custom geoprocessing tool for ArcGIS Pro that has options for grouping data when exported to an Excel document.

super_excel_export/
├─ tool_entry.py           # ArcGIS Pro script tool entry point
├─ super_export/
│  ├─ __init__.py          # Package initializer
│  ├─ core.py              # run_export() and overall workflow
│  ├─ data_utils.py        # Layer reading, fields, domains, DataFrame ops
│  └─ excel_utils.py       # Sheet naming, TOC, formatting, Excel writing



This project contains a multi-file Python implementation of an advanced ArcGIS Pro geoprocessing tool for exporting feature layer data to Excel.  

The tool supports:

- Exporting all fields or a custom field list
- Optional domain code → description** mapping
- Optional field alias output
- Single-sheet or multi-sheet grouped Excel export
- Automatic column width adjustment
- Freeze header row option
- TOC (Table of Contents) sheet creation
- Full compatibility with ArcGIS Pro script tools



---

Workflow

```mermaid
flowchart TD
    A[Start: GP tool runs] --> B[Read tool parameters]
    B --> C[Resolve input layer & field lists]
    C --> D[Detect selection and read features into DataFrame]
    D --> E{Use coded domain descriptions?}

    E -->|Yes| F[Get domain mappings<br/>(GDB or service)]
    F --> G[Apply domain mappings<br/>codes → descriptions]
    E -->|No| H[Skip domain mapping]

    G --> I[Clean column headers<br/>(aliases or raw names)]
    H --> I[Clean column headers<br/>(aliases or raw names)]

    I --> J[Ensure group-by column exists<br/>and is normalized]

    J --> K{Export mode?}
    K -->|single_sheet| L[Sort DataFrame<br/>by group and other fields]
    L --> M[Write single 'Data' sheet]

    K -->|sheets| N[Group DataFrame by group-by field]
    N --> O[For each group:<br/>build safe, unique sheet name<br/>and write sheet]

    M --> P{TOC enabled?}
    O --> P{TOC enabled?}

    P -->|Yes| Q[Build TOC sheet<br/>with sheet names & counts]
    P -->|No| R[Skip TOC]

    Q --> S[Apply formatting<br/>(auto-width, freeze header)]
    R --> S[Apply formatting<br/>(auto-width, freeze header)]

    S --> T[Set Excel path as output<br/>and log success]
    T --> U[End]
