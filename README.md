# PO-Region-Classifier
Excel VBA tool that classifies Purchase Orders into global regions (AMER / APAC / EMEA) using prefix-based mapping. Includes sample data and modular .bas files.

# PO Region Classifier (Excel + VBA)

This project provides an Excel-based automation tool that classifies Purchase Order values into global regions (**AMER / APAC / EMEA**) using prefix-based lookup logic. The macros read PO identifiers, assign regions automatically, and generate summary counts without manual filtering or pivot tables.

This repository contains **a fully sanitized demo version** with dummy sample data and standalone `.bas` files that can be imported into any workbook.

---

## 🚀 Features

- 🔹 Classifies POs into **AMER / APAC / EMEA**
- 🔹 Uses a **country prefix → region mapping sheet**
- 🔹 Outputs **region tags in Column B**
- 🔹 Generates **summary counts automatically**
- 🔹 Includes a macro to **clear previous results**
- 🔹 Lightweight, no external dependencies

---

## 📂 Repository Structure

PO-Region-Classifier/
├── src/
│ ├── po_region_classifier.bas
│ ├── clear_po.bas
├── data/
│ └── PO_Mapping_sheet.xlsx
├── demo/
│ └── demo-po-classification.gif (coming soon)
└── README.md

## 🔧 How to Use

1. Open a new Excel workbook.
2. Press **Alt + F11** to open the VBA editor.
3. Go to **File → Import File…**
4. Import:
   - `po_region_classifier.bas`
   - `clear_po.bas`
5. Add PO numbers in **Sheet `POData`**, Column A.
6. Add the mapping sheet as **`POMappings`** (or rename accordingly).
7. Run the macro:

**From Excel:**
- `Alt + F8` → `Classify_ByRegion_ApacSet`

or attach to a button for quick access.

---

## 🧪 Example Output

| PO Number | Region |
|-----------|--------|
| AM10001   | AMER   |
| CN84010   | APAC   |
| FR65020   | EMEA   |

Auto-generated totals:

| Region | Count |
|--------|--------|
| AMER   | 77     |
| APAC   | 15     |
| EMEA   | 20     |
| Total  | 112    |

---

## 🗂 Data Included

The repository includes a sample region mapping sheet:

| Prefix | Country | GlobalRegion |
|--------|----------|--------------|
| US     | United States | AMER |
| IN     | India | APAC |
| DE     | Germany | EMEA |
| …      | …         | …

This dataset is **generic and non-confidential**.

---

## 🛠 Skills Demonstrated

- Excel automation in VBA  
- Dictionary-based classification logic  
- Data transformation for operations reporting  
- Modular code structure & reusable macros  
- Clean demo-based portfolio design

---

## 👤 Author

*Madhumitha Sekar*  
Practical automation projects for operations & procurement workflows.
