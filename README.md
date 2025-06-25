# De-Dup
A Fuzzy Token Sort Ratio-Based Method for Handling Naming and Address Diversity in Deduplication

---

## ⚙️ Requirements

This tool requires:

- **Python version 3.11.4 or higher**

To check your Python version:

```bash
python --version
```

If needed, download the latest version from: [https://www.python.org/downloads/](https://www.python.org/downloads/)

---

## 📦 Installing Python Dependencies

After installing Python, you need to install a few external modules. Other required libraries are part of the Python Standard Library.

### ✅ Install All Required Modules

```bash
pip install pandas numpy fuzzywuzzy tqdm
```

> Modules like `sys`, `re`, `math`, `time`, `argparse`, `os`, `shutil`, and `xml.etree.ElementTree` are part of the **Python standard library** and do not require installation.

### Optional: Using `requirements.txt`

If you prefer using a dependency file:

1. Create a file named `requirements.txt` with the following contents:

    ```txt
    pandas
    numpy
    fuzzywuzzy
    tqdm
    ```

2. Install all modules at once:

    ```bash
    pip install -r requirements.txt
    ```

---

## 📁 Setup Instructions

1. Install Python and dependencies (as above).
2. **Download all repository files** into a single folder. Make sure the following files are present:
   - `DeDup.py`
   - `config.txt`
   - `ICD.txt`
   - Any additional required files

---

## ▶️ Running the Tool

Use the following command to execute the program:

```bash
python DeDup.py -f1 q.xlsx -f2 t.xlsx -j TEST
```

### 🔹 Argument Details:

| Argument | Description                                | Required |
|----------|--------------------------------------------|----------|
| `-f1`    | First input Excel file for comparison       | ✅       |
| `-f2`    | Second input Excel file for comparison      | ✅       |
| `-j`     | Job name (used for output and tracking)     | ✅       |

> All three arguments are **mandatory**.

---

## 🧾 Column Mapping Table

This table defines the standard column names, their data types, and the customizable equivalents used during processing.

| VARIABLE      | DATA_TYPE | EQU_C_NAME   |
|---------------|-----------|--------------|
| REF           | UAN       | REF          |
| FULL_NAME     | AN        | FULL_NAME    |
| FULL_ADDRESS  | AN        | FULL_ADDRESS |
| AGE           | N         | AGE          |
| GENDER        | AN        | GENDER       |
| ICD           | AN        | ICD          |
| PINCODE       | N         | PINCODE      |
| RELATIVE      | AN        | RELATIVE     |

### 📝 Notes:

- Columns `VARIABLE` and `DATA_TYPE` must **not be modified** unless you plan to update the Python code logic.
- You may **modify the `EQU_C_NAME` values** to match your input Excel column names.
- Example: If your Excel file uses `Person_Name` instead of `FULL_NAME`, update `EQU_C_NAME` accordingly.

---

## 🧩 Data Type Legend

- **UAN** – Unique Alphanumeric  
- **AN** – Alphanumeric  
- **N** – Numeric  

---

## ⚙️ Configuration File: `config.txt`

The `config.txt` file defines threshold-based decision rules in an XML-like format. These are used within nested conditions in the `main()` function.

### Example Content:

```xml
<BASE_CONDITION>
    <C0>THRESHOLD:0.0</C0>
    <C1>FULL_NAME:75</C1>
    <C1a>FULL_NAME:50</C1a>
    <C1b>ICD:1</C1b>
    ...
    <C5e>TOTAL_SCORE:105</C5e>
</BASE_CONDITION>
```

### Key Notes:

- `C0` defines a **global threshold** (default `0.0`). You can increase this if needed.
- Other tags (e.g., `<C1>`, `<C3b6>`) represent custom thresholds used within matching logic.
- Edit values **only if you fully understand** how they affect rule evaluation in the `main()` function.

---

## 🧾 ICD Code File: `ICD.txt`

This file contains a list of **ICD-10 cancer codes** considered equivalent for matching purposes.

### ✅ Behavior

- If one record has ICD `C10` and another has `C26`, and `C26` is listed in `ICD.txt`, they are treated as a **match**.
- You may **add more cancer-related ICD-10 codes** if you want them to be treated as matchable.

### 🛑 Disabling ICD Matching

If you don't want any ICD equivalency logic:

1. **Do not leave the file empty** (this causes errors).
2. Instead, write the following on the **first line**:

   ```
   XXX
   ```

This disables ICD equivalence matching — but:

> Records with the **same ICD code** (e.g., `C26` in both records) will still be treated as a match.

---


## 📤 Output Description

After successfully running the tool, the program produces **four output files** in the working directory based on the provided job name.

### 📁 Output Files:

| File Name       | Description                                                                 |
|------------------|-----------------------------------------------------------------------------|
| `QC.xlsx`        | Cleaned version of **Input File 1** (`-f1`)                                 |
| `TC.xlsx`        | Cleaned version of **Input File 2** (`-f2`)                                 |
| `results.xlsx`   | Final result file with **matched record pairs** and corresponding **match scores** |
| `score.xlsx`     | Contains only the **score summary** (probability values) for each matched pair |

---

### 📘 File Details

#### 🔹 `results.xlsx`

- Contains matched pairs of records from both input files.
- Each match includes two rows of data (one from each file) followed by a third row that shows the **match scores for each variable** used in comparison.

#### 🔹 `score.xlsx`

- Contains a **summary view** of the matching results.
- Includes unique identifiers from both input files along with a **combined match probability score** for each matched pair.
- This file can be used for **filtering high-confidence matches** based on probability thresholds.

## 🧰 Support

For issues, questions, or suggestions, please open an issue in this repository or contact saravanan.vij@icmr.gov.in or brsaran@gmail.com.
