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
        pandas>=2.2.3
        numpy>=2.0.0,<3.0.0
        openpyxl>=3.1.5
        fuzzywuzzy==0.18.0
        tqdm>=4.67.1
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

## Module Compatibility

The following table lists the recommended Python and module versions for running the DeDup and DeDup_multi code.

| Dependency | Version |
| :--- | :--- |
| Python | 3.11 or 3.12+ |
| pandas | 2.2.3 or newer |
| NumPy | 2.x |
| openpyxl | 3.1.5 or newer |
| fuzzywuzzy | 0.18.0 |
| tqdm | 4.67.1 or newer |


### Compatibility Notes

- The updated code uses `DataFrame.map()` instead of the deprecated `DataFrame.applymap()`. This requires pandas 2.1.0 or newer.
- The updated code replaces the private `DataFrame._append()` method with `pandas.concat()`.


Use the following command to execute the program: * The tool is OS independent and execute python code using corresponding command. Below is for running in windows OS

```bash
python DeDup.py -f1 q.xlsx -f2 t.xlsx -j TEST
or
python DeDup_multi.py -f1 q.xlsx -f2 t.xlsx -j TEST -w 8
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

The file S_column.xlsx is used for column mapping. This table defines the standard column names, their data types, and the customizable equivalents used during processing.

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

## Disclaimer

This program is distributed in the hope that it will be useful, but **WITHOUT ANY WARRANTY**; without even the implied warranty of **MERCHANTABILITY** or **FITNESS FOR A PARTICULAR PURPOSE**.  
See the [GNU General Public License](https://www.gnu.org/licenses/gpl-3.0.html) for more details.

The authors shall not be held liable for any direct, indirect, incidental, special, exemplary, or consequential damages arising in any way out of the use of this software.
