# iOBR Extractor

Cross-platform desktop utility for splitting `OBR_Output.xlsx` into account-level iOBR workbooks.

## Run

```bash
python3 iobr_app.py
```

On Windows, use:

```bash
python iobr_app.py
```

## Workflow

1. Select the base `OBR_Output.xlsx` file.
2. Click `Load Filters`.
3. Select optional filters. Leaving a filter empty means all values.
4. Select the output folder.
5. Click `Generate`.

The app creates one folder per Sold-to and writes the extraction files inside it. On each new run, only the files for the selected account batches are overwritten. Existing Sold-to folders and files for accounts that are not part of the current run remain unchanged.

When Sold-to and D-Account are left empty or both set to all values, the app still generates separate report files per account and places them in the corresponding Sold-to folder. If one or more Sold-tos are selected, the D-Account selector is narrowed to only accounts connected to those Sold-tos.

## Current Report Rules

- `Line_Matched`: global filters plus `REJECT_QTY = 0`.
- `Country-of-Origin`: `COMMENT_COO = Check`.
- `IHDvsCRD`: `COMMENT_CRD_VS_IHD` is not `OK` / `On time or early`, and `REJECT_QTY = 0`.
- `Style-ID`: `COMMENT_PROD_CD = Check`, and `REJECT_QTY = 0`.
- `Cancellations`: `COMMENT_QUANTITY_PO_LEVEL = Check rejection`, and `REPORTING_CONFIRMED_QTY_PO_LEVEL = 0`. `Nike SCM comment` is not generated.
- `PO not received - Cancellations`: only `D_ACCOUNT` filter is applied, plus `COMMENTS = Check OFOA status`.
- `Qty by Size`: size-level quantity issue comments, and `REJECT_QTY = 0`.
- `Nike Data Only`: global filters available on `Line_Nike_Data_Only`, excluding rows where `SO_STATUS_L2_DESC` contains `Cancel`.

`9897` is currently fixed in the generated filenames.

## CLI

The same extractor can run without the UI:

```bash
python3 iobr_app.py --input data/OBR_Output.xlsx --output output --season HO2026 --d-account GER-38
```

## Build Windows Executable

Windows executables must be built on Windows. The repository includes a GitHub Actions workflow that builds `iOBR Extractor.exe` automatically on every push to `main`, and can also be run manually from the GitHub `Actions` tab.

Download path in GitHub:

1. Open the repository on GitHub.
2. Go to `Actions`.
3. Open the latest `Build Windows App` run.
4. Download the `iOBR-Extractor-Windows` artifact.
5. Unzip it and run `iOBR Extractor.exe`.

To build directly on a Windows laptop with Python installed, run PowerShell from the project folder:

```powershell
.\build_windows.ps1
```
