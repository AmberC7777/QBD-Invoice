# QBD Invoice Importer

Console tool for importing invoice headers and lines from CSV files into QuickBooks Desktop using the QBFC16 SDK.

## Requirements

- Windows machine with QuickBooks Desktop installed.
- QBFC16 SDK (COM reference `QBFC16Lib`) registered on the machine.
- .NET Framework 4.7.2 build environment.
- NuGet packages restored (see `packages.config`).

## Input CSVs

The app reads two CSV files from the working directory when you run the executable:

- `InvoiceHeader.csv`
- `InvoiceLines.csv`

### `InvoiceHeader.csv`

Required columns (header row must match property names):

| Column | Type | Notes |
| --- | --- | --- |
| `InvoiceID` | string | Correlates to line items in `InvoiceLines.csv`. |
| `CustomerRef` | string | QuickBooks customer full name. |
| `TxnDate` | date | Invoice transaction date. |
| `RefNumber` | string | Invoice reference number. |

### `InvoiceLines.csv`

Required columns (header row must match property names):

| Column | Type | Notes |
| --- | --- | --- |
| `InvoiceID` | string | Must match a header `InvoiceID`. |
| `LineNum` | integer | Ordering or line identifier. |
| `ItemRef` | string | QuickBooks item full name. |
| `Desc` | string | Optional description. |
| `Quantity` | decimal | Optional; ignored for percent-based items. |
| `Rate` | decimal | Optional; sent as `RatePercent` when `ItemRef` is `OOP`. |
| `Amount` | decimal | Optional; currently not mapped. |

> **Note**: When `ItemRef` equals `OOP` (case-insensitive), the `Rate` column is sent as a percent (e.g., `5` for 5%) rather than a standard rate.

## Build and Run

1. Restore NuGet packages.
2. Build the project in Visual Studio (or with MSBuild) targeting **.NET Framework 4.7.2** and **x86** (Debug).
3. Place `InvoiceHeader.csv` and `InvoiceLines.csv` next to the executable or run from the directory that contains them.
4. Run the executable; it will push invoices to QuickBooks Desktop.

## Output

- Success and error messages are printed to the console.
- If any invoices fail, details are written to `InvoiceImportErrors.log` in the working directory.

## Troubleshooting

- Ensure QuickBooks Desktop is running and that the application is authorized in QuickBooks.
- Verify that item and customer names match exactly what is in QuickBooks.
- Confirm the QBFC16 SDK is installed and registered.
