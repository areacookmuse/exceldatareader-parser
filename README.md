<h1 align="center">ExcelDataReader</h1>

<p align="center">
  Lightweight C# library for reading Microsoft Excel files, including XLS, XLSX, XLSB, and CSV.
</p>

<p align="center">
  <a href="https://www.nuget.org/packages/ExcelDataReader">
    <img src="https://img.shields.io/nuget/v/ExcelDataReader.svg" alt="NuGet">
  </a>
</p>

## 📘 Overview

ExcelDataReader is a lightweight and fast library written in C# for reading Microsoft Excel files. It supports Excel formats from Excel 2.0 through Excel 2021 and Microsoft 365, including legacy binary workbooks, modern OpenXML workbooks, binary OpenXML workbooks, and CSV files.

The library is designed for scenarios where applications need to read spreadsheet data programmatically without requiring Excel itself. It provides a low-level reader interface for direct row-by-row access and an optional DataSet extension for loading spreadsheet content into `System.Data.DataSet`.

ExcelDataReader is split into two NuGet packages:

| Package | Purpose | Target frameworks |
|---|---|---|
| `ExcelDataReader` | Base package with the low-level reader interface | `net462`, `netstandard2.0`, `netstandard2.1` |
| `ExcelDataReader.DataSet` | Extension package providing `AsDataSet()` | `net462`, `netstandard2.0`, `netstandard2.1` |

## ✅ Supported File Formats

| File Type | Container Format | File Format | Excel Version(s) |
|---|---|---|---|
| `.xlsx` | ZIP, CFB+ZIP | OpenXml | 2007 and newer |
| `.xlsb` | ZIP, CFB | OpenXml | 2007 and newer |
| `.xls` | CFB | BIFF8 | 97, 2000, XP, 2003; 98, 2001, v.X, 2004 for Mac |
| `.xls` | CFB | BIFF5 | 5.0, 95 |
| `.xls` | - | BIFF4 | 4.0 |
| `.xls` | - | BIFF3 | 3.0 |
| `.xls` | - | BIFF2 | 2.0, 2.2 |
| `.csv` | - | CSV | All |

## 📦 Packages

Use NuGet to reference the package that matches your needs.

For the low-level reader API:

```powershell
Install-Package ExcelDataReader
```

For `AsDataSet()` support:

```powershell
Install-Package ExcelDataReader.DataSet
```

The `ExcelDataReader.DataSet` package depends on the base `ExcelDataReader` package.

## 🚀 Basic Usage

The reader can auto-detect supported Excel formats from a stream.

```csharp
using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
{
    using (var reader = ExcelReaderFactory.CreateReader(stream))
    {
        do
        {
            while (reader.Read())
            {
                // Example:
                // var value = reader.GetValue(0);
            }
        } while (reader.NextResult());
    }
}
```

To load workbook data into a `DataSet`, reference `ExcelDataReader.DataSet` and use `AsDataSet()`:

```csharp
using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
{
    using (var reader = ExcelReaderFactory.CreateReader(stream))
    {
        var result = reader.AsDataSet();

        // Spreadsheet data is available in result.Tables
    }
}
```

## 📄 Reading CSV Files

For plain text comma-separated values, use `CreateCsvReader`:

```csharp
using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
{
    while (reader.Read())
    {
        var value = reader.GetValue(0);
    }
}
```

CSV behavior includes:

- CSV field values are returned as strings.
- The caller is responsible for interpreting numbers, dates, or other types.
- `FallbackEncoding` can be used when the CSV lacks a BOM and does not parse as UTF-8.
- `AutodetectSeparators` can be used to configure separator detection.
- The CSV input is parsed in advance to determine `FieldCount`, `RowCount`, encoding, and separator.
- `AnalyzeInitialCsvRows` can limit CSV pre-analysis when full upfront row counting is not required.

## 🧭 Reader API

`IExcelDataReader` extends `System.Data.IDataReader` and `IDataRecord`, making it suitable for row-by-row reading and direct access to cell values.

Common members include:

| Member | Description |
|---|---|
| `Read()` | Reads a row from the current sheet |
| `NextResult()` | Moves to the next sheet |
| `ResultsCount` | Returns the number of sheets in the workbook |
| `Name` | Returns the current sheet name |
| `CodeName` | Returns the VBA code name identifier for the current sheet |
| `FieldCount` | Returns the number of columns in the current sheet |
| `RowCount` | Returns the number of rows in the current sheet |
| `HeaderFooter` | Returns header and footer information, or `null` |
| `MergeCells` | Returns merged cell ranges |
| `RowHeight` | Returns the visual row height in points |
| `GetColumnWidth()` | Returns the column width in character units |
| `GetFieldType()` | Returns the value type for the current row and column |
| `IsDBNull()` | Checks whether a value is null |
| `GetValue()` | Returns a cell value as `object` |
| `GetDouble()` | Returns a value as `double` |
| `GetInt32()` | Returns a value as `int` |
| `GetBoolean()` | Returns a value as `bool` |
| `GetDateTime()` | Returns a value as `DateTime` |
| `GetString()` | Returns a value as `string` |
| `GetNumberFormatString()` | Returns the number format string for a cell |
| `GetNumberFormatIndex()` | Returns the number format index for a cell |
| `GetCellStyle()` | Returns style information for a cell |

Typed `Get*()` methods throw `InvalidCastException` unless the value type matches exactly.

## ⚙️ Reader Configuration

Reader factory methods accept an optional `ExcelReaderConfiguration` object.

```csharp
var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
{
    FallbackEncoding = Encoding.GetEncoding(1252),
    Password = "password",
    AutodetectSeparators = new char[] { ',', ';', '\t', '|', '#' },
    QuoteChar = '"',
    TrimWhiteSpace = true,
    LeaveOpen = false,
    AnalyzeInitialCsvRows = 0,
    SinglePassMode = false,
    EscapeChar = null,
});
```

Configuration options include:

| Option | Description |
|---|---|
| `FallbackEncoding` | Encoding used when an XLS file lacks a CodePage record, or when CSV lacks a BOM and does not parse as UTF-8 |
| `Password` | Password used to open password-protected workbooks |
| `AutodetectSeparators` | CSV separator candidates |
| `QuoteChar` | CSV quote character |
| `TrimWhiteSpace` | Controls whether CSV whitespace values are trimmed |
| `LeaveOpen` | Controls whether the stream remains open after disposing the reader |
| `AnalyzeInitialCsvRows` | Limits CSV analysis to a specified number of rows |
| `SinglePassMode` | Skips the initial full scan for XLS and XLSX/XLSB files |
| `EscapeChar` | Optional escape character for CSV quoted fields |

## 🧾 DataSet Configuration

`AsDataSet()` accepts an `ExcelDataSetConfiguration` object for controlling how workbook data is converted into tables.

```csharp
var result = reader.AsDataSet(new ExcelDataSetConfiguration()
{
    UseColumnDataType = true,

    FilterSheet = (tableReader, sheetIndex) => true,

    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
    {
        CaseSensitive = false,
        EmptyColumnNamePrefix = "Column",
        UseHeaderRow = false,

        ReadHeaderRow = (rowReader) =>
        {
            rowReader.Read();
        },

        FilterRow = (rowReader) => true,
        FilterColumn = (rowReader, columnIndex) => true,
        FillMergedCellsValue = false,
    }
});
```

Useful options include:

| Option | Description |
|---|---|
| `UseColumnDataType` | Sets `DataColumn.DataType` in a second pass |
| `FilterSheet` | Controls whether a sheet is included |
| `ConfigureDataTable` | Provides per-table configuration |
| `CaseSensitive` | Sets whether the resulting `DataTable` is case-sensitive |
| `EmptyColumnNamePrefix` | Prefix for generated column names |
| `UseHeaderRow` | Uses a row from the data as column names |
| `ReadHeaderRow` | Selects the row to use as headers |
| `FilterRow` | Controls whether a row is included |
| `FilterColumn` | Controls whether a column is included |
| `FillMergedCellsValue` | Fills merged cells with the top-left cell value |

## 🎨 Formatting Values

ExcelDataReader does not apply cell formatting directly. It can expose Excel number format information so callers can decide how to format values.

Use:

```csharp
reader.GetNumberFormatString(columnIndex);
reader.GetNumberFormatIndex(columnIndex);
```

For formatting values, the original documentation references the third-party `ExcelNumberFormat` library.

Example:

```csharp
string GetFormattedValue(IExcelDataReader reader, int columnIndex, CultureInfo culture)
{
    var value = reader.GetValue(columnIndex);
    var formatString = reader.GetNumberFormatString(columnIndex, culture);

    if (formatString != null)
    {
        var format = new NumberFormat(formatString);
        return format.Format(value, culture);
    }

    return Convert.ToString(value, culture);
}
```

Related links:

- https://github.com/andersnm/ExcelNumberFormat
- https://www.nuget.org/packages/ExcelNumberFormat

## 🧩 .NET Core and Encoding Notes

On .NET Core and .NET 5.0 or later, ExcelDataReader may throw:

```text
No data is available for encoding 1252.
```

To resolve this, add a dependency on `System.Text.Encoding.CodePages` and register the provider during application initialization:

```csharp
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
```

This is required for parsing strings in binary BIFF2-5 Excel documents encoded with older DOS-era code pages. These encodings are available by default in the full .NET Framework, but not in .NET Core or .NET 5.0 and later.

## 🔄 Upgrading from ExcelDataReader 2.x

ExcelDataReader 3 introduced breaking changes. Older code may produce errors such as:

```text
'IExcelDataReader' does not contain a definition for 'AsDataSet'
'IExcelDataReader' does not contain a definition for 'IsFirstRowAsColumnNames'
```

To update older code:

1. Rename namespace references from `Excel` to `ExcelDataReader`.
2. Reference `ExcelDataReader.DataSet` when using `AsDataSet()`.
3. Replace `IsFirstRowAsColumnNames` with `ExcelDataSetConfiguration`.

Example:

```csharp
var result = reader.AsDataSet(new ExcelDataSetConfiguration()
{
    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
    {
        UseHeaderRow = true
    }
});
```

## 🧪 Continuous Integration

The project uses GitHub Actions for continuous integration.

[![CI](https://github.com/ExcelDataReader/ExcelDataReader/actions/workflows/ci.yml/badge.svg?branch=develop)](https://github.com/ExcelDataReader/ExcelDataReader/actions/workflows/ci.yml)

## ❓ FAQ

### What file types can ExcelDataReader read?

ExcelDataReader supports `.xls`, `.xlsx`, `.xlsb`, and `.csv` files. It supports Excel formats ranging from Excel 2.0 through Excel 2021 and Microsoft 365.

### Does ExcelDataReader require Microsoft Excel?

The README describes ExcelDataReader as a C# library for reading Excel files. It does not state that Microsoft Excel is required.

### When should I use `ExcelDataReader.DataSet`?

Use `ExcelDataReader.DataSet` when you want the `AsDataSet()` extension method to load spreadsheet data into a `System.Data.DataSet`.

### Are CSV values automatically converted to numbers or dates?

No. CSV field values are returned as strings. The caller is responsible for interpreting CSV values.

### Does ExcelDataReader apply Excel cell formatting?

No. ExcelDataReader does not support formatting directly, but it can expose number format strings and format indices so callers can handle formatting separately.

### What should I provide when reporting an issue?

When reporting an issue, it is useful to supply an example Excel file because it makes debugging easier. Without an example file, some problems may be difficult to resolve.

## 🤝 Contributing

Forks and pull requests are welcome. The original README asks contributors to submit pull requests to the `develop` branch.

When reporting issues, include an example Excel file when possible. This helps reproduce and debug format-specific problems.
