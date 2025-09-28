using System.Text.Json;

namespace Itenium.ExcelCreator.Library.Models;

/// <summary>
/// Both the data to be put in the Excel and the
/// configuration on how to format the data.
/// </summary>
public class FullExcelData
{
    public JsonElement[][] Data { get; set; } = [];
    public ExcelConfiguration Config { get; set; } = new();

    public override string ToString() => $"Rows={Data.Length}, Config={Config}";
}

public class ExcelConfiguration
{
    public string FileName { get; set; } = "";
    public string SheetName { get; set; } = "Sheet1";
    public ColumnConfiguration[] Columns { get; set; } = [];

    public override string ToString() => $"{SheetName}: {string.Join(", ", Columns.Select(x => x.Header))}";
}

public class ColumnConfiguration
{
    /// <summary>
    /// The label for the column in the top row
    /// </summary>
    public string Header { get; set; } = "";
    public ColumnType Type { get; set; }
    /// <summary>
    /// Can start with '='.
    /// Use {row} to be replaced with the current row number.
    /// Example: "=A{row}+B{row}".
    /// </summary>
    public string? Formula { get; set; }

    public override string ToString() => $"{Header}: Type={Type}, Formula={Formula?.TrimStart('=')}";
}

/// <summary>
/// How to format and display an Excel cell.
/// 
/// If there is an issue parsing, it will revert to
/// trying to just display it (ie numbers remain
/// numbers, ... or display as a string)
/// </summary>
public enum ColumnType
{
    String,
    /// <summary>
    /// mm/dd/yyyy
    /// </summary>
    Date,
    /// <summary>
    /// Value passed should be 0-100.
    /// Formatting: 0.00%.
    /// </summary>
    Percentage,
    /// <summary>
    /// #,##0
    /// </summary>
    Integer,
    /// <summary>
    /// € #,##0.00
    /// </summary>
    Money,
    /// <summary>
    /// #,##0.00
    /// </summary>
    Decimal,
    /// <summary>
    /// true/false
    /// </summary>
    Boolean,
}
