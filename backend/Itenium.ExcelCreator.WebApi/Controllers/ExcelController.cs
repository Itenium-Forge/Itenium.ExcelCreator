using ClosedXML.Excel;
using Itenium.ExcelCreator.Client;
using Itenium.ExcelCreator.WebApi.Helpers;
using Microsoft.AspNetCore.Mvc;

namespace Itenium.ExcelCreator.WebApi.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ExcelController : ControllerBase, IExcelService
{
    private readonly ILogger<ExcelController> _logger;

    public ExcelController(ILogger<ExcelController> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Creates an Excel from the data and configuration posted in the body
    /// </summary>
    /// <returns>The generated Excel</returns>
    [HttpPost]
    public FileStreamResult Create([FromBody] FullExcelData data)
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(data.Config.SheetName);

        int cols = data.Config.Columns.Length;

        if (data.Config.Columns.Length > 0)
        {
            for (int i = 0; i < data.Config.Columns.Length; i++)
            {
                var cell = ws.Cell(1, i + 1);
                cell.Value = data.Config.Columns[i].Header;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
            }
        }

        // Add data rows
        for (int rowIndex = 0; rowIndex < data.Data.Length; rowIndex++)
        {
            var row = data.Data[rowIndex];
            for (int colIndex = 0; colIndex < row.Length; colIndex++)
            {
                var cell = ws.Cell(rowIndex + 2, colIndex + 1); // +2 because row 1 is headers
                var value = row[colIndex];

                // Apply formatting based on column type
                if (colIndex < data.Config.Columns.Length)
                {
                    var columnConfig = data.Config.Columns[colIndex];
                    FormatCellBasedOnType(cell, value, columnConfig.Type);
                }
                else
                {
                    // Default formatting
                    cell.Value = value?.ToString();
                }
            }
        }

        // var lastColumn = Math.Max(, data.Data[0].Length);
        var lastRow = data.Data.Length + 1; // +1 for header row
        ws.Range(1, 1, lastRow, cols).SetAutoFilter();

        ws.ColumnsUsed().AdjustToContents();
        return wb.Deliver(data.Config.FileName);
    }

    private static void FormatCellBasedOnType(IXLCell cell, object? value, ColumnType columnType)
    {
        if (value == null)
        {
            cell.Value = "";
            return;
        }

        switch (columnType)
        {
            case ColumnType.String:
                cell.Value = value.ToString();
                break;

            case ColumnType.Date:
                if (DateTime.TryParse(value.ToString(), out DateTime dateValue))
                {
                    cell.Value = dateValue;
                    cell.Style.DateFormat.Format = "mm/dd/yyyy";
                }
                else
                {
                    cell.Value = value.ToString();
                }
                break;

            case ColumnType.Percentage:
                if (double.TryParse(value.ToString(), out double percentValue))
                {
                    // Assuming the value is already in percentage form (e.g., 25 for 25%)
                    cell.Value = percentValue / 100;
                    cell.Style.NumberFormat.Format = "0.00%";
                }
                else
                {
                    cell.Value = value.ToString();
                }
                break;

            default:
                cell.Value = value.ToString();
                break;
        }
    }

    public class FullExcelData
    {
        public object?[][] Data { get; set; } = [];
        public ExcelConfiguration Config { get; set; } = new();
    }

    public class ExcelConfiguration
    {
        public string FileName { get; set; } = "";
        public string SheetName { get; set; } = "";
        public ColumnConfiguration[] Columns { get; set; } = [];
    }

    public class ColumnConfiguration
    {
        public string Header { get; set; } = "";
        public ColumnType Type { get; set; }
    }

    public enum ColumnType
    {
        String,
        Date,
        Percentage,
        Number,
        Money,
    }
}
