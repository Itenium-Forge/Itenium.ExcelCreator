using ClosedXML.Excel;
using Itenium.ExcelCreator.Library.Models;
using System.Globalization;
using System.Text.Json;

namespace Itenium.ExcelCreator.Library;

public class ExcelService
{
    public XLWorkbook CreateExcel(FullExcelData data)
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(data.Config.SheetName);

        for (int i = 0; i < data.Config.Columns.Length; i++)
        {
            var cell = ws.Cell(1, i + 1);
            cell.Value = data.Config.Columns[i].Header;
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
        }

        for (int rowIndex = 0; rowIndex < data.Data.Length; rowIndex++)
        {
            var row = data.Data[rowIndex];
            for (int colIndex = 0; colIndex < Math.Max(row.Length, data.Config.Columns.Length); colIndex++)
            {
                var cell = ws.Cell(rowIndex + 2, colIndex + 1);

                ColumnConfiguration? columnConfig = null;
                if (colIndex < data.Config.Columns.Length)
                {
                    columnConfig = data.Config.Columns[colIndex];
                    FormatCell(cell, columnConfig.Type);
                }

                if (!string.IsNullOrWhiteSpace(columnConfig?.Formula))
                {
                    string formula = columnConfig.Formula.Replace("{row}", cell.Address.RowNumber.ToString());
                    cell.FormulaA1 = formula;
                    
                    if (colIndex < row.Length && 
                        row[colIndex].ValueKind is not JsonValueKind.Null and not JsonValueKind.Undefined)
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.Red;
                        cell.CreateComment().AddText($"ERR: Both formula ({formula}) and data ({row[colIndex]})");
                    }
                }
                else if (colIndex < row.Length)
                {
                    var value = row[colIndex];
                    if (columnConfig == null)
                    {
                        cell.Value = GetStringValue(value);
                    }
                    else
                    {
                        SetCellValue(cell, value, columnConfig);
                    }
                }
            }
        }

        int lastRow = data.Data.Length + 1;
        int cols = data.Config.Columns.Length;
        ws.Range(1, 1, lastRow, cols).SetAutoFilter();

        if (data.Config.FreezeColumns is > 0)
        {
            ws.SheetView.FreezeColumns(data.Config.FreezeColumns.Value);
        }

        wb.RecalculateAllFormulas();

        ws.ColumnsUsed().AdjustToContents();
        return wb;
    }

    private static void FormatCell(IXLCell cell, ColumnType columnType)
    {
        switch (columnType)
        {
            case ColumnType.String:
                break;

            case ColumnType.Date:
                // cell.Style.DateFormat.NumberFormatId = (int)XLPredefinedFormat.DateTime.DayMonthYear4WithSlashes;
                cell.Style.DateFormat.Format = "dd/mm/yyyy";
                break;

            case ColumnType.Percentage:
                cell.Style.NumberFormat.Format = "0.00%";
                break;

            case ColumnType.Integer:
                // cell.Style.NumberFormat.NumberFormatId = (int)XLPredefinedFormat.Number.Precision2WithSeparator;
                cell.Style.NumberFormat.Format = "#,##0";
                break;

            case ColumnType.Money:
                cell.Style.NumberFormat.Format = "€ #,##0.00";
                break;

            case ColumnType.Decimal:
                cell.Style.NumberFormat.Format = "#,##0.00";
                break;

            case ColumnType.Boolean:
                break;

            default:
                throw new NotImplementedException($"Still have to implement ColumnType.{columnType}!");
        }
    }

    private static void SetCellValue(IXLCell cell, JsonElement value, ColumnConfiguration column)
    {
        if (value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            cell.Value = "";
            return;
        }

        switch (column.Type)
        {
            case ColumnType.String:
                cell.Value = GetStringValue(value);
                break;

            case ColumnType.Date:
                if (TryGetDateValue(value, out DateTime dateValue))
                {
                    cell.Value = dateValue;
                }
                else
                {
                    cell.Value = GetStringValue(value);
                }
                break;

            case ColumnType.Percentage:
                if (TryGetDecimalValue(value, out decimal percentValue))
                {
                    cell.Value = percentValue / 100;
                }
                else
                {
                    cell.Value = GetStringValue(value);
                }
                break;

            case ColumnType.Integer:
                if (TryGetIntegerValue(value, out int intValue))
                {
                    cell.Value = intValue;
                }
                else
                {
                    cell.Value = GetStringValue(value);
                }
                break;

            case ColumnType.Money:
                if (TryGetDecimalValue(value, out decimal currencyValue))
                {
                    cell.Value = currencyValue;
                }
                else
                {
                    cell.Value = GetStringValue(value);
                }
                break;

            case ColumnType.Decimal:
                if (TryGetDecimalValue(value, out decimal decimalValue))
                {
                    cell.Value = decimalValue;
                }
                else
                {
                    cell.Value = GetStringValue(value);
                }
                break;

            case ColumnType.Boolean:
                cell.Value = value.ValueKind switch
                {
                    JsonValueKind.True => true,
                    JsonValueKind.False => false,
                    _ => GetStringValue(value)
                };
                break;

            default:
                throw new NotImplementedException($"Still have to implement ColumnType.{column.Type}!");
        }
    }

    private static string GetStringValue(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString() ?? "",
            JsonValueKind.Number => element.GetDecimal().ToString(CultureInfo.InvariantCulture),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => "",
            JsonValueKind.Undefined => "",
            _ => element.ToString()
        };
    }

    private static bool TryGetDoubleValue(JsonElement element, out double value)
    {
        value = 0;

        return element.ValueKind switch
        {
            JsonValueKind.Number => element.TryGetDouble(out value),
            JsonValueKind.String => double.TryParse(element.GetString(), CultureInfo.InvariantCulture, out value),
            _ => false
        };
    }

    private static bool TryGetDecimalValue(JsonElement element, out decimal value)
    {
        value = 0;

        return element.ValueKind switch
        {
            JsonValueKind.Number => element.TryGetDecimal(out value),
            JsonValueKind.String => decimal.TryParse(element.GetString(), CultureInfo.InvariantCulture, out value),
            _ => false
        };
    }

    private static bool TryGetIntegerValue(JsonElement element, out int value)
    {
        value = 0;

        return element.ValueKind switch
        {
            JsonValueKind.Number => element.TryGetInt32(out value),
            JsonValueKind.String => int.TryParse(element.GetString(), CultureInfo.InvariantCulture, out value),
            _ => false
        };
    }

    private static bool TryGetDateValue(JsonElement element, out DateTime value)
    {
        value = default;
        if (element.ValueKind == JsonValueKind.String)
        {
            return DateTime.TryParse(element.GetString(), out value);
        }
        return false;
    }
}
