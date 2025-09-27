using System.Globalization;
using Itenium.ExcelCreator.Library.Models;
using ClosedXML.Excel;
using System.Text.Json;

namespace Itenium.ExcelCreator.Library.Tests;

public class ExcelServiceTests
{
    private ExcelService _service;

    [SetUp]
    public void Setup()
    {
        _service = new ExcelService();
    }

    [Test]
    public void CreateExcel_WithBasicConfiguration_CreatesWorkbookWithCorrectSheetName()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                SheetName = "TestSheet",
            },
        };

        var workbook = _service.CreateExcel(data);

        Assert.That(workbook.Worksheets.Count, Is.EqualTo(1));
        Assert.That(workbook.Worksheet(1).Name, Is.EqualTo("TestSheet"));
    }

    [Test]
    public void CreateExcel_WithHeaders_CreatesCorrectHeaders()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Name", Type = ColumnType.String },
                    new() { Header = "Age", Type = ColumnType.Integer },
                    new() { Header = "Salary", Type = ColumnType.Money }
                ]
            },
            Data = CreateTestData([
                ["John Doe", 30, 50000.50],
                ["Jane Smith", 25, 60000.75]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(1, 1).Value.ToString(), Is.EqualTo("Name"));
        Assert.That(worksheet.Cell(1, 2).Value.ToString(), Is.EqualTo("Age"));
        Assert.That(worksheet.Cell(1, 3).Value.ToString(), Is.EqualTo("Salary"));
        
        Assert.That(worksheet.Cell(1, 1).Style.Font.Bold, Is.True);
        Assert.That(worksheet.Cell(1, 1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.LightGray));

        Assert.That(worksheet.Cell(2, 1).Style.Font.Bold, Is.False);
        Assert.That(worksheet.Cell(2, 1).Style.Fill.BackgroundColor, Is.Not.EqualTo(XLColor.LightGray));
    }

    [Test]
    public void CreateExcel_WithStringData_FormatsCorrectly()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Name", Type = ColumnType.String }
                ]
            },
            Data = CreateTestData([
                ["John Doe"],
                ["Jane Smith"]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value.ToString(), Is.EqualTo("John Doe"));
        Assert.That(worksheet.Cell(3, 1).Value.ToString(), Is.EqualTo("Jane Smith"));
    }

    [Test]
    public void CreateExcel_WithIntegerData_FormatsWithThousandsSeparator()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Count", Type = ColumnType.Integer }
                ]
            },
            Data = CreateTestData([
                [1000],
                [2500000]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value, Is.EqualTo(1000));
        Assert.That(worksheet.Cell(3, 1).Value, Is.EqualTo(2500000));
        Assert.That(worksheet.Cell(2, 1).Style.NumberFormat.Format, Is.EqualTo("#,##0"));
        Assert.That(worksheet.Cell(3, 1).Style.NumberFormat.Format, Is.EqualTo("#,##0"));
    }

    [Test]
    public void CreateExcel_WithMoneyData_FormatsCurrency()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Salary", Type = ColumnType.Money }
                ]
            },
            Data = CreateTestData([
                [1250.75],
                [95.2301]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value, Is.EqualTo(1250.75));
        Assert.That(worksheet.Cell(3, 1).Value, Is.EqualTo(95.2301));
        Assert.That(worksheet.Cell(2, 1).Style.NumberFormat.Format, Is.EqualTo("€ #,##0.00"));
        Assert.That(worksheet.Cell(3, 1).Style.NumberFormat.Format, Is.EqualTo("€ #,##0.00"));
    }

    [Test]
    public void CreateExcel_WithMoneyData_WithTooHighPrecision_Truncates()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Salary", Type = ColumnType.Money }
                ]
            },
            Data = CreateTestData([
                [1234.123456789012345m],
                ["1234.123456789012345"]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value.ToString(CultureInfo.InvariantCulture), Is.EqualTo("1234.1234567890122"));
        Assert.That(worksheet.Cell(2, 1).GetValue<decimal>(), Is.EqualTo(1234.12345678901m));

        Assert.That(worksheet.Cell(3, 1).Value.ToString(CultureInfo.InvariantCulture), Is.EqualTo("1234.1234567890122"));
        Assert.That(worksheet.Cell(3, 1).GetValue<decimal>(), Is.EqualTo(1234.12345678901m));
    }

    [Test]
    public void CreateExcel_WithDecimalData_FormatsDecimalPlaces()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Value", Type = ColumnType.Decimal }
                ]
            },
            Data = CreateTestData([
                [123.456],
                [78.90]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value, Is.EqualTo(123.456));
        Assert.That(worksheet.Cell(3, 1).Value, Is.EqualTo(78.90));
        Assert.That(worksheet.Cell(2, 1).Style.NumberFormat.Format, Is.EqualTo("#,##0.00"));
        Assert.That(worksheet.Cell(3, 1).Style.NumberFormat.Format, Is.EqualTo("#,##0.00"));
    }

    [Test]
    public void CreateExcel_WithPercentageData_FormatsAsPercentage()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Score", Type = ColumnType.Percentage }
                ]
            },
            Data = CreateTestData([
                [85.5m],
                [92.3m]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value, Is.EqualTo(0.855));
        Assert.That(worksheet.Cell(3, 1).Value, Is.EqualTo(0.923));
        Assert.That(worksheet.Cell(2, 1).Style.NumberFormat.Format, Is.EqualTo("0.00%"));
        Assert.That(worksheet.Cell(3, 1).Style.NumberFormat.Format, Is.EqualTo("0.00%"));
    }

    [Test]
    public void CreateExcel_WithBooleanData_DisplaysBooleanValues()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Active", Type = ColumnType.Boolean }
                ]
            },
            Data = CreateTestData([
                [true],
                [false]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value, Is.EqualTo(true));
        Assert.That(worksheet.Cell(3, 1).Value, Is.EqualTo(false));
    }

    [Test]
    public void CreateExcel_WithDateData_FormatsAsDate()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Date", Type = ColumnType.Date }
                ]
            },
            Data = CreateTestData([
                ["2024-12-15T10:30:00"]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        var expectedDate = new DateTime(2024, 12, 15, 10, 30, 0);
        Assert.That(worksheet.Cell(2, 1).Value, Is.EqualTo(expectedDate));
        Assert.That(worksheet.Cell(2, 1).Style.DateFormat.Format, Is.EqualTo("mm/dd/yyyy"));
    }

    [Test]
    public void CreateExcel_WithNullValues_HandlesNullsCorrectly()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Data", Type = ColumnType.String }
                ]
            },
            Data = CreateTestData([
                [null],
                ["ValidValue"]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value.ToString(), Is.EqualTo(""));
        Assert.That(worksheet.Cell(3, 1).Value.ToString(), Is.EqualTo("ValidValue"));
    }

    [Test]
    public void CreateExcel_WithInvalidDateString_FallsBackToStringValue()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Date", Type = ColumnType.Date }
                ]
            },
            Data = CreateTestData([
                ["not-a-date"]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value.ToString(), Is.EqualTo("not-a-date"));
    }

    [Test]
    public void CreateExcel_WithInvalidNumericString_FallsBackToStringValue()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Amount", Type = ColumnType.Money }
                ]
            },
            Data = CreateTestData([
                ["not-a-number"]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value.ToString(), Is.EqualTo("not-a-number"));
    }

    [Test]
    public void CreateExcel_WithMixedDataTypes_HandlesAllTypesCorrectly()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Name", Type = ColumnType.String },
                    new() { Header = "Age", Type = ColumnType.Integer },
                    new() { Header = "Salary", Type = ColumnType.Money },
                    new() { Header = "Active", Type = ColumnType.Boolean },
                    new() { Header = "StartDate", Type = ColumnType.Date },
                    new() { Header = "Score", Type = ColumnType.Percentage }
                ]
            },
            Data = CreateTestData([
                [
                    "John Doe",
                    30,
                    50000.50,
                    true,
                    "2024-01-15",
                    95.5
                ]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value.ToString(), Is.EqualTo("John Doe"));
        Assert.That(worksheet.Cell(2, 2).Value, Is.EqualTo(30));
        Assert.That(worksheet.Cell(2, 3).Value, Is.EqualTo(50000.50));
        Assert.That(worksheet.Cell(2, 4).Value, Is.EqualTo(true));
        Assert.That(worksheet.Cell(2, 5).Value, Is.EqualTo(DateTime.Parse("2024-01-15")));
        Assert.That(worksheet.Cell(2, 6).Value, Is.EqualTo(0.955));

        Assert.That(worksheet.Cell(2, 2).Style.NumberFormat.Format, Is.EqualTo("#,##0"));
        Assert.That(worksheet.Cell(2, 3).Style.NumberFormat.Format, Is.EqualTo("€ #,##0.00"));
        Assert.That(worksheet.Cell(2, 5).Style.DateFormat.Format, Is.EqualTo("mm/dd/yyyy"));
        Assert.That(worksheet.Cell(2, 6).Style.NumberFormat.Format, Is.EqualTo("0.00%"));
    }

    [Test]
    public void CreateExcel_WithMoreDataColumnsThanConfigColumns_HandlesExtraColumns()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Name", Type = ColumnType.String },
                    new() { Header = "Age", Type = ColumnType.Integer }
                ]
            },
            Data = CreateTestData([
                ["John", 30, "Extra Data"]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value.ToString(), Is.EqualTo("John"));
        Assert.That(worksheet.Cell(2, 2).Value, Is.EqualTo(30));
        Assert.That(worksheet.Cell(2, 3).Value.ToString(), Is.EqualTo("Extra Data"));
    }

    [Test]
    public void CreateExcel_WithEmptyData_CreatesWorkbookWithHeadersOnly()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Name", Type = ColumnType.String },
                    new() { Header = "Age", Type = ColumnType.Integer }
                ]
            },
            Data = []
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(1, 1).Value.ToString(), Is.EqualTo("Name"));
        Assert.That(worksheet.Cell(1, 2).Value.ToString(), Is.EqualTo("Age"));
        Assert.That(worksheet.Cell(2, 1).IsEmpty(), Is.True);
        Assert.That(worksheet.Cell(2, 2).IsEmpty(), Is.True);
    }

    [Test]
    public void CreateExcel_SetsAutoFilter()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Name", Type = ColumnType.String },
                    new() { Header = "Age", Type = ColumnType.Integer }
                ]
            },
            Data = CreateTestData([
                ["John", 30]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.AutoFilter.IsEnabled, Is.True);
        Assert.That(worksheet.AutoFilter.Range.RangeAddress.ToString(), Is.EqualTo("A1:B2"));
    }

    [Test]
    public void CreateExcel_AdjustsColumnsToContents()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Description", Type = ColumnType.String }
                ]
            },
            Data = CreateTestData([
                ["This is a very long text that should cause the column to be adjusted"]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        double width = worksheet.Column(1).Width;
        Assert.That(width, Is.GreaterThan(15.0));
    }

    private static JsonElement[][] CreateTestData(object?[][] rawData)
    {
        var jsonDoc = JsonSerializer.SerializeToDocument(rawData);
        return rawData.Select((row, i) =>
            row.Select((cell, j) => jsonDoc.RootElement[i][j]).ToArray()
        ).ToArray();
    }
}
