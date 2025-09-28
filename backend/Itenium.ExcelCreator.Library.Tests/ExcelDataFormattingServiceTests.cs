using System.Globalization;
using Itenium.ExcelCreator.Library.Models;

namespace Itenium.ExcelCreator.Library.Tests;

public class ExcelDataFormattingServiceTests
{
    private ExcelService _service;

    [SetUp]
    public void Setup()
    {
        _service = new ExcelService();
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
            Data = Helpers.CreateTestData([
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
}
