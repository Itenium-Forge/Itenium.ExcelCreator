using Itenium.ExcelCreator.Library.Models;

namespace Itenium.ExcelCreator.Library.Tests;

public class SpecialExcelServiceTests
{
    private ExcelService _service;

    [SetUp]
    public void Setup()
    {
        _service = new ExcelService();
    }

    [Test]
    public void Formula_WithTemplate()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Value1", Type = ColumnType.Integer },
                    new() { Header = "Value2", Type = ColumnType.Integer },
                    new() { Header = "Sum", Type = ColumnType.Integer, Formula = "=A{row}+B{row}"},
                ]
            },
            Data = Helpers.CreateTestData([
                [100, 200, null],
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value, Is.EqualTo(100));
        Assert.That(worksheet.Cell(2, 2).Value, Is.EqualTo(200));

        Assert.That(worksheet.Cell(2, 3).FormulaA1, Is.EqualTo("A2+B2"));
        Assert.That(worksheet.Cell(2, 3).Value, Is.EqualTo(300));
        Assert.That(worksheet.Cell(2, 3).Style.NumberFormat.Format, Is.EqualTo("#,##0"));
    }

    [Test]
    public void FormulaNotStartingWithEqualSign_AlsoWorks()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Value1", Type = ColumnType.Integer },
                    new() { Header = "Value2", Type = ColumnType.Integer },
                    new() { Header = "Sum", Type = ColumnType.Integer, Formula = "A{row}+B{row}"},
                ]
            },
            Data = Helpers.CreateTestData([
                [100, 200, null],
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 3).FormulaA1, Is.EqualTo("A2+B2"));
        Assert.That(worksheet.Cell(2, 3).Value, Is.EqualTo(300));
    }

    [Test]
    public void Formula_WithoutDataValue()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Sum", Type = ColumnType.Integer, Formula = "=10+10"},
                ]
            },
            Data = Helpers.CreateTestData([
                [], // No value for Sum column
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value, Is.EqualTo(20));
    }

    [Test]
    public void Formula_WithDataValue()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Sum", Type = ColumnType.Integer, Formula = "=10+10"},
                ]
            },
            Data = Helpers.CreateTestData([
                [200],
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.Cell(2, 1).Value, Is.EqualTo(20));
    }
}
