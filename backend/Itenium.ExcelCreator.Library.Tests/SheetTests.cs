using Itenium.ExcelCreator.Library.Models;
using ClosedXML.Excel;

namespace Itenium.ExcelCreator.Library.Tests;

public class SheetTests
{
    private ExcelService _service;

    [SetUp]
    public void Setup()
    {
        _service = new ExcelService();
    }

    [Test]
    public void CreateWorkbookWithCorrectSheetName()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                SheetName = "TestSheet",
            },
        };

        var workbook = _service.CreateExcel(data);

        Assert.That(workbook.Worksheets, Has.Count.EqualTo(1));
        Assert.That(workbook.Worksheet(1).Name, Is.EqualTo("TestSheet"));
    }

    [Test]
    public void Headers_AreBoldAndColored()
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
            Data = Helpers.CreateTestData([
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
    public void MoreDataColumnsThanConfigColumns_HandlesExtraColumns_AsStrings()
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
            Data = Helpers.CreateTestData([
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
    public void NoData_CreatesWorkbookWithHeadersOnly()
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
    public void SetsAutoFilter()
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
            Data = Helpers.CreateTestData([
                ["John", 30]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.AutoFilter.IsEnabled, Is.True);
        Assert.That(worksheet.AutoFilter.Range.RangeAddress.ToString(), Is.EqualTo("A1:B2"));
    }

    [Test]
    public void AdjustsColumnsToContents()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Description", Type = ColumnType.String }
                ]
            },
            Data = Helpers.CreateTestData([
                ["This is a very long text that should cause the column to be adjusted"]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        double width = worksheet.Column(1).Width;
        Assert.That(width, Is.GreaterThan(15.0));
    }

    [Test]
    public void FreezesColumns_IsZero_WhenNotSet()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Name", Type = ColumnType.String },
                    new() { Header = "Age", Type = ColumnType.Integer },
                    new() { Header = "Salary", Type = ColumnType.Money }
                ],
            },
            Data = Helpers.CreateTestData([
                ["John", 30, 50000]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.SheetView.SplitColumn, Is.EqualTo(0));
    }

    [Test]
    public void FreezesColumns()
    {
        var data = new FullExcelData
        {
            Config = new ExcelConfiguration
            {
                Columns = [
                    new() { Header = "Name", Type = ColumnType.String },
                    new() { Header = "Age", Type = ColumnType.Integer },
                    new() { Header = "Salary", Type = ColumnType.Money }
                ],
                FreezeColumns = 2
            },
            Data = Helpers.CreateTestData([
                ["John", 30, 50000]
            ])
        };

        var workbook = _service.CreateExcel(data);
        var worksheet = workbook.Worksheet(1);

        Assert.That(worksheet.SheetView.SplitColumn, Is.EqualTo(2));
    }
}
