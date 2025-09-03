using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace Itenium.ExcelCreator.WebApi.Helpers;

public static class Extensions
{
    private const string ExcelContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public static FileStreamResult Deliver(this IXLWorkbook workbook, string fileName)
    {
        var memoryStream = new MemoryStream();
        workbook.SaveAs(memoryStream);
        memoryStream.Seek(0, SeekOrigin.Begin);
        return new FileStreamResult(memoryStream, ExcelContentType) { FileDownloadName = fileName };
    }
}
