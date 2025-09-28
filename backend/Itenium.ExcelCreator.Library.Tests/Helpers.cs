using System.Text.Json;

namespace Itenium.ExcelCreator.Library.Tests;

public static class Helpers
{
    public static JsonElement[][] CreateTestData(object?[][] rawData)
    {
        var jsonDoc = JsonSerializer.SerializeToDocument(rawData);
        return rawData.Select((row, i) =>
            row.Select((cell, j) => jsonDoc.RootElement[i][j]).ToArray()
        ).ToArray();
    }
}
