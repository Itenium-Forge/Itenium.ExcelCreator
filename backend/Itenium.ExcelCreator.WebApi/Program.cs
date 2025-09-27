using Itenium.ExcelCreator.Library;
using Itenium.ExcelCreator.Library.Models;
using Itenium.ExcelCreator.WebApi;
using Itenium.Forge.Controllers;
using Itenium.Forge.Logging;
using Itenium.Forge.Settings;
using Itenium.Forge.Swagger;
using Serilog;

Log.Logger = LoggingExtensions.CreateBootstrapLogger();

try
{
    var builder = WebApplication.CreateBuilder(args);
    builder.AddForgeSettings<ExcelCreatorSettings>();
    builder.AddForgeLogging();

    builder.AddForgeControllers();
    builder.AddForgeSwagger(typeof(ColumnType));

    builder.Services.AddScoped<ExcelService>();

    WebApplication app = builder.Build();
    app.UseForgeLogging();

    app.UseForgeControllers();
    app.UseForgeSwagger();

    app.Run();
}
catch (Exception ex)
{
    Log.Fatal(ex, "Application terminated unexpectedly");
}
finally
{
    await Log.CloseAndFlushAsync();
}
