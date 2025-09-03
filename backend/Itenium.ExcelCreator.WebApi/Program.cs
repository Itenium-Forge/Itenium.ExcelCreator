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
    var settings = builder.AddForgeSettings<ExcelCreatorSettings>();
    builder.AddForgeLogging();

    builder.AddForgeControllers();
    builder.AddForgeSwagger();

    WebApplication app = builder.Build();
    app.UseForgeLogging();

    // app.UseAuthorization();

    app.UseForgeControllers();
    //if (app.Environment.IsDevelopment())
    app.UseForgeSwagger();

    var logger = app.Services.GetRequiredService<ILogger<Program>>();
    logger.LogInformation("DOTNET_ENVIRONMENT: " + Environment.GetEnvironmentVariable("DOTNET_ENVIRONMENT"));
    logger.LogInformation("ASPNETCORE_ENVIRONMENT: " + Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT"));

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
