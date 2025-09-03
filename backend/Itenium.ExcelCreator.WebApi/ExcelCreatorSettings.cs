using Itenium.Forge.Core;
using Itenium.Forge.Settings;

namespace Itenium.ExcelCreator.WebApi;

public class ExcelCreatorSettings : IForgeSettings
{
    public ForgeSettings Forge { get; } = new();
}
