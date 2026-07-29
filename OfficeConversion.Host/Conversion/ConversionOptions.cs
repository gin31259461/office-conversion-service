using System.ComponentModel.DataAnnotations;

namespace OfficeConversion.Conversion;

public sealed class ConversionOptions
{
    public const string SectionName = "Conversion";

    [Range(1, 1000)]
    public int QueueCapacity { get; init; } = 20;

    [Range(10, 3600)]
    public int TimeoutSeconds { get; init; } = 120;
}
