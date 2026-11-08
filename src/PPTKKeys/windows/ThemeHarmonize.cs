namespace PPTKKeys;

/// <summary>
/// Deterministic PPTK color mapping for Fix Slide Colors.
/// This is the PPTK rule, not a claim of auxi parity.
/// VBA ModulePPTKFixColors uses the same thresholds and formulas.
/// </summary>
public enum ThemeSlot
{
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
}

public readonly record struct Rgb(byte R, byte G, byte B);

public readonly record struct ThemePalette(
    Rgb Dark1,
    Rgb Light1,
    Rgb Dark2,
    Rgb Light2,
    Rgb Accent1,
    Rgb Accent2,
    Rgb Accent3,
    Rgb Accent4,
    Rgb Accent5,
    Rgb Accent6);

public readonly record struct HarmonizedColor(bool Change, ThemeSlot Slot, double Brightness);

public static class ThemeHarmonize
{
    public const double NeutralChroma = 0.12;
    public const double WhiteLuminance = 0.90;
    public const double BlackLuminance = 0.12;
    public const double LightNeutralLuminance = 0.72;
    public const double DarkNeutralLuminance = 0.28;
    public const double MaxBrightness = 0.55;
    public const double MinimumContrast = 3.0;

    public static double Channel(double value)
    {
        double c = value / 255.0;
        if (c <= 0.04045) return c / 12.92;
        return Math.Pow((c + 0.055) / 1.055, 2.4);
    }

    public static double Luminance(Rgb color) =>
        0.2126 * Channel(color.R) + 0.7152 * Channel(color.G) + 0.0722 * Channel(color.B);

    public static double Chroma(Rgb color)
    {
        double r = color.R / 255.0;
        double g = color.G / 255.0;
        double b = color.B / 255.0;
        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));
        return max - min;
    }

    public static double Hue(Rgb color)
    {
        double r = color.R / 255.0;
        double g = color.G / 255.0;
        double b = color.B / 255.0;
        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));
        double d = max - min;
        if (d < 1e-8) return 0;
        double h;
        if (max == r) h = (g - b) / d;
        else if (max == g) h = (b - r) / d + 2;
        else h = (r - g) / d + 4;
        if (h < 0) h += 6;
        return h * 60.0;
    }

    public static double CircularDistance(double a, double b)
    {
        double d = Math.Abs(a - b) % 360.0;
        if (d > 180) d = 360 - d;
        return d;
    }

    public static double ContrastRatio(double a, double b)
    {
        double lighter = Math.Max(a, b);
        double darker = Math.Min(a, b);
        return (lighter + 0.05) / (darker + 0.05);
    }

    public static HarmonizedColor Map(Rgb source, ThemePalette palette)
    {
        double luminance = Luminance(source);
        if (Chroma(source) <= NeutralChroma)
        {
            if (luminance >= WhiteLuminance)
                return new HarmonizedColor(true, ThemeSlot.Light1, 0);
            if (luminance <= BlackLuminance)
                return new HarmonizedColor(true, ThemeSlot.Dark1, 0);
            if (luminance >= LightNeutralLuminance)
                return Neutral(ThemeSlot.Light2, luminance, Luminance(palette.Light2));
            if (luminance <= DarkNeutralLuminance)
                return Neutral(ThemeSlot.Dark2, luminance, Luminance(palette.Dark2));
            return new HarmonizedColor(false, ThemeSlot.Dark1, 0);
        }

        var accents = new (ThemeSlot Slot, Rgb Color)[]
        {
            (ThemeSlot.Accent1, palette.Accent1),
            (ThemeSlot.Accent2, palette.Accent2),
            (ThemeSlot.Accent3, palette.Accent3),
            (ThemeSlot.Accent4, palette.Accent4),
            (ThemeSlot.Accent5, palette.Accent5),
            (ThemeSlot.Accent6, palette.Accent6),
        };

        ThemeSlot chosen = ThemeSlot.Accent1;
        Rgb chosenColor = palette.Accent1;
        double best = double.MaxValue;
        bool found = false;
        double sourceHue = Hue(source);
        foreach ((ThemeSlot slot, Rgb color) in accents)
        {
            if (Chroma(color) <= NeutralChroma) continue;
            double distance = CircularDistance(sourceHue, Hue(color));
            if (!found || distance < best - 1e-9)
            {
                found = true;
                best = distance;
                chosen = slot;
                chosenColor = color;
            }
        }

        double brightness = Luminance(source) - Luminance(chosenColor);
        if (brightness > MaxBrightness) brightness = MaxBrightness;
        if (brightness < -MaxBrightness) brightness = -MaxBrightness;
        return new HarmonizedColor(true, chosen, brightness);
    }

    public static bool NeedsContrastFlip(double textLuminance, double backgroundLuminance) =>
        ContrastRatio(textLuminance, backgroundLuminance) < MinimumContrast;

    public static ThemeSlot ContrastTextSlot(double backgroundLuminance) =>
        backgroundLuminance >= 0.5 ? ThemeSlot.Dark1 : ThemeSlot.Light1;

    static HarmonizedColor Neutral(ThemeSlot slot, double sourceLuminance, double slotLuminance)
    {
        double brightness = sourceLuminance - slotLuminance;
        if (brightness > MaxBrightness) brightness = MaxBrightness;
        if (brightness < -MaxBrightness) brightness = -MaxBrightness;
        return new HarmonizedColor(true, slot, brightness);
    }
}
