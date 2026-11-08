using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

public class ThemeHarmonizeTests
{
    static readonly ThemePalette Office = new(
        new Rgb(0, 0, 0),
        new Rgb(255, 255, 255),
        new Rgb(68, 84, 106),
        new Rgb(231, 230, 230),
        new Rgb(192, 0, 0),
        new Rgb(237, 125, 49),
        new Rgb(255, 192, 0),
        new Rgb(112, 173, 71),
        new Rgb(91, 155, 213),
        new Rgb(112, 48, 160));

    [Fact]
    public void White_and_black_stay_on_the_light_and_dark_theme_slots()
    {
        HarmonizedColor white = ThemeHarmonize.Map(new Rgb(255, 255, 255), Office);
        HarmonizedColor black = ThemeHarmonize.Map(new Rgb(0, 0, 0), Office);
        Assert.True(white.Change);
        Assert.Equal(ThemeSlot.Light1, white.Slot);
        Assert.Equal(0, white.Brightness);
        Assert.True(black.Change);
        Assert.Equal(ThemeSlot.Dark1, black.Slot);
        Assert.Equal(0, black.Brightness);
    }

    [Fact]
    public void Mid_gray_is_preserved_and_dark_gray_uses_dark2()
    {
        HarmonizedColor mid = ThemeHarmonize.Map(new Rgb(188, 188, 188), Office);
        HarmonizedColor dark = ThemeHarmonize.Map(new Rgb(128, 128, 128), Office);
        Assert.False(mid.Change);
        Assert.True(dark.Change);
        Assert.Equal(ThemeSlot.Dark2, dark.Slot);
    }

    [Fact]
    public void Distinct_hues_use_distinct_accents()
    {
        HarmonizedColor red = ThemeHarmonize.Map(new Rgb(192, 0, 0), Office);
        HarmonizedColor blue = ThemeHarmonize.Map(new Rgb(91, 155, 213), Office);
        Assert.Equal(ThemeSlot.Accent1, red.Slot);
        Assert.Equal(ThemeSlot.Accent5, blue.Slot);
        Assert.NotEqual(red.Slot, blue.Slot);
        Assert.InRange(red.Brightness, -ThemeHarmonize.MaxBrightness, ThemeHarmonize.MaxBrightness);
    }

    [Fact]
    public void Light_and_dark_reds_share_an_accent_with_opposite_brightness()
    {
        HarmonizedColor light = ThemeHarmonize.Map(new Rgb(255, 140, 140), Office);
        HarmonizedColor dark = ThemeHarmonize.Map(new Rgb(90, 0, 0), Office);
        Assert.Equal(ThemeSlot.Accent1, light.Slot);
        Assert.Equal(ThemeSlot.Accent1, dark.Slot);
        Assert.True(light.Brightness > 0);
        Assert.True(dark.Brightness < 0);
    }

    [Fact]
    public void Equal_hue_distance_picks_the_earlier_accent()
    {
        var palette = Office with
        {
            Accent1 = new Rgb(255, 128, 0),
            Accent2 = new Rgb(255, 0, 128),
        };
        HarmonizedColor red = ThemeHarmonize.Map(new Rgb(220, 0, 0), palette);
        double hue = ThemeHarmonize.Hue(new Rgb(220, 0, 0));
        double towardFirst = ThemeHarmonize.CircularDistance(hue, ThemeHarmonize.Hue(palette.Accent1));
        double towardSecond = ThemeHarmonize.CircularDistance(hue, ThemeHarmonize.Hue(palette.Accent2));
        Assert.Equal(towardFirst, towardSecond, 1);
        Assert.Equal(ThemeSlot.Accent1, red.Slot);
    }

    [Fact]
    public void Low_contrast_text_flips_toward_the_opposite_neutral()
    {
        Assert.True(ThemeHarmonize.NeedsContrastFlip(0.85, 0.90));
        Assert.Equal(ThemeSlot.Dark1, ThemeHarmonize.ContrastTextSlot(0.90));
        Assert.Equal(ThemeSlot.Light1, ThemeHarmonize.ContrastTextSlot(0.10));
        Assert.False(ThemeHarmonize.NeedsContrastFlip(0.0, 1.0));
    }
}
