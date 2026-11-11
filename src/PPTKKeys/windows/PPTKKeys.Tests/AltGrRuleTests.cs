using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

public class AltGrRuleTests
{
    [Theory]
    [InlineData(0xA2, 0x21Du, true)]   // synthetic Left Ctrl sent with Right Alt on AltGr layouts
    [InlineData(0xA2, 0x01Du, false)]  // real Left Ctrl
    [InlineData(0xA3, 0x11Du, false)]  // Right Ctrl
    [InlineData(0xA5, 0x038u, false)]  // Right Alt itself
    public void Only_the_scan_21D_left_ctrl_is_synthetic(int vk, uint scan, bool synthetic) =>
        Assert.Equal(synthetic, AltGrRule.IsSyntheticCtrl(vk, scan));

    [Fact]
    public void Right_alt_on_an_altgr_layout_acts_as_alt_for_a_key_without_an_altgr_character() =>
        Assert.True(AltGrRule.ActsAsAlt(syntheticCtrlHeld: true, rightAltHeld: true, realCtrlHeld: false, keyTypesAltGrCharacter: false));

    [Fact]
    public void Altgr_typing_is_kept_when_the_layout_types_a_character() =>
        Assert.False(AltGrRule.ActsAsAlt(true, true, false, keyTypesAltGrCharacter: true));

    [Fact]
    public void Real_ctrl_plus_alt_stays_ctrl_alt() =>
        Assert.False(AltGrRule.ActsAsAlt(true, true, realCtrlHeld: true, false));

    [Fact]
    public void Without_the_synthetic_ctrl_nothing_changes()
    {
        // Left Ctrl + Left Alt pressed by hand.
        Assert.False(AltGrRule.ActsAsAlt(syntheticCtrlHeld: false, rightAltHeld: false, realCtrlHeld: true, false));
        // Synthetic Ctrl seen but Right Alt already released.
        Assert.False(AltGrRule.ActsAsAlt(true, rightAltHeld: false, false, false));
    }

    [Fact]
    public void Plain_alt_period_chord_is_what_the_rule_restores()
    {
        // After the rule clears Ctrl, the chord is the same as Left Alt + period.
        Assert.Equal("ALT .", ShortcutChords.Token(KeyMark.Period, 0xBE, ctrl: false, shift: false, alt: true, win: false));
        Assert.Equal("ALT 2", ShortcutChords.Token(KeyMark.None, 0x32, false, false, true, false));
        Assert.Equal("ALT 3", ShortcutChords.Token(KeyMark.None, 0x33, false, false, true, false));
        // Unchanged Ctrl+Alt is still left to the layout.
        Assert.Null(ShortcutChords.Token(KeyMark.Period, 0xBE, ctrl: true, shift: false, alt: true, win: false));
    }
}
