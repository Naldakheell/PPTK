using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

public class LayoutRuleTests
{
    [Theory]
    [InlineData(0x41, true)]   // A
    [InlineData(0x44, true)]   // D: Arabic Shift+D types ]
    [InlineData(0x46, true)]   // F: Arabic Shift+F types [
    [InlineData(0x4C, true)]   // L: Arabic Shift+L types /
    [InlineData(0x5A, true)]   // Z
    [InlineData(0xBE, false)]  // period key
    [InlineData(0xBF, false)]  // slash key
    [InlineData(0xDB, false)]
    [InlineData(0xDD, false)]
    [InlineData(0x6E, false)]  // numpad decimal
    [InlineData(0x32, false)]  // 2
    public void Letter_keys_never_carry_punctuation_shortcuts(int virtualKey, bool letter) =>
        Assert.Equal(letter, LayoutRule.IsLetterKey(virtualKey));

    [Theory]
    [InlineData('ز', true)]   // Arabic 101 period key
    [InlineData('ظ', true)]   // Arabic 101 slash key
    [InlineData('ج', true)]   // Arabic 101 [ key
    [InlineData('د', true)]   // Arabic 101 ] key
    [InlineData('.', false)]  // US: the key already is period
    [InlineData(';', false)]  // French: punctuation stays with the layout
    [InlineData(':', false)]
    [InlineData('\0', false)] // dead key or nothing
    public void Us_punctuation_key_is_kept_only_where_the_layout_types_a_letter(char produced, bool keep) =>
        Assert.Equal(keep, LayoutRule.KeepsUsPunctuation(produced));
}
