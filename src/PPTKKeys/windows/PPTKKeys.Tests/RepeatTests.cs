using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

public class RepeatTests
{
    static SequencePress Plus(bool shift = false, bool ctrl = false, bool alt = false) =>
        new(SequenceSymbol.Plus, '\0', ctrl, alt, false, shift);

    static SequencePress Minus(bool shift = false) =>
        new(SequenceSymbol.Minus, '\0', false, false, false, shift);

    [Theory]
    [InlineData("H +", true)]
    [InlineData("H -", true)]
    [InlineData("V +", true)]
    [InlineData("V -", true)]
    [InlineData("A A", false)]
    [InlineData("X T", false)]
    [InlineData("A L", false)]
    [InlineData("D H", false)]
    [InlineData(". T", false)]
    [InlineData("T", false)]
    [InlineData("R", false)]
    [InlineData("N", false)]
    [InlineData("I", false)]
    [InlineData("ALT .", false)]
    [InlineData("CTRL K", false)]
    [InlineData("CTRL 2", false)]
    [InlineData("ALT 3", false)]
    [InlineData(null, false)]
    public void Only_spacing_repeats(string? token, bool repeatable) =>
        Assert.Equal(repeatable, ShortcutRepeat.IsRepeatable(token));

    [Fact]
    public void Held_plus_repeats_after_the_previous_run_and_interval()
    {
        Assert.Equal(RepeatStep.Repeat, ShortcutRepeat.Decide("H +", Plus(shift: true), true, false, 1100, 1000));
        Assert.Equal(RepeatStep.Repeat, ShortcutRepeat.Decide("V +", Plus(), true, false, 1050, 1000));
        Assert.Equal(RepeatStep.Repeat, ShortcutRepeat.Decide("H -", Minus(), true, false, 2000, 1000));
        Assert.Equal(RepeatStep.Repeat, ShortcutRepeat.Decide("V -", Minus(), true, false, 2000, 1000));
    }

    [Fact]
    public void Waits_while_a_run_is_in_flight_or_inside_the_interval()
    {
        Assert.Equal(RepeatStep.Wait, ShortcutRepeat.Decide("H +", Plus(), true, inFlight: true, 5000, 1000));
        Assert.Equal(RepeatStep.Wait, ShortcutRepeat.Decide("H +", Plus(), true, false, 1049, 1000));
    }

    [Fact]
    public void Tick_wrap_does_not_block_repeat()
    {
        Assert.Equal(RepeatStep.Repeat, ShortcutRepeat.Decide("H +", Plus(), true, false, 30, uint.MaxValue - 40));
    }

    [Fact]
    public void Stops_when_the_canvas_is_gone()
    {
        Assert.Equal(RepeatStep.Stop, ShortcutRepeat.Decide("H +", Plus(), allowCanvas: false, false, 2000, 1000));
    }

    [Fact]
    public void Stops_when_a_command_modifier_joins_the_hold()
    {
        Assert.Equal(RepeatStep.Stop, ShortcutRepeat.Decide("H +", Plus(ctrl: true), true, false, 2000, 1000));
        Assert.Equal(RepeatStep.Stop, ShortcutRepeat.Decide("H +", Plus(alt: true), true, false, 2000, 1000));
    }

    [Fact]
    public void Stops_when_the_held_key_changes_meaning()
    {
        // Shift released on US =/+ : the key now types =, which is not plus.
        var equalsKey = new SequencePress(SequenceSymbol.Other, '\0', false, false, false, false);
        Assert.Equal(RepeatStep.Stop, ShortcutRepeat.Decide("H +", equalsKey, true, false, 2000, 1000));
        // Shift pressed on the minus key: underscore is not minus.
        Assert.Equal(RepeatStep.Stop, ShortcutRepeat.Decide("H -", Minus(shift: true), true, false, 2000, 1000));
        Assert.Equal(RepeatStep.Stop, ShortcutRepeat.Decide("H +", Minus(), true, false, 2000, 1000));
    }

    [Fact]
    public void Nothing_held_never_repeats()
    {
        Assert.Equal(RepeatStep.Stop, ShortcutRepeat.Decide(null, Plus(), true, false, 2000, 1000));
        Assert.Equal(RepeatStep.Stop, ShortcutRepeat.Decide("A A", Plus(), true, false, 2000, 1000));
    }

    [Fact]
    public void Layout_plus_and_minus_map_to_the_repeat_symbols()
    {
        // US: Shift+= types +, numpad + and - always count, OEM minus without Shift is minus.
        Assert.Equal(SequenceSymbol.Plus, ShortcutSequence.PressFromLayout(0xBB, false, false, false, true, '+').Symbol);
        Assert.Equal(SequenceSymbol.Plus, ShortcutSequence.PressFromLayout(0x6B, false, false, false, false, '\0').Symbol);
        Assert.Equal(SequenceSymbol.Minus, ShortcutSequence.PressFromLayout(0x6D, false, false, false, false, '\0').Symbol);
        Assert.Equal(SequenceSymbol.Minus, ShortcutSequence.PressFromLayout(0xBD, false, false, false, false, '-').Symbol);
        Assert.Equal(SequenceSymbol.Other, ShortcutSequence.PressFromLayout(0xBB, false, false, false, false, '\0').Symbol);
    }
}
