using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

public class SequenceTests
{
    static SequencePress Letter(char letter, bool shift = false, bool ctrl = false) =>
        new(SequenceSymbol.Letter, letter, ctrl, false, false, shift);

    static SequenceResult Tap(SequenceState state, SequencePress press, uint now, bool allowArm) =>
        ShortcutSequence.Next(state, press, now, allowArm);

    [Theory]
    [InlineData('A', 'L', "A L")]
    [InlineData('A', 'R', "A R")]
    [InlineData('A', 'T', "A T")]
    [InlineData('A', 'B', "A B")]
    [InlineData('A', 'M', "A M")]
    [InlineData('A', 'C', "A C")]
    [InlineData('A', 'A', "A A")]
    [InlineData('A', 'P', "A P")]
    [InlineData('A', 'X', "A X")]
    [InlineData('A', 'Z', "A Z")]
    [InlineData('S', 'W', "S W")]
    [InlineData('S', 'H', "S H")]
    [InlineData('S', 'S', "S S")]
    [InlineData('D', 'H', "D H")]
    [InlineData('D', 'V', "D V")]
    [InlineData('X', 'T', "X T")]
    public void Letter_sequences_dispatch_on_the_canvas(char prefix, char second, string token)
    {
        SequenceResult armed = Tap(SequenceState.Idle, Letter(prefix), 1000, allowArm: true);
        Assert.True(armed.Swallow);
        Assert.Null(armed.Token);
        Assert.Equal(prefix, armed.State.Prefix);

        SequenceResult done = Tap(armed.State, Letter(second), 1500, allowArm: true);
        Assert.True(done.Swallow);
        Assert.Equal(token, done.Token);
        Assert.False(done.State.IsArmed);
    }

    [Fact]
    public void Spacing_accepts_shift_plus_and_unshifted_minus()
    {
        SequenceState horizontal = Tap(SequenceState.Idle, Letter('H'), 10, true).State;
        SequencePress plus = ShortcutSequence.PressFromLayout(0xBB, false, false, false, shift: true, produced: '+');
        SequenceResult wider = Tap(horizontal, plus, 20, true);
        Assert.Equal("H +", wider.Token);

        SequenceState again = Tap(SequenceState.Idle, Letter('H'), 10, true).State;
        SequencePress minus = ShortcutSequence.PressFromLayout(0xBD, false, false, false, shift: false, produced: '-');
        Assert.Equal("H -", Tap(again, minus, 20, true).Token);

        SequenceState vertical = Tap(SequenceState.Idle, Letter('V'), 10, true).State;
        SequencePress numpadPlus = ShortcutSequence.PressFromLayout(0x6B, false, false, false, shift: false, produced: '\0');
        Assert.Equal("V +", Tap(vertical, numpadPlus, 20, true).Token);

        SequenceState verticalMinus = Tap(SequenceState.Idle, Letter('V'), 10, true).State;
        SequencePress numpadMinus = ShortcutSequence.PressFromLayout(0x6D, false, false, false, shift: false, produced: '\0');
        Assert.Equal("V -", Tap(verticalMinus, numpadMinus, 20, true).Token);
    }

    [Fact]
    public void Layout_fallbacks_cover_shift_plus_and_oem_minus()
    {
        Assert.Equal(SequenceSymbol.Plus, ShortcutSequence.PressFromLayout(0xBB, false, false, false, true, '\0').Symbol);
        Assert.Equal(SequenceSymbol.Minus, ShortcutSequence.PressFromLayout(0xBD, false, false, false, false, '\0').Symbol);
        Assert.Equal(SequenceSymbol.Plus, ShortcutSequence.PressFromLayout(0xBB, false, false, false, false, '+').Symbol);
        Assert.Equal(SequenceSymbol.Other, ShortcutSequence.PressFromLayout(0xBD, false, false, false, true, '_').Symbol);
        Assert.Equal(SequenceSymbol.Other, ShortcutSequence.PressFromLayout(0xBB, false, false, false, false, '=').Symbol);
    }

    [Fact]
    public void Shifted_minus_and_shifted_letters_do_not_complete()
    {
        SequenceState horizontal = Tap(SequenceState.Idle, Letter('H'), 10, true).State;
        SequencePress underscore = ShortcutSequence.PressFromLayout(0xBD, false, false, false, shift: true, produced: '_');
        SequenceResult cancelled = Tap(horizontal, underscore, 20, true);
        Assert.Null(cancelled.Token);
        Assert.False(cancelled.Swallow);

        SequenceState align = Tap(SequenceState.Idle, Letter('A'), 10, true).State;
        SequenceResult shifted = Tap(align, Letter('L', shift: true), 20, true);
        Assert.Null(shifted.Token);
        Assert.False(shifted.Swallow);
    }

    [Fact]
    public void Prefixes_do_not_arm_off_the_canvas()
    {
        foreach (char prefix in new[] { 'A', 'S', 'D', 'H', 'V', 'X' })
        {
            SequenceResult typed = Tap(SequenceState.Idle, Letter(prefix), 10, allowArm: false);
            Assert.False(typed.Swallow);
            Assert.False(typed.State.IsArmed);
            Assert.Null(typed.Token);
        }
    }

    [Fact]
    public void Escape_timeout_invalid_key_and_lost_canvas_cancel()
    {
        SequenceState armed = Tap(SequenceState.Idle, Letter('A'), 1000, true).State;

        SequenceResult escape = Tap(armed, ShortcutSequence.PressFromLayout(0x1B, false, false, false, false, '\0'), 1100, true);
        Assert.True(escape.Swallow);
        Assert.Null(escape.Token);
        Assert.False(escape.State.IsArmed);

        SequenceResult invalid = Tap(armed, Letter('Q'), 1200, true);
        Assert.Null(invalid.Token);
        Assert.False(invalid.Swallow);
        Assert.False(invalid.State.IsArmed);

        SequenceResult away = Tap(armed, Letter('L'), 1300, allowArm: false);
        Assert.Null(away.Token);
        Assert.False(away.Swallow);

        SequenceResult late = Tap(armed, Letter('L'), 1000 + ShortcutSequence.TimeoutMs, true);
        Assert.Null(late.Token);
        Assert.False(late.State.IsArmed);
    }

    [Fact]
    public void Invalid_prefix_rearms_and_direct_insert_letter_is_released()
    {
        SequenceState armed = Tap(SequenceState.Idle, Letter('D'), 1000, true).State;
        SequenceResult again = Tap(armed, Letter('D'), 1100, true);
        Assert.True(again.State.IsArmed);
        Assert.Equal('D', again.State.Prefix);
        Assert.True(again.Swallow);
        Assert.Null(again.Token);

        SequenceState align = Tap(SequenceState.Idle, Letter('A'), 1000, true).State;
        SequenceResult insert = Tap(align, Letter('U'), 1100, true);
        Assert.False(insert.Swallow);
        Assert.Null(insert.Token);
        Assert.False(insert.State.IsArmed);
    }

    [Fact]
    public void Command_modifiers_do_not_arm_or_complete()
    {
        SequenceResult ctrlA = Tap(SequenceState.Idle, Letter('A', ctrl: true), 10, allowArm: false);
        Assert.False(ctrlA.State.IsArmed);

        SequenceState armed = Tap(SequenceState.Idle, Letter('A'), 10, true).State;
        SequenceResult ctrlL = Tap(armed, Letter('L', ctrl: true), 20, allowArm: false);
        Assert.Null(ctrlL.Token);
        Assert.False(ctrlL.Swallow);
    }

    [Fact]
    public void Direct_gate_still_rejects_sequence_tokens_and_the_prefix_letters()
    {
        Assert.False(ShortcutGate.IsToken("A L"));
        Assert.True(ShortcutSequence.IsSequenceToken("A L"));
        Assert.True(ShortcutGate.IsRunnable("A L"));
        Assert.True(ShortcutGate.IsRunnable("T"));
        Assert.False(ShortcutGate.IsRunnable("CMD /"));
        Assert.True(ShortcutGate.IsRunnable("CTRL /"));
        Assert.True(ShortcutGate.IsRunnable(". T"));
        Assert.False(ShortcutGate.Decide("/", false, false, FocusKind.Pane, Surface.Canvas).Dispatch);
        Assert.Null(ShortcutGate.TokenForVirtualKey(0x41));
        Assert.False(ShortcutGate.Decide("A", false, false, FocusKind.Pane, Surface.Canvas).Dispatch);
        Assert.False(ShortcutGate.Decide("A L", false, false, FocusKind.Pane, Surface.Canvas).Dispatch);
        Assert.True(ShortcutGate.Decide("R", false, false, FocusKind.Pane, Surface.Canvas).Dispatch);
        Assert.True(ShortcutGate.Decide("N", false, false, FocusKind.Pane, Surface.Slides).Dispatch);
    }

    [Fact]
    public void X_then_T_unifies_shape_type_only_on_the_canvas()
    {
        Assert.True(ShortcutSequence.IsPrefixLetter('X'));
        Assert.True(ShortcutSequence.IsPrefixVirtualKey(0x58));
        Assert.True(ShortcutSequence.IsSequenceToken("X T"));
        Assert.True(ShortcutSequence.IsSequenceToken("A A"));
        Assert.True(ShortcutGate.IsRunnable("X T"));
        Assert.Equal("X T", ShortcutSequence.Match('X', Letter('T')));
        Assert.Null(ShortcutSequence.Match('X', Letter('T', shift: true)));
        Assert.Null(ShortcutSequence.Match('X', Letter('T', ctrl: true)));

        SequenceResult armed = Tap(SequenceState.Idle, Letter('X'), 1000, allowArm: true);
        Assert.True(armed.Swallow);
        Assert.Null(armed.Token);
        Assert.Equal('X', armed.State.Prefix);

        SequenceResult unified = Tap(armed.State, Letter('T'), 1500, allowArm: true);
        Assert.Equal("X T", unified.Token);
        Assert.True(unified.Swallow);
        Assert.False(unified.State.IsArmed);

        SequenceResult offCanvas = Tap(SequenceState.Idle, Letter('X'), 1000, allowArm: false);
        Assert.False(offCanvas.Swallow);
        Assert.False(offCanvas.State.IsArmed);
        Assert.Null(offCanvas.Token);

        SequenceResult stillOff = Tap(offCanvas.State, Letter('T'), 1200, allowArm: false);
        Assert.False(stillOff.Swallow);
        Assert.Null(stillOff.Token);
        Assert.False(stillOff.State.IsArmed);
    }

    [Fact]
    public void Lone_X_arms_for_one_second_then_expires()
    {
        const uint armedAt = 5000;
        SequenceResult armed = Tap(SequenceState.Idle, Letter('X'), armedAt, allowArm: true);
        Assert.True(armed.State.IsArmed);
        Assert.Equal('X', armed.State.Prefix);
        Assert.True(armed.Swallow);
        Assert.Null(armed.Token);

        SequenceResult inTime = Tap(
            armed.State,
            Letter('T'),
            armedAt + ShortcutSequence.TimeoutMs - 1,
            allowArm: true);
        Assert.Equal("X T", inTime.Token);
        Assert.True(inTime.Swallow);
        Assert.False(inTime.State.IsArmed);

        SequenceResult armedAgain = Tap(SequenceState.Idle, Letter('X'), armedAt, allowArm: true);
        SequenceResult late = Tap(
            armedAgain.State,
            Letter('T'),
            armedAt + ShortcutSequence.TimeoutMs,
            allowArm: true);
        Assert.Null(late.Token);
        Assert.False(late.Swallow);
        Assert.False(late.State.IsArmed);

        SequenceResult rectangle = Tap(armedAgain.State, Letter('R'), armedAt + 100, allowArm: true);
        Assert.Null(rectangle.Token);
        Assert.False(rectangle.Swallow);
        Assert.False(rectangle.State.IsArmed);
    }

    [Fact]
    public void Align_pairs_period_T_and_alt_period_stay()
    {
        Assert.Equal("A X", Tap(Tap(SequenceState.Idle, Letter('A'), 10, true).State, Letter('X'), 20, true).Token);
        Assert.Equal("A A", Tap(Tap(SequenceState.Idle, Letter('A'), 10, true).State, Letter('A'), 20, true).Token);
        Assert.Equal("A T", Tap(Tap(SequenceState.Idle, Letter('A'), 10, true).State, Letter('T'), 20, true).Token);

        SequencePress period = ShortcutSequence.PressFromLayout(0xBE, false, false, false, false, '.');
        Assert.Equal(". T", Tap(Tap(SequenceState.Idle, period, 10, true).State, Letter('T'), 20, true).Token);

        Assert.Equal("ALT .", ShortcutChords.Token(KeyMark.Period, 0xBE, false, false, true, false));
        Assert.Null(ShortcutChords.Token(KeyMark.Period, 0xBE, true, false, true, false));
        Assert.Null(ShortcutChords.Token(KeyMark.Period, 0xBE, false, false, false, false));
    }

    [Fact]
    public void Period_then_T_merges_only_on_the_canvas()
    {
        SequencePress period = ShortcutSequence.PressFromLayout(0xBE, false, false, false, false, '.');
        Assert.Equal(SequenceSymbol.Period, period.Symbol);

        SequenceResult armed = Tap(SequenceState.Idle, period, 1000, allowArm: true);
        Assert.True(armed.Swallow);
        Assert.Equal('.', armed.State.Prefix);

        SequenceResult merged = Tap(armed.State, Letter('T'), 1200, true);
        Assert.Equal(". T", merged.Token);
        Assert.True(merged.Swallow);

        SequenceResult typed = Tap(SequenceState.Idle, period, 1000, allowArm: false);
        Assert.False(typed.Swallow);
        Assert.False(typed.State.IsArmed);

        SequencePress altPeriod = ShortcutSequence.PressFromLayout(0xBE, false, alt: true, false, false, '.');
        SequenceResult margins = Tap(SequenceState.Idle, altPeriod, 1000, allowArm: true);
        Assert.False(margins.State.IsArmed);
        Assert.False(margins.Swallow);
    }

    [Fact]
    public void Chords_require_their_modifiers_and_ignore_altgr()
    {
        Assert.Equal("/", ShortcutChords.Token(KeyMark.Slash, 0xBF, false, false, false, false));
        Assert.Equal("CTRL /", ShortcutChords.Token(KeyMark.Slash, 0xBF, true, false, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.Slash, 0xBF, true, false, true, false));
        Assert.Null(ShortcutChords.Token(KeyMark.Slash, 0xBF, false, false, false, true));
        Assert.Equal("ALT .", ShortcutChords.Token(KeyMark.Period, 0xBE, false, false, true, false));
        Assert.Null(ShortcutChords.Token(KeyMark.Period, 0xBE, false, false, false, false));
        Assert.Equal("CTRL ]", ShortcutChords.Token(KeyMark.CloseBracket, 0xDD, true, false, false, false));
        Assert.Equal("CTRL [", ShortcutChords.Token(KeyMark.OpenBracket, 0xDB, true, false, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.CloseBracket, 0xDD, false, false, false, false));
        Assert.Equal("CTRL SHIFT ]", ShortcutChords.Token(KeyMark.None, 0xDD, true, true, false, false, unshiftedCloseBracket: true));
        Assert.Equal("CTRL SHIFT [", ShortcutChords.Token(KeyMark.None, 0xDB, true, true, false, false, unshiftedOpenBracket: true));
        Assert.Equal("CTRL ]", ShortcutChords.Token(KeyMark.CloseBracket, 0xDD, true, true, false, false, unshiftedCloseBracket: true));
        Assert.Null(ShortcutChords.Token(KeyMark.None, 0xDD, true, true, true, false, unshiftedCloseBracket: true));
        Assert.False(ShortcutGate.IsToken("I"));
        Assert.True(ShortcutGate.IsRunnable("I"));
        Assert.Equal("I", ShortcutGate.DecideIcon(FocusKind.Pane, Surface.Canvas).Token);
        Assert.False(ShortcutGate.DecideIcon(FocusKind.Pane, Surface.Text).Dispatch);
        Assert.False(ShortcutGate.DecideIcon(FocusKind.Pane, Surface.Blocked).Dispatch);
        Assert.False(ShortcutGate.DecideIcon(FocusKind.RibbonOrEdit, Surface.Canvas).Dispatch);
        Assert.False(ShortcutGate.DecideIcon(FocusKind.MenuOrDialog, Surface.Canvas).Dispatch);
        Assert.False(ShortcutGate.DecideIcon(FocusKind.Outside, Surface.Canvas).Dispatch);
        Assert.False(ShortcutGate.Decide("I", false, false, FocusKind.Pane, Surface.Canvas).Dispatch);
        Assert.Equal("CTRL SHIFT C", ShortcutChords.Token(KeyMark.None, 0x43, true, true, false, false));
        Assert.Equal("CTRL SHIFT V", ShortcutChords.Token(KeyMark.None, 0x56, true, true, false, false));
        Assert.Equal("CTRL SHIFT T", ShortcutChords.Token(KeyMark.None, 0x54, true, true, false, false));
        Assert.True(ShortcutChords.IsChordToken("CTRL SHIFT T"));
        Assert.Null(ShortcutChords.Token(KeyMark.None, 0x54, true, false, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.None, 0x54, false, true, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.None, 0x43, true, false, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.None, 0x56, true, false, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.None, 0x43, true, true, true, false));
        Assert.Equal("CTRL 2", ShortcutChords.Token(KeyMark.None, 0x32, true, false, false, false));
        Assert.Equal("CTRL 3", ShortcutChords.Token(KeyMark.None, 0x63, true, false, false, false));
        Assert.Equal("ALT 2", ShortcutChords.Token(KeyMark.None, 0x32, false, false, true, false));
        Assert.Equal("ALT 3", ShortcutChords.Token(KeyMark.None, 0x33, false, false, true, false));
        Assert.Equal("CTRL 4", ShortcutChords.Token(KeyMark.None, 0x34, true, false, false, false));
        Assert.Equal("CTRL 5", ShortcutChords.Token(KeyMark.None, 0x35, true, false, false, false));
        Assert.Equal("CTRL 6", ShortcutChords.Token(KeyMark.None, 0x36, true, false, false, false));
        Assert.Equal("CTRL 7", ShortcutChords.Token(KeyMark.None, 0x37, true, false, false, false));
        Assert.Equal("CTRL K", ShortcutChords.Token(KeyMark.None, 0x4B, true, false, false, false));
        Assert.Equal("CTRL SPACE", ShortcutChords.Token(KeyMark.None, 0x20, true, false, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.None, 0x32, true, true, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.None, 0x32, true, false, true, false));
        Assert.Null(ShortcutChords.Token(KeyMark.None, 0x32, false, false, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.None, 0x4B, false, false, false, false));
        Assert.True(ShortcutChords.IsChordToken("CTRL 2"));
        Assert.True(ShortcutGate.IsRunnable("CTRL K"));
        Assert.True(ShortcutGate.IsRunnable("A P"));
        Assert.True(ShortcutSequence.IsSequenceToken("A X"));
        Assert.True(ShortcutSequence.IsSequenceToken("A Z"));
        Assert.False(ShortcutGate.IsToken("A P"));
    }
}
