using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

public class ReplayGuardTests
{
    static ReplayOrigin Origin() => new(100, 42);

    static ReplayObservation Active(FocusKind focus = FocusKind.Pane, bool menu = false, bool textEdit = false) =>
        new(true, 100, 42, true, focus, menu, textEdit);

    [Fact]
    public void PowerPoint_pane_remains_an_allowed_replay_target()
    {
        Assert.True(ReplayGuard.TryCreateOrigin(true, 100, 42, true, FocusKind.Pane, out ReplayOrigin origin));
        Assert.Equal(ReplayDecision.Send, ReplayGuard.Decide(origin, Active(), powerpointRunning: true, pendingCharacters: 1));
        Assert.True(ReplayGuard.CanReplay(origin, Active(FocusKind.RibbonOrEdit, textEdit: true), powerpointRunning: true));
    }

    [Fact]
    public void Another_application_cancels_replay()
    {
        var other = new ReplayObservation(false, 200, 99, false, FocusKind.Outside, false, false);
        Assert.Equal(ReplayDecision.Discard, ReplayGuard.Decide(Origin(), other, powerpointRunning: true, pendingCharacters: 4));
        var otherPowerPoint = Active() with { ProcessId = 300, FrameWindow = 70 };
        Assert.False(ReplayGuard.CanReplay(Origin(), otherPowerPoint, powerpointRunning: true));
    }

    [Fact]
    public void PowerPoint_dialog_cancels_replay()
    {
        var dialogWindow = new ReplayObservation(true, 100, 77, false, FocusKind.NotFrame, true, false);
        var menu = Active(FocusKind.MenuOrDialog, menu: true);
        var fileDialog = Active(FocusKind.RibbonOrEdit, menu: true);
        Assert.False(ReplayGuard.CanReplay(Origin(), dialogWindow, powerpointRunning: true));
        Assert.False(ReplayGuard.CanReplay(Origin(), menu, powerpointRunning: true));
        Assert.False(ReplayGuard.CanReplay(Origin(), fileDialog, powerpointRunning: true));
    }

    [Fact]
    public void PowerPoint_closing_cancels_replay()
    {
        Assert.Equal(ReplayDecision.Discard, ReplayGuard.Decide(Origin(), Active(), powerpointRunning: false, pendingCharacters: 2));
        var gone = new ReplayObservation(false, 0, 0, false, FocusKind.Outside, false, false);
        Assert.False(ReplayGuard.CanReplay(Origin(), gone, powerpointRunning: true));
    }

    [Fact]
    public void Multiple_queued_keystrokes_follow_the_same_target_rule()
    {
        Assert.Equal(ReplayDecision.Send, ReplayGuard.Decide(Origin(), Active(), powerpointRunning: true, pendingCharacters: 5));
        Assert.Equal(ReplayDecision.Discard, ReplayGuard.Decide(Origin(), Active(), powerpointRunning: true, pendingCharacters: 0));
        var elsewhere = new ReplayObservation(false, 8, 8, false, FocusKind.Outside, false, false);
        Assert.Equal(ReplayDecision.Discard, ReplayGuard.Decide(Origin(), elsewhere, powerpointRunning: true, pendingCharacters: 5));
    }

    [Fact]
    public void Ribbon_and_uncertain_controls_are_not_replay_targets()
    {
        Assert.False(ReplayGuard.CanReplay(Origin(), Active(FocusKind.RibbonOrEdit), powerpointRunning: true));
        Assert.False(ReplayGuard.CanReplay(Origin(), Active(FocusKind.Uncertain), powerpointRunning: true));
        Assert.False(ReplayGuard.TryCreateOrigin(true, 100, 42, true, FocusKind.MenuOrDialog, out _));
        Assert.False(ReplayGuard.TryCreateOrigin(true, 100, 0, true, FocusKind.Pane, out _));
    }

    [Fact]
    public void Existing_canvas_sequence_still_resolves()
    {
        SequenceResult armed = ShortcutSequence.Next(
            SequenceState.Idle,
            ShortcutSequence.PressFromLayout(0x41, false, false, false, false, '\0'),
            now: 10,
            allowArm: true);
        SequenceResult completed = ShortcutSequence.Next(
            armed.State,
            ShortcutSequence.PressFromLayout(0x4C, false, false, false, false, '\0'),
            now: 20,
            allowArm: true);
        Assert.Equal("A L", completed.Token);
        Assert.False(ReplayGuard.CanReplay(
            Origin(),
            new ReplayObservation(false, 1, 1, false, FocusKind.Outside, false, false),
            powerpointRunning: true));
    }
}
