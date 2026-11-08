using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

public class ShortcutGateTests
{
    static readonly string[] Tokens = ["T", "R", "U", "O", "C", "P", "G", "M", "N", "E"];

    [Fact]
    public void Canvas_dispatches_every_bare_token()
    {
        foreach (string token in Tokens)
        {
            GateDecision decision = ShortcutGate.Decide(token, false, false, FocusKind.Pane, Surface.Canvas);
            Assert.True(decision.Dispatch);
            Assert.True(decision.Swallow);
            Assert.Equal(token, decision.Token);
        }
    }

    [Theory]
    [InlineData(Surface.Text)]
    [InlineData(Surface.Blocked)]
    [InlineData(Surface.Unknown)]
    public void Text_blocked_and_unknown_pass(Surface surface)
    {
        foreach (string token in Tokens)
        {
            GateDecision decision = ShortcutGate.Decide(token, false, false, FocusKind.Pane, surface);
            Assert.False(decision.Dispatch);
            Assert.False(decision.Swallow);
        }
    }

    [Fact]
    public void Thumbnail_and_sorter_allow_only_N()
    {
        foreach (string token in Tokens)
        {
            GateDecision decision = ShortcutGate.Decide(token, false, false, FocusKind.Pane, Surface.Slides);
            Assert.Equal(token == "N", decision.Dispatch);
        }
    }

    [Theory]
    [InlineData(FocusKind.Outside)]
    [InlineData(FocusKind.NotFrame)]
    [InlineData(FocusKind.MenuOrDialog)]
    [InlineData(FocusKind.RibbonOrEdit)]
    [InlineData(FocusKind.Uncertain)]
    public void Non_pane_focus_passes_even_on_canvas(FocusKind focus)
    {
        GateDecision decision = ShortcutGate.Decide("R", false, false, focus, Surface.Canvas);
        Assert.False(decision.Dispatch);
    }

    [Fact]
    public void Modifiers_and_injected_keys_pass()
    {
        Assert.False(ShortcutGate.Decide("T", modified: true, injected: false, FocusKind.Pane, Surface.Canvas).Dispatch);
        Assert.False(ShortcutGate.Decide("T", modified: false, injected: true, FocusKind.Pane, Surface.Canvas).Dispatch);
        Assert.True(ShortcutGate.Decide("Q", false, false, FocusKind.Pane, Surface.Canvas).Dispatch);
        Assert.False(ShortcutGate.Decide("Q", false, false, FocusKind.Pane, Surface.Text).Dispatch);
        Assert.False(ShortcutGate.Decide("Q", false, false, FocusKind.Pane, Surface.Slides).Dispatch);
    }

    [Fact]
    public void Virtual_keys_map_the_canvas_letters()
    {
        Assert.Equal("T", ShortcutGate.TokenForVirtualKey(0x54));
        Assert.Equal("R", ShortcutGate.TokenForVirtualKey(0x52));
        Assert.Equal("U", ShortcutGate.TokenForVirtualKey(0x55));
        Assert.Equal("O", ShortcutGate.TokenForVirtualKey(0x4F));
        Assert.Equal("C", ShortcutGate.TokenForVirtualKey(0x43));
        Assert.Equal("P", ShortcutGate.TokenForVirtualKey(0x50));
        Assert.Equal("G", ShortcutGate.TokenForVirtualKey(0x47));
        Assert.Equal("M", ShortcutGate.TokenForVirtualKey(0x4D));
        Assert.Equal("N", ShortcutGate.TokenForVirtualKey(0x4E));
        Assert.Equal("E", ShortcutGate.TokenForVirtualKey(0x45));
        Assert.Equal("Q", ShortcutGate.TokenForVirtualKey(0x51));
        Assert.Null(ShortcutGate.TokenForVirtualKey(0x41));
        Assert.False(ShortcutGate.Decide("E", modified: true, injected: false, FocusKind.Pane, Surface.Canvas).Dispatch);
        Assert.False(ShortcutGate.Decide("E", false, false, FocusKind.Pane, Surface.Text).Dispatch);
        Assert.True(ShortcutGate.Decide("E", false, false, FocusKind.Pane, Surface.Canvas).Dispatch);
    }

    [Fact]
    public void Classify_matches_the_windows_pane_matrix()
    {
        Assert.Equal(Surface.Blocked, ShortcutGate.Classify(slideshow: true, ShortcutGate.ViewSlide, ShortcutGate.SelectionNone, true));
        Assert.Equal(Surface.Unknown, ShortcutGate.Classify(false, ShortcutGate.ViewSlide, ShortcutGate.SelectionNone, readOk: false));
        Assert.Equal(Surface.Text, ShortcutGate.Classify(false, ShortcutGate.ViewSlide, ShortcutGate.SelectionText, true));
        Assert.Equal(Surface.Text, ShortcutGate.Classify(false, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionText, true));
        Assert.Equal(Surface.Text, ShortcutGate.Classify(false, ShortcutGate.ViewNotesPage, ShortcutGate.SelectionText, true));
        Assert.Equal(Surface.Canvas, ShortcutGate.Classify(false, ShortcutGate.ViewSlide, ShortcutGate.SelectionNone, true));
        Assert.Equal(Surface.Canvas, ShortcutGate.Classify(false, ShortcutGate.ViewSlide, ShortcutGate.SelectionShapes, true));
        Assert.Equal(Surface.Slides, ShortcutGate.Classify(false, ShortcutGate.ViewSlide, ShortcutGate.SelectionSlides, true));
        Assert.Equal(Surface.Slides, ShortcutGate.Classify(false, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionNone, true));
        Assert.Equal(Surface.Slides, ShortcutGate.Classify(false, ShortcutGate.ViewSlideSorter, ShortcutGate.SelectionSlides, true));
        Assert.Equal(Surface.Blocked, ShortcutGate.Classify(false, ShortcutGate.ViewNotesPage, ShortcutGate.SelectionShapes, true));
        Assert.Equal(Surface.Blocked, ShortcutGate.Classify(false, ShortcutGate.ViewOutline, ShortcutGate.SelectionSlides, true));
        Assert.Equal(Surface.Blocked, ShortcutGate.Classify(false, ShortcutGate.ViewPrintPreview, ShortcutGate.SelectionNone, true));
        Assert.Equal(Surface.Unknown, ShortcutGate.Classify(false, ShortcutGate.ViewNormal, ShortcutGate.SelectionShapes, true));
        Assert.Equal(0, ShortcutGate.SelectionNone);
        Assert.Equal(1, ShortcutGate.SelectionSlides);
        Assert.Equal(2, ShortcutGate.SelectionShapes);
        Assert.Equal(3, ShortcutGate.SelectionText);
        Assert.Equal(Surface.Text, ShortcutGate.Classify(false, 1, 3, true));
        Assert.Equal(Surface.Canvas, ShortcutGate.Classify(false, 1, 2, true));
        Assert.Equal(Surface.Canvas, ShortcutGate.Classify(false, 1, 0, true));
        Assert.Equal(Surface.Slides, ShortcutGate.Classify(false, 1, 1, true));
        Assert.Equal(Surface.Slides, ShortcutGate.Classify(false, 11, 0, true));
        Assert.Equal(Surface.Blocked, ShortcutGate.Classify(false, 3, 2, true));
    }

    [Fact]
    public void Text_hold_ignores_stale_canvas_until_a_selection_event()
    {
        HoldResult armed = ShortcutGate.Reduce(HoldState.Armed, Surface.Canvas, "stale", false, false, false);
        Assert.True(armed.Hold.Held);
        Assert.False(armed.Hold.ReleaseAllowed);
        Assert.Equal(Surface.Text, armed.Surface);

        HoldResult finished = ShortcutGate.Reduce(HoldState.Armed, Surface.Text, "pane=1 selection=2", false, runFinished: true, false);
        Assert.False(finished.Hold.Held);
        Assert.Equal(Surface.Text, finished.Surface);

        HoldResult waiting = ShortcutGate.Reduce(HoldState.Armed, Surface.Canvas, "stale", false, runFinished: true, false);
        Assert.Equal(HoldState.WaitingForExit, waiting.Hold);
        Assert.Equal(Surface.Text, waiting.Surface);

        HoldResult refresh = ShortcutGate.Reduce(waiting.Hold, Surface.Canvas, "refresh", false, false, selectionEvent: false);
        Assert.Equal(HoldState.WaitingForExit, refresh.Hold);
        Assert.Equal(Surface.Text, refresh.Surface);

        HoldResult left = ShortcutGate.Reduce(waiting.Hold, Surface.Canvas, "pane=1 selection=1", false, false, selectionEvent: true);
        Assert.False(left.Hold.Held);
        Assert.Equal(Surface.Canvas, left.Surface);

        HoldResult show = ShortcutGate.Reduce(waiting.Hold, Surface.Blocked, "slideshow", slideshow: true, false, false);
        Assert.False(show.Hold.Held);
        Assert.Equal(Surface.Blocked, show.Surface);
    }

    [Fact]
    public void Queued_text_is_the_typed_letter_without_ctrl_alt_or_windows()
    {
        Assert.True(ShortcutGate.TryQueueText(0x52, false, false, false, false, false, out char lower));
        Assert.Equal('r', lower);
        Assert.True(ShortcutGate.TryQueueText(0x52, false, false, false, shift: true, false, out char upper));
        Assert.Equal('R', upper);
        Assert.False(ShortcutGate.TryQueueText(0x52, ctrl: true, false, false, false, false, out _));
        Assert.False(ShortcutGate.TryQueueText(0x52, false, alt: true, false, false, false, out _));
        Assert.False(ShortcutGate.TryQueueText(0x52, false, false, win: true, false, false, out _));
        Assert.True(ShortcutGate.TryQueueText(0x20, false, false, false, false, false, out char space));
        Assert.Equal(' ', space);
        Assert.False(ShortcutGate.TryQueueText(0x31, false, false, false, shift: true, false, out _));
    }

    [Fact]
    public void Backslash_saves_only_selected_thumbnail_slides()
    {
        Assert.True(ShortcutGate.IsBareBackslash(KeyMark.Backslash, false, false, false));
        Assert.False(ShortcutGate.IsBareBackslash(KeyMark.Backslash, ctrl: true, false, false));
        Assert.False(ShortcutGate.IsBareBackslash(KeyMark.Backslash, false, alt: true, false));
        Assert.False(ShortcutGate.IsBareBackslash(KeyMark.Backslash, true, true, false));
        Assert.False(ShortcutGate.IsBareBackslash(KeyMark.Backslash, false, false, win: true));
        Assert.False(ShortcutGate.IsBareBackslash(KeyMark.Slash, false, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.Backslash, 0xDC, false, false, false, false));
        Assert.Null(ShortcutChords.Token(KeyMark.Backslash, 0xDC, true, true, false, false));
        Assert.Equal("/", ShortcutChords.Token(KeyMark.Slash, 0xBF, false, false, false, false));

        Assert.False(ShortcutGate.IsToken("\\"));
        Assert.False(ShortcutChords.IsChordToken("\\"));
        Assert.True(ShortcutGate.IsSaveSlidesToken("\\"));
        Assert.True(ShortcutGate.IsRunnable("\\"));
        Assert.False(ShortcutGate.Decide("\\", false, false, FocusKind.Pane, Surface.Canvas).Dispatch);
        Assert.False(ShortcutGate.Decide("\\", false, false, FocusKind.Pane, Surface.Slides).Dispatch);

        GateDecision saved = ShortcutGate.DecideSaveSlides(
            FocusKind.Pane, Surface.Slides, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionSlides);
        Assert.True(saved.Dispatch);
        Assert.True(saved.Swallow);
        Assert.Equal("\\", saved.Token);

        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.Pane, Surface.Slides, ShortcutGate.ViewSlideSorter, ShortcutGate.SelectionSlides).Dispatch);
        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.Pane, Surface.Slides, ShortcutGate.ViewSlide, ShortcutGate.SelectionSlides).Dispatch);
        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.Pane, Surface.Slides, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionNone).Dispatch);
        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.Pane, Surface.Slides, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionShapes).Dispatch);
        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.Pane, Surface.Text, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionText).Dispatch);
        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.Pane, Surface.Canvas, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionSlides).Dispatch);
        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.Pane, Surface.Blocked, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionSlides).Dispatch);
        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.Pane, Surface.Unknown, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionSlides).Dispatch);
        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.MenuOrDialog, Surface.Slides, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionSlides).Dispatch);
        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.RibbonOrEdit, Surface.Slides, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionSlides).Dispatch);
        Assert.False(ShortcutGate.DecideSaveSlides(
            FocusKind.Outside, Surface.Slides, ShortcutGate.ViewThumbnails, ShortcutGate.SelectionSlides).Dispatch);
        Assert.False(ShortcutGate.IsRunnable(";"));
        Assert.False(ShortcutGate.IsRunnable("EMAIL SLIDES"));
    }
}
