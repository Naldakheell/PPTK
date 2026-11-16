using System.Text.Json.Nodes;
using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

public class UpdateNoticeTests
{
    static readonly NoticeObservation Idle = new(PowerPointForeground: true, Slideshow: false, DialogOrMenu: false, SinceLastKey: TimeSpan.FromSeconds(30));

    // ---- read-only inspection

    [Fact]
    public void Inspect_reports_a_verified_pending_update_and_writes_nothing()
    {
        using var f = new PendingFixture(keysChanged: true);
        f.SetProcesses("POWERPNT");   // PowerPoint open: inspection does not care
        var before = f.Snapshot();
        var i = f.Engine().Inspect();
        Assert.True(i.HasPending && i.Valid, i.Problem);
        Assert.Equal(f.Pending, i.Version);
        Assert.False(i.Notified);
        Assert.Equal(before, f.Snapshot());
        Assert.False(File.Exists(f.Paths.LockPath));   // not even the lock file is created
    }

    [Theory]
    [InlineData("sha")]
    [InlineData("missing-ppam")]
    [InlineData("keys-flag")]
    [InlineData("version")]
    [InlineData("installed-addin")]
    [InlineData("journal")]
    public void Inspect_reports_an_unverified_pending_update_and_writes_nothing(string damage)
    {
        using var f = new PendingFixture();
        var s = f.State();
        switch (damage)
        {
            case "sha": File.AppendAllText(f.PendingFile, "tampered"); break;
            case "missing-ppam": File.Delete(f.PendingFile); break;
            case "keys-flag": s["pending"]!["keysChanged"] = true; File.WriteAllText(f.Paths.StatePath, s.ToJsonString()); break;
            case "version": s["pending"]!["version"] = "0.9.0"; File.WriteAllText(f.Paths.StatePath, s.ToJsonString()); break;
            case "installed-addin": File.WriteAllText(f.Paths.AddinFile, "someone else's add-in"); break;
            case "journal": File.WriteAllText(f.Paths.JournalPath, "{\"operation\":\"install\"}"); break;
        }
        var before = f.Snapshot();
        var i = f.Engine().Inspect();
        Assert.False(i.Valid);
        Assert.NotNull(i.Problem);
        Assert.Equal(before, f.Snapshot());                       // no lastError, nothing else
        Assert.Equal(NoticeAction.None, NotificationPolicy.Decide(i, Idle).Action);
    }

    [Theory]
    [InlineData("{ not json")]
    [InlineData("[1]")]
    public void Inspect_of_a_corrupt_state_is_not_verified(string text)
    {
        using var f = new PendingFixture();
        File.WriteAllText(f.Paths.StatePath, text);
        var i = f.Engine().Inspect();
        Assert.False(i.HasPending || i.Valid);
        Assert.Equal(NoticeAction.None, NotificationPolicy.Decide(i, Idle).Action);
    }

    [Fact]
    public void After_the_update_is_applied_there_is_nothing_to_announce()
    {
        using var f = new PendingFixture();
        Assert.Equal(PendingUpdateOutcome.Applied, f.Engine().TryApply().Outcome);
        var i = f.Engine().Inspect();
        Assert.False(i.HasPending);
        Assert.Equal(NoticeAction.None, NotificationPolicy.Decide(i, Idle).Action);
    }

    // ---- policy

    [Fact]
    public void Policy_shows_only_when_PowerPoint_is_in_front_and_the_user_is_not_busy()
    {
        var p = new PendingInspection(true, true, "1.0.1", false, null);
        Assert.Equal(NoticeAction.Show, NotificationPolicy.Decide(p, Idle).Action);
        Assert.Equal(NoticeAction.Wait, NotificationPolicy.Decide(p, Idle with { PowerPointForeground = false }).Action);
        Assert.Equal(NoticeAction.Wait, NotificationPolicy.Decide(p, Idle with { Slideshow = true }).Action);
        Assert.Equal(NoticeAction.Wait, NotificationPolicy.Decide(p, Idle with { DialogOrMenu = true }).Action);
        Assert.Equal(NoticeAction.Wait, NotificationPolicy.Decide(p, Idle with { SinceLastKey = TimeSpan.FromSeconds(2) }).Action);
        Assert.Equal(NoticeAction.None, NotificationPolicy.Decide(p with { Notified = true }, Idle).Action);
        Assert.Equal(NoticeAction.None, NotificationPolicy.Decide(p with { Valid = false, Problem = "x" }, Idle).Action);
        Assert.Equal(NoticeAction.None, NotificationPolicy.Decide(new PendingInspection(false, false, null, false, "no pending update"), Idle).Action);
        Assert.Equal("1.0.1", NotificationPolicy.Decide(p, Idle).Version);
    }

    [Fact]
    public void Typing_delays_the_notice_until_the_user_pauses()
    {
        var p = new PendingInspection(true, true, "1.0.1", false, null);
        var scheduler = new NoticeScheduler();
        // Ticks while the user types (last key 0-4 s ago): wait. Then 5 s of quiet: show.
        foreach (var since in new[] { 0.0, 1.2, 0.3, 4.9 })
        {
            var d = scheduler.Next(p, Idle with { SinceLastKey = TimeSpan.FromSeconds(since) });
            Assert.Equal(NoticeAction.Wait, d.Action);
            Assert.Equal("the user is typing", d.Reason);
        }
        Assert.Equal(NoticeAction.Show, scheduler.Next(p, Idle with { SinceLastKey = TimeSpan.FromSeconds(5) }).Action);
    }

    [Fact]
    public void A_slideshow_or_dialog_delays_the_notice_and_it_appears_afterwards()
    {
        var p = new PendingInspection(true, true, "1.0.1", false, null);
        var scheduler = new NoticeScheduler();
        Assert.Equal(NoticeAction.Wait, scheduler.Next(p, Idle with { Slideshow = true }).Action);
        Assert.Equal(NoticeAction.Wait, scheduler.Next(p, Idle with { DialogOrMenu = true }).Action);
        Assert.Equal(NoticeAction.Wait, scheduler.Next(p, Idle with { PowerPointForeground = false }).Action);
        Assert.Equal(NoticeAction.Show, scheduler.Next(p, Idle).Action);
    }

    // ---- the "shown" flag: only after the window appeared, once per version

    [Fact]
    public void Shown_notice_is_recorded_once_and_not_shown_again()
    {
        using var f = new PendingFixture();
        var scheduler = new NoticeScheduler();
        var first = scheduler.Next(f.Engine().Inspect(), Idle);
        Assert.Equal(NoticeAction.Show, first.Action);
        // (The window appears here in 2C-2/2C-3; only then is the flag written.)
        Assert.False(f.Engine().Inspect().Notified);
        Assert.Equal(MarkResult.Marked, scheduler.OnShown(first.Version!, v => PendingNotice.MarkNotified(f.Paths, v)));
        var after = f.Engine().Inspect();
        Assert.True(after.Notified && after.Valid);
        Assert.Equal(NoticeAction.None, scheduler.Next(after, Idle).Action);
        Assert.Equal(NoticeAction.None, new NoticeScheduler().Next(after, Idle).Action);   // a restarted helper: still not again
        Assert.Equal(MarkResult.AlreadyMarked, PendingNotice.MarkNotified(f.Paths, f.Pending));
        var s = f.State();
        Assert.Equal(f.Installed, (string?)s["installed"]);        // only the flag changed
        Assert.NotNull((string?)s["pending"]!["notifiedUtc"]);
    }

    [Fact]
    public void A_notice_that_was_never_shown_is_not_recorded()
    {
        using var f = new PendingFixture();
        var scheduler = new NoticeScheduler();
        Assert.Equal(NoticeAction.Show, scheduler.Next(f.Engine().Inspect(), Idle).Action);
        // The window failed to appear: OnShown is not called, so the flag stays false and the next tick tries again.
        Assert.False(f.Engine().Inspect().Notified);
        Assert.Equal(NoticeAction.Show, scheduler.Next(f.Engine().Inspect(), Idle).Action);
    }

    [Fact]
    public void Busy_lock_after_showing_retries_the_flag_without_showing_again()
    {
        using var f = new PendingFixture();
        var scheduler = new NoticeScheduler();
        var d = scheduler.Next(f.Engine().Inspect(), Idle);
        MarkResult result;
        using (new FileStream(f.Paths.LockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))   // an installer runs
            result = scheduler.OnShown(d.Version!, v => PendingNotice.MarkNotified(f.Paths, v));
        Assert.Equal(MarkResult.Busy, result);
        Assert.False(f.Engine().Inspect().Notified);
        var retry = scheduler.Next(f.Engine().Inspect(), Idle);
        Assert.Equal(NoticeAction.MarkShown, retry.Action);              // save the flag; do not show a second time
        Assert.Equal(MarkResult.Marked, scheduler.OnShown(retry.Version!, v => PendingNotice.MarkNotified(f.Paths, v)));
        Assert.Equal(NoticeAction.None, scheduler.Next(f.Engine().Inspect(), Idle).Action);
    }

    [Fact]
    public void Flag_is_refused_for_another_version_or_without_a_pending_update()
    {
        using var f = new PendingFixture();
        var before = File.ReadAllText(f.Paths.StatePath);
        Assert.Equal(MarkResult.VersionMismatch, PendingNotice.MarkNotified(f.Paths, "1.0.2"));
        File.WriteAllText(f.Paths.JournalPath, "{\"operation\":\"install\"}");
        Assert.Equal(MarkResult.Busy, PendingNotice.MarkNotified(f.Paths, f.Pending));   // interrupted operation: leave it alone
        File.Delete(f.Paths.JournalPath);
        Assert.Equal(before, File.ReadAllText(f.Paths.StatePath));
        Assert.Equal(PendingUpdateOutcome.Applied, f.Engine().TryApply().Outcome);
        Assert.Equal(MarkResult.NoPending, PendingNotice.MarkNotified(f.Paths, f.Pending));
    }

    [Fact]
    public void A_newer_pending_version_is_announced_again()
    {
        using var f = new PendingFixture();
        var scheduler = new NoticeScheduler();
        var d = scheduler.Next(f.Engine().Inspect(), Idle);
        scheduler.OnShown(d.Version!, v => PendingNotice.MarkNotified(f.Paths, v));
        Assert.Equal(NoticeAction.None, scheduler.Next(f.Engine().Inspect(), Idle).Action);
        // Intune replaces the staged update with a newer one (not yet announced).
        var next = new PendingInspection(true, true, "1.0.2", false, null);
        Assert.Equal(NoticeAction.Show, scheduler.Next(next, Idle).Action);
    }

    [Fact]
    public void The_flag_does_not_disturb_applying_the_update()
    {
        using var f = new PendingFixture(keysChanged: true);
        Assert.Equal(MarkResult.Marked, PendingNotice.MarkNotified(f.Paths, f.Pending));
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Applied, r.Outcome);
        Assert.Null(f.State()["pending"]);
    }

    [Fact]
    public void Notice_text_and_duration_are_the_approved_ones()
    {
        Assert.Equal("A PPTK update is ready.\nPlease close and reopen PowerPoint\nto complete the update.", UpdateNoticeText.Message);
        Assert.Equal(TimeSpan.FromSeconds(10), UpdateNoticeText.DisplayFor);
        Assert.DoesNotContain("Apply", UpdateNoticeText.Message);
        Assert.DoesNotContain("Download", UpdateNoticeText.Message);
    }
}
