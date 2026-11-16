using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

/// <summary>Phase 2C-3: the notice as the helper drives it (fake window and observations, real state files).</summary>
public class UpdateNoticeControllerTests
{
    static readonly NoticeObservation Idle = new(true, false, false, TimeSpan.FromSeconds(30));
    static readonly PendingInspection Verified = new(true, true, "1.0.1", false, null);

    sealed class Harness
    {
        public NoticeObservation Seen = Idle;
        public int Shows, Marks;
        public Func<bool> ShowResult = () => true;
        public Func<string, MarkResult> MarkResult = _ => PPTKKeys.MarkResult.Marked;
        public readonly List<string> Log = [];
        public readonly Queue<Action> Background = new();
        public bool RunBackgroundAtOnce = true;
        public readonly UpdateNoticeController Controller;

        public Harness()
        {
            Controller = new UpdateNoticeController(() => Seen, _ => { Shows++; return ShowResult(); },
                v => { Marks++; return MarkResult(v); },
                work => { if (RunBackgroundAtOnce) work(); else Background.Enqueue(work); },
                m => Log.Add(m));
        }
        public void Tick(int times = 1) { for (int i = 0; i < times; i++) Controller.Tick(); }
        public void DrainBackground() { while (Background.TryDequeue(out var w)) w(); }
    }

    [Fact]
    public void Nothing_is_shown_before_the_first_inspection_or_without_a_verified_update()
    {
        var h = new Harness();
        h.Tick(3);
        h.Controller.OnInspection(new PendingInspection(false, false, null, false, "no pending update"));
        h.Tick(3);
        h.Controller.OnInspection(Verified with { Valid = false, Problem = "SHA-256 mismatch" });
        h.Tick(3);
        Assert.Equal(0, h.Shows);
    }

    [Fact]
    public void A_verified_pending_update_is_shown_once_even_while_the_inspection_is_stale()
    {
        var h = new Harness();
        h.Controller.OnInspection(Verified);
        h.Tick(10);                          // the inspection still says notified=false (refreshed only every few seconds)
        Assert.Equal(1, h.Shows);
        Assert.Equal(1, h.Marks);
        h.Controller.OnInspection(Verified with { Notified = true });
        h.Tick(10);
        Assert.Equal(1, h.Shows);
    }

    [Fact]
    public void Typing_delays_the_notice_until_a_pause()
    {
        var h = new Harness();
        h.Controller.OnInspection(Verified);
        foreach (var s in new[] { 0.1, 2.0, 0.5, 4.9 }) { h.Seen = Idle with { SinceLastKey = TimeSpan.FromSeconds(s) }; h.Tick(); }
        Assert.Equal(0, h.Shows);
        Assert.Contains(h.Log, l => l.Contains("the user is typing"));
        h.Seen = Idle with { SinceLastKey = TimeSpan.FromSeconds(5.1) };
        h.Tick();
        Assert.Equal(1, h.Shows);
    }

    [Theory]
    [InlineData("slideshow")]
    [InlineData("dialog")]
    [InlineData("background")]
    public void Slideshow_dialog_or_PowerPoint_in_the_background_delay_the_notice(string busy)
    {
        var h = new Harness();
        h.Controller.OnInspection(Verified);
        h.Seen = busy switch { "slideshow" => Idle with { Slideshow = true }, "dialog" => Idle with { DialogOrMenu = true }, _ => Idle with { PowerPointForeground = false } };
        h.Tick(5);
        Assert.Equal(0, h.Shows);
        h.Seen = Idle;
        h.Tick();
        Assert.Equal(1, h.Shows);
    }

    [Fact]
    public void A_window_that_did_not_appear_is_not_recorded_and_is_tried_again()
    {
        var h = new Harness { ShowResult = () => false };
        h.Controller.OnInspection(Verified);
        h.Tick();
        Assert.Equal(1, h.Shows);
        Assert.Equal(0, h.Marks);
        Assert.Contains(h.Log, l => l.Contains("did not appear; not recorded"));
        h.ShowResult = () => true;
        h.Tick();
        Assert.Equal(2, h.Shows);
        Assert.Equal(1, h.Marks);
    }

    [Fact]
    public void A_window_that_throws_is_logged_and_never_breaks_the_tick()
    {
        var h = new Harness { ShowResult = () => throw new InvalidOperationException("no desktop") };
        h.Controller.OnInspection(Verified);
        h.Tick(3);                            // no exception escapes
        Assert.Equal(0, h.Marks);
        Assert.Contains(h.Log, l => l.Contains("could not be shown (InvalidOperationException: no desktop)"));
    }

    [Fact]
    public void A_busy_flag_write_is_retried_without_showing_again()
    {
        var h = new Harness { MarkResult = _ => MarkResult.Busy };
        h.Controller.OnInspection(Verified);
        h.Tick();
        Assert.Equal(1, h.Shows);
        h.Tick(3);                            // retries the flag
        Assert.Equal(1, h.Shows);
        Assert.True(h.Marks >= 2);
        h.MarkResult = _ => MarkResult.Marked;
        h.Tick();
        int marks = h.Marks;
        h.Tick(5);
        Assert.Equal(marks, h.Marks);         // saved: no more writes
        Assert.Equal(1, h.Shows);
    }

    [Fact]
    public void A_failing_flag_write_is_logged_and_never_breaks_the_tick()
    {
        var h = new Harness { MarkResult = _ => throw new IOException("disk") };
        h.Controller.OnInspection(Verified);
        h.Tick(3);
        Assert.Equal(1, h.Shows);
        Assert.Contains(h.Log, l => l.Contains("flag not saved (IOException: disk)"));
    }

    [Fact]
    public void Ticks_while_the_flag_is_being_written_never_show_twice_or_write_twice_at_once()
    {
        var h = new Harness { RunBackgroundAtOnce = false };
        h.Controller.OnInspection(Verified);
        h.Tick(5);                            // the write is still queued
        Assert.Equal(1, h.Shows);
        Assert.Single(h.Background);
        h.DrainBackground();
        h.Tick(5);
        Assert.Equal(1, h.Shows);
        Assert.Equal(1, h.Marks);
    }

    [Fact]
    public void A_newer_pending_version_is_announced_once_more()
    {
        var h = new Harness();
        h.Controller.OnInspection(Verified);
        h.Tick();
        h.Controller.OnInspection(Verified with { Notified = true });
        h.Tick(3);
        h.Controller.OnInspection(new PendingInspection(true, true, "1.0.2", false, null));
        h.Tick(3);
        Assert.Equal(2, h.Shows);
    }

    // ---- with the watcher and real state files

    [Fact]
    public async Task Watcher_feeds_the_notice_only_while_PowerPoint_is_open_and_the_flag_survives_a_helper_restart()
    {
        using var f = new PendingFixture();
        var probe = new FakeProbe();
        probe.Processes.Add(new PowerPointProcess(1, DateTime.UtcNow, true, "explorer.exe"));
        int shows = 0;
        UpdateNoticeController Controller() => new(() => Idle, _ => { shows++; return true; }, v => PendingNotice.MarkNotified(f.Paths, v), w => w(), _ => { });

        var notice = Controller();
        var watcher = new PendingUpdateWatcher(f.Paths, () => f.Engine(), probe, _ => { }, onInspection: notice.OnInspection);
        Assert.True(watcher.Trigger("poll")); await watcher.Current.WaitAsync(TimeSpan.FromSeconds(30));
        notice.Tick();
        Assert.Equal(1, shows);
        Assert.True(f.Engine().Inspect().Notified);            // saved in state.json
        Assert.Equal(0, watcher.EngineRuns);                    // PowerPoint open: the engine waited

        // Helper restart: a new watcher and controller read notified=true and stay quiet.
        var restarted = Controller();
        var watcher2 = new PendingUpdateWatcher(f.Paths, () => f.Engine(), probe, _ => { }, onInspection: restarted.OnInspection);
        Assert.True(watcher2.Trigger("startup")); await watcher2.Current.WaitAsync(TimeSpan.FromSeconds(30));
        restarted.Tick(); restarted.Tick();
        Assert.Equal(1, shows);

        // PowerPoint closes: the update applies; afterwards there is nothing to announce.
        probe.Processes.Clear();
        Assert.True(watcher2.Trigger("poll")); await watcher2.Current.WaitAsync(TimeSpan.FromSeconds(30));
        Assert.Equal(PendingUpdateOutcome.Applied, watcher2.LastResult!.Outcome);
        restarted.Tick();
        Assert.Equal(1, shows);
    }

    [Fact]
    public async Task A_failing_notice_never_blocks_the_update()
    {
        using var f = new PendingFixture();
        var probe = new FakeProbe();
        var log = new List<string>();
        var watcher = new PendingUpdateWatcher(f.Paths, () => f.Engine(), probe, m => { lock (log) log.Add(m); },
            onInspection: _ => throw new InvalidOperationException("notice broken"));
        probe.Processes.Add(new PowerPointProcess(1, DateTime.UtcNow, true, "explorer.exe"));
        Assert.True(watcher.Trigger("poll")); await watcher.Current.WaitAsync(TimeSpan.FromSeconds(30));
        Assert.Contains(log, l => l.Contains("inspection not delivered (InvalidOperationException: notice broken)"));
        probe.Processes.Clear();
        Assert.True(watcher.Trigger("poll")); await watcher.Current.WaitAsync(TimeSpan.FromSeconds(30));
        Assert.Equal(PendingUpdateOutcome.Applied, watcher.LastResult!.Outcome);
        Assert.Equal(f.Pending, (string?)f.State()["installed"]);
    }

    [Fact]
    public async Task An_unverified_pending_update_is_never_announced()
    {
        using var f = new PendingFixture();
        File.AppendAllText(f.PendingFile, "tampered");
        var probe = new FakeProbe();
        probe.Processes.Add(new PowerPointProcess(1, DateTime.UtcNow, true, "explorer.exe"));
        int shows = 0;
        var notice = new UpdateNoticeController(() => Idle, _ => { shows++; return true; }, v => PendingNotice.MarkNotified(f.Paths, v), w => w(), _ => { });
        var watcher = new PendingUpdateWatcher(f.Paths, () => f.Engine(), probe, _ => { }, onInspection: notice.OnInspection);
        Assert.True(watcher.Trigger("poll")); await watcher.Current.WaitAsync(TimeSpan.FromSeconds(30));
        notice.Tick(); notice.Tick();
        Assert.Equal(0, shows);
        Assert.False(f.Engine().Inspect().Notified);
    }
}
