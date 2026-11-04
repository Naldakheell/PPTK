using System.Runtime.InteropServices;
using Ppt = Microsoft.Office.Interop.PowerPoint;

namespace PPTKKeys;

internal sealed class Snapshot
{
    public static readonly Snapshot Unknown = new(Surface.Unknown, "unknown");

    public Surface Surface { get; }
    public string Detail { get; }
    public int Pane { get; }
    public int Selection { get; }

    public Snapshot(Surface surface, string detail, int pane = 0, int selection = 0)
    {
        Surface = surface;
        Detail = detail;
        Pane = pane;
        Selection = selection;
    }
}

/// <summary>
/// PowerPoint COM on the helper STA. The keyboard hook never calls this.
/// Events and posted refreshes update one snapshot. Reads that fail publish Unknown.
/// </summary>
internal sealed class PowerPointSession : IDisposable
{
    static readonly Guid PowerPointClsid = new("91493441-5A91-11CF-8700-00AA0060263B");

    Ppt.Application? _app;
    HoldState _hold = HoldState.Off;
    bool _insideT;
    string _lastError = "";

    public Snapshot Current { get; private set; } = Snapshot.Unknown;

    public bool IsAttached => _app != null;

    public bool TextHold => _hold.Held;

    public Surface EffectiveSurface =>
        ShortcutGate.EffectiveSurface(Current.Surface, _hold.Held, Current.Surface == Surface.Blocked && IsSlideshow(Current));

    public bool TryAttach()
    {
        if (_app != null) return true;
        try
        {
            Guid clsid = PowerPointClsid;
            NativeMethods.GetActiveObject(ref clsid, IntPtr.Zero, out object unknown);
            var app = (Ppt.Application)unknown;
            app.WindowSelectionChange += OnWindowSelectionChange;
            app.SlideShowBegin += OnSlideShowBegin;
            app.SlideShowEnd += OnSlideShowEnd;
            app.WindowActivate += OnWindowActivate;
            _app = app;
            Console.WriteLine("PowerPoint: attached");
            return true;
        }
        catch (COMException)
        {
            return false;
        }
        catch (InvalidCastException ex)
        {
            NoteError("PowerPoint COM cast failed: " + ex.Message);
            return false;
        }
    }

    public void Refresh()
    {
        if (_app == null && !TryAttach())
        {
            PublishObserved(new Snapshot(Surface.Unknown, "PowerPoint not running"), runFinished: false, selectionEvent: false);
            return;
        }

        PublishObserved(Read(), runFinished: false, selectionEvent: false);
    }

    public void ArmTextHold()
    {
        _hold = HoldState.Armed;
        Current = new Snapshot(Surface.Text, "hold after T");
    }

    public void FinishTextInsert()
    {
        _insideT = true;
        try
        {
            Run(ShortcutGate.MacroName, "T");
        }
        finally
        {
            _insideT = false;
        }

        PublishObserved(Read(), runFinished: true, selectionEvent: false);
    }

    static int _iconBusy;

    public void Dispatch(string token)
    {
        if (token == "T" || !ShortcutGate.IsRunnable(token)) return;
        if (token == "I")
        {
            BeginInsertIcon();
            return;
        }

        Run(ShortcutGate.MacroName, token);
    }

    /// <summary>
    /// Insert Icon is a modal PowerPoint dialog. Run it on a separate STA so this
    /// helper thread can keep pumping the keyboard hook, and ignore I while it is open.
    /// </summary>
    static void BeginInsertIcon()
    {
        if (Interlocked.CompareExchange(ref _iconBusy, 1, 0) != 0)
        {
            Console.Error.WriteLine("Icon dialog already open");
            return;
        }

        var thread = new Thread(InsertIconWorker)
        {
            IsBackground = true,
            Name = "PPTKKeys.Icon",
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    static void InsertIconWorker()
    {
        object? unknown = null;
        try
        {
            Guid clsid = PowerPointClsid;
            NativeMethods.GetActiveObject(ref clsid, IntPtr.Zero, out unknown);
            var app = (Ppt.Application)unknown;
            object[] args = ["I"];
            app.Run(ShortcutGate.MacroName, ref args);
            Console.WriteLine("Dispatched IconInsertFromFile");
        }
        catch (Exception ex) when (ex is COMException or InvalidCastException or InvalidComObjectException or System.Reflection.TargetException)
        {
            Console.Error.WriteLine("Icon dialog failed: " + ex.GetType().Name);
        }
        finally
        {
            if (unknown != null) Marshal.ReleaseComObject(unknown);
            Interlocked.Exchange(ref _iconBusy, 0);
        }
    }

    public void CancelHold()
    {
        _hold = HoldState.Off;
        Current = Snapshot.Unknown;
    }

    public Snapshot Peek()
    {
        if (_app == null && !TryAttach()) return new Snapshot(Surface.Unknown, "PowerPoint not running");
        return Read();
    }

    public void Dispose() => Drop();

    void OnWindowSelectionChange(Ppt.Selection selection)
    {
        if (_insideT || _app == null) return;
        try
        {
            int selected = (int)selection.Type;
            int shows = _app.SlideShowWindows.Count;
            if (shows > 0)
            {
                PublishObserved(new Snapshot(Surface.Blocked, "slideshow"), runFinished: false, selectionEvent: true);
                return;
            }

            int pane = (int)_app.ActiveWindow.ActivePane.ViewType;
            Surface surface = ShortcutGate.Classify(false, pane, selected, true);
            PublishObserved(new Snapshot(surface, Detail(pane, selected), pane, selected), runFinished: false, selectionEvent: true);
        }
        catch (Exception ex) when (ex is COMException or InvalidComObjectException or System.Reflection.TargetException)
        {
            if (IsDisconnected(ex)) Drop();
            PublishObserved(new Snapshot(Surface.Unknown, "selection event failed"), runFinished: false, selectionEvent: true);
        }
    }

    void OnSlideShowBegin(Ppt.SlideShowWindow window) =>
        PublishObserved(new Snapshot(Surface.Blocked, "slideshow"), runFinished: false, selectionEvent: false);

    void OnSlideShowEnd(Ppt.Presentation presentation) => Refresh();

    void OnWindowActivate(Ppt.Presentation presentation, Ppt.DocumentWindow window) => Refresh();

    void Run(string macro, string token)
    {
        if (!ShortcutGate.IsRunnable(token) || macro != ShortcutGate.MacroName)
            throw new InvalidOperationException("Refusing an unlisted macro call.");
        if (_app == null && !TryAttach())
        {
            Console.Error.WriteLine($"Dispatch failed {token}: PowerPoint is not running");
            return;
        }

        try
        {
            object[] args = [token];
            _app!.Run(macro, ref args);
            Console.WriteLine($"Dispatched {macro} {token}");
        }
        catch (Exception ex) when (ex is COMException or InvalidComObjectException or System.Reflection.TargetException)
        {
            if (IsDisconnected(ex)) Drop();
            Console.Error.WriteLine($"Dispatch failed {token}: {ex.GetType().Name}");
        }
    }

    Snapshot Read()
    {
        if (_app == null) return new Snapshot(Surface.Unknown, "PowerPoint not running");
        try
        {
            int shows = _app.SlideShowWindows.Count;
            if (shows > 0) return new Snapshot(Surface.Blocked, "slideshow");
            var window = _app.ActiveWindow;
            int pane = (int)window.ActivePane.ViewType;
            int selected = (int)window.Selection.Type;
            Surface surface = ShortcutGate.Classify(false, pane, selected, true);
            return new Snapshot(surface, Detail(pane, selected), pane, selected);
        }
        catch (Exception ex) when (ex is COMException or InvalidComObjectException or System.Reflection.TargetException)
        {
            if (IsDisconnected(ex)) Drop();
            return new Snapshot(Surface.Unknown, "read failed");
        }
    }

    public void Invalidate()
    {
        if (_hold.Held) return;
        Current = new Snapshot(Surface.Unknown, "focus changed");
    }

    void PublishObserved(Snapshot observed, bool runFinished, bool selectionEvent)
    {
        bool slideshow = IsSlideshow(observed);
        HoldResult result = ShortcutGate.Reduce(
            _hold,
            observed.Surface,
            observed.Detail,
            slideshow,
            runFinished,
            selectionEvent);
        _hold = result.Hold;
        Current = new Snapshot(result.Surface, result.Detail, observed.Pane, observed.Selection);
    }

    static string Detail(int pane, int selection) =>
        $"pane={pane}({ShortcutGate.PaneName(pane)}) selection={selection}({ShortcutGate.SelectionName(selection)})";

    static bool IsSlideshow(Snapshot snapshot) =>
        snapshot.Surface == Surface.Blocked && snapshot.Detail.Contains("slideshow", StringComparison.Ordinal);

    void Drop()
    {
        var app = _app;
        _app = null;
        _hold = HoldState.Off;
        Current = Snapshot.Unknown;
        if (app == null) return;
        try
        {
            app.WindowSelectionChange -= OnWindowSelectionChange;
            app.SlideShowBegin -= OnSlideShowBegin;
            app.SlideShowEnd -= OnSlideShowEnd;
            app.WindowActivate -= OnWindowActivate;
        }
        catch (Exception ex) when (ex is COMException or InvalidComObjectException)
        {
        }

        try
        {
            Marshal.ReleaseComObject(app);
        }
        catch (Exception ex) when (ex is COMException or InvalidComObjectException)
        {
        }

        Console.WriteLine("PowerPoint: detached");
    }

    void NoteError(string message)
    {
        if (message == _lastError) return;
        _lastError = message;
        Console.Error.WriteLine(message);
    }

    static bool IsDisconnected(Exception ex)
    {
        if (ex is COMException com)
        {
            int code = com.HResult;
            return code is unchecked((int)0x800706BA) or unchecked((int)0x80010108) or unchecked((int)0x800706BE);
        }

        return ex is InvalidComObjectException;
    }
}
