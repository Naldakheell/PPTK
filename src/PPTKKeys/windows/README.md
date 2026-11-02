# PPTKKeys for Windows

Development helper for the nine bare-letter PPTK inserts. It calls the existing VBA entry point only:

`Application.Run("PPTK_DispatchShortcut", token)`

`token` is one of `T R U O C P G M N`. The macOS helper is not involved.

This project was compiled on macOS. The keyboard hook, PowerPoint COM, and the manual matrix below have not been executed on Windows.

## Requirements

- Windows 10 or later
- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0) (building and `dotnet run` also need the .NET 10 desktop runtime, which the SDK includes)
- Desktop PowerPoint, already running, with `PPTK_DispatchShortcut` available to `Application.Run`
- No administrator account and no registry changes

## Build

From this directory:

```bat
build.cmd
```

Or:

```bat
dotnet restore PPTKKeys.csproj
dotnet build PPTKKeys.csproj -c Release
```

The decision matrix alone can be tested on macOS or Windows:

```bat
dotnet test PPTKKeys.Tests\PPTKKeys.Tests.csproj
```

## Run

Start PowerPoint and open the PPTK presentation first, then:

```bat
run.cmd
```

Or:

```bat
dotnet run --project PPTKKeys.csproj -c Release
```

Ctrl+C stops the helper. A second copy exits immediately.

Print the current gate without installing the hook:

```bat
dotnet run --project PPTKKeys.csproj -c Release -- --describe
```

`--describe` prints the pane id, selection id, focus window classes, and whether each token would dispatch. Use it on the canvas, in a text box, on the thumbnail pane, in Notes, and in the ribbon before trusting the hook.

## What the hook does

`WH_KEYBOARD_LL` handles key-down for the nine letters. Ctrl, Shift, Alt, and the Windows key are ignored. The hook does not call PowerPoint. It reads:

- foreground process name `POWERPNT.EXE`
- foreground class `PPTFrameClass` (a dialog or slideshow window is not the frame)
- focused window class (`NetUIHWND`, `Edit`, RichEdit, `#32770`, command bars, and menu mode fail open)
- a cached COM snapshot

The snapshot is updated from `WindowSelectionChange`, `SlideShowBegin`, `SlideShowEnd`, and `WindowActivate`, plus a refresh after the foreground or PowerPoint focus window changes. A failed read publishes an unknown state, and an unknown state does not swallow the key.

`T` arms a text hold before the hook returns, so the next bare letter cannot dispatch a shape. Letters typed before the text box exists are queued and replayed with Unicode input after `PPTK_DispatchShortcut "T"` returns. If that call does not leave a text selection, bare letters stay disabled until a later selection event reports the canvas or the thumbnail pane.

## Manual tests

Run these on Windows with the helper in a console and PowerPoint in front. Each insert action should happen once, and the original letter should not appear, except where the expected result is normal typing.

A. Slide canvas. Press `T R U O C P G M N`. Each PPTK action runs once. The letter is not typed.

B. An existing text box, caret inside it. Type `T R U O C P G M N`. Every letter appears. No insert runs.

C. On the canvas, press `T`, then immediately type `R U O C P G M N`. Those letters go into the new text box. No shapes are inserted.

D. Thumbnail pane. `N` should add a slide when `--describe` reports pane `thumbnails` (11) or `sorter` (7), or selection `slides` (1) on the slide pane. `R` must not insert a rectangle. Selection `text` is 3, not slides. If the thumbnail pane reports as `outline` (6), `N` stays canvas-only.

E. Notes pane. None of the nine letters insert a shape.

F. Ribbon font box, Tell Me, and a modal dialog. None of the nine letters insert a shape.

G. Another application. All letters type normally.

H. Slide show. None of the nine letters insert a shape.
