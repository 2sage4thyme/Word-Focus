# Word Focus

**Highlight the word under your mouse, plus the next few words, in any Windows app.** A free reading-focus aid that helps your eyes stay on the line. It may help with ADHD, dyslexia, or plain reading fatigue.

Press a hotkey, move your mouse over text, and the word you're pointing at lights up with a soft color that fades across the next few words. Your mouse pointer turns almost invisible while it's over text, so it doesn't block what you're reading.

- Works across your desktop, not just in a browser: Claude, Chrome, Edge, Word, Notepad, PDFs opened in a browser, web textbook readers, and more
- Choose how many words to highlight: 1 to 10, or the rest of the line
- Ombre fade or one solid color
- Six preset colors plus a custom color picker
- Turn it on and off with **Ctrl+Alt+H**
- Nothing to install: it runs with PowerShell, which is built into Windows

## Download and run

1. Click the green **Code** button on this page, then **Download ZIP**.
2. Unzip it anywhere, for example into your Documents folder.
3. Double-click **Start Word Focus.cmd**.
   - If Windows shows "Windows protected your PC," click **More info**, then **Run anyway**. This appears because the file came from the internet and isn't signed by a company.
4. A small icon appears in the system tray (bottom-right of the taskbar; you may need to click the **^** arrow).
5. Press **Ctrl+Alt+H** and hover over some text.

To quit, right-click the tray icon and choose **Exit**.

## Controls

| Key | What it does |
|---|---|
| Ctrl+Alt+H | Turn the highlight on or off |
| Ctrl+Alt+K | Open settings (word count, ombre on/off, color) |
| Tray icon, right-click | On/off, settings, exit |

Settings reset to the defaults each time you start it (3 words, ombre on, yellow). To change the defaults, open `WordFocus.ps1` in Notepad and edit the values marked "Easy-to-change defaults".

## Where it works, and where it doesn't

Word Focus asks Windows' accessibility system (UI Automation, the same one screen readers use) which word is under the mouse. So it works in apps that share their text with that system:

- **Works:** Claude desktop app, Chrome, Edge, most websites, Word, Notepad, PDFs in Edge or Chrome, most web-based textbook readers
- **Usually doesn't:** games, text inside images, scanned PDFs, some e-reader apps, some music and creative software

## What it does to your computer

- It doesn't install anything, start with Windows, or change saved settings.
- It never connects to the internet and writes no files.
- While a highlight is showing, it swaps your arrow, text, and link pointers for see-through copies, kept in memory only. It puts your normal pointers back when you stop hovering, turn it off, or exit.
  - If Word Focus is force-closed (for example from Task Manager) while the pointer is faded, start it and exit it once, or sign out and back in, to get your normal pointer back.
- `Start Word Focus.cmd` runs PowerShell with `-ExecutionPolicy Bypass`, which lets this one script run without changing your computer's script settings.

You can read all of the code in `WordFocus.ps1` before running it.

## Requirements

- Windows 10 (tested). Windows 11 should work but hasn't been tested yet.

## How it was made

Word Focus was designed, directed, and tested by [2sage4thyme](https://github.com/2sage4thyme). The code was written by Claude, Anthropic's AI assistant, during that process.

## License

[Big Time Public License 2.0.2](LICENSE.md). In short:

- **Free** for personal use, students, schools, charities, nonprofits, government, and small businesses (fewer than 20 people and under about $1M in revenue).
- **Big businesses** need a paid license. To ask about one, email 2sage4thyme [at] gmail [dot] com or open an issue on this repository.

This is a short summary. The license text is what counts.

## Contact and feedback

2sage4thyme can be reached at **2sage4thyme [at] gmail [dot] com**.

Found an app where it doesn't work, or have an idea? Email, or [open an issue](../../issues).
