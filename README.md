# Word Focus

**Highlight the word under your mouse, plus the next few words, in any Windows app.** A free reading-focus aid that helps your eyes stay on the line. It may help with ADHD, dyslexia, or plain reading fatigue.

Press a hotkey, move your mouse over text, and the word you're pointing at lights up with a soft color that fades across the next few words. Your mouse pointer fades while it's over text (you choose how much), so it doesn't block what you're reading.

- Works across your desktop, not just in a browser: Claude, Chrome, Edge, Word, Notepad, PDFs opened in a browser, web textbook readers, and more
- Choose how many words to highlight: 1 to 10, to the end of the sentence, to the end of the line, or the end of the line plus the first word of the next line
- Highlight, underline, or both. In "both", the highlight uses your color and the underline automatically turns white on dark pages and black on light ones
- Ombre fade or one solid color, with or without thin overlap lines where words meet
- Adjustable highlight strength, from faint to solid
- Six preset colors plus a custom color picker
- Turn it on and off with **Ctrl+Alt+H**
- **Paragraph Focus** and **Page Focus**: point at text and press a hotkey to softly light up the whole paragraph, or all the body text visible in the window (skipping headings, buttons, and side menus, and following along as you scroll), while the hover highlight keeps working on top
- **Word Focus Reader**: inside a Paragraph / Page Focus area, Ctrl+click a word and a few words at a time light up and move along from there. Ctrl+click again to pause; the next Ctrl+click starts from wherever you click at the speed you choose, pausing briefly at commas, periods, and other punctuation (adjustable), and twice as long at paragraph breaks. The Reader has its own style (highlight, underline, or both) and color, set in Advanced. Not sure what speed to pick? **Test my speed...** in Advanced explains the test, lets you pick a short informational passage (cheese history, the quadratic formula, or how ecosystems work), and has you read it while following along with the mouse, then suggests a speed The Play/Pause key on your keyboard pauses and resumes
- **Read aloud**: the Reader can speak the words as it highlights them, using Windows' built-in voices. With read aloud on, the voice sets the pace and the highlight follows it
- A hotkey for anything Word Focus can do (word count, style, colors, strength, pointer, Focus layers, and more), all optional and all changeable
- Optional **Start with Windows** box, so it's always ready
- Nothing to install: it runs with PowerShell, which is built into Windows

## Download and run

1. Click the green **Code** button on this page, then **Download ZIP**.
2. Unzip it anywhere, for example into your Documents folder.
3. Double-click **Start Word Focus.cmd**.
   - If Windows shows "Windows protected your PC," click **More info**, then **Run anyway**. This appears because the file came from the internet and isn't signed by a company.
4. A small icon appears in the system tray (bottom-right of the taskbar; you may need to click the **^** arrow).
5. Press **Ctrl+Alt+H** and hover over some text.

Word Focus keeps running in the tray until you exit or restart. To have it start automatically, press **Ctrl+Alt+K** and tick **Start with Windows**. It comes back with your last settings, including whether the highlight was on.

To quit, right-click the tray icon and choose **Exit**.

## Controls

| Key | What it does |
|---|---|
| Ctrl+Alt+H | Turn the highlight on or off |
| Ctrl+Alt+K | Open settings (word count, ombre, overlap lines, highlight strength, pointer visibility, color, start with Windows) |
| Ctrl+Alt+A | Paragraph Focus: light up the paragraph under the mouse (press again to turn off) |
| Ctrl+Alt+B | Page Focus: light up the body text visible in the window (press again to turn off) |
| Ctrl + click (inside a Focus area) | Reader: 1st click starts at that word, 2nd click pauses, 3rd click starts at the new word, and so on |
| Play/Pause key | Reader: pause / resume (only while the Reader is running) |
| Tray icon, right-click | On/off, settings, exit |

To change the style (highlight, underline, or both) or the Paragraph / Page Focus layer, open settings and click **Advanced...**. To set hotkeys, click **Hotkeys...** in Advanced: every action is listed, and anything without a hotkey shows "none" until you give it one. If another program already uses one of your hotkeys, Word Focus tells you when it starts.

Your settings, including whether the highlight was on, are remembered between sessions and restarts. They're saved in `%APPDATA%\WordFocus\settings.txt`. To go back to the defaults (3 words, ombre on, overlap lines on, 60% strength, pointer 10% visible, yellow), exit Word Focus and delete that file.

## Where it works, and where it doesn't

Word Focus asks Windows' accessibility system (UI Automation, the same one screen readers use) which word is under the mouse. So it works in apps that share their text with that system:

- **Works:** Claude desktop app, Chrome, Edge, most websites, Word, Notepad, PDFs in Edge or Chrome, most web-based textbook readers
- **Usually doesn't:** games, text inside images, scanned PDFs, some e-reader apps, some music and creative software

## What it does to your computer

- It doesn't install anything or change saved settings.
- It only starts with Windows if you tick **Start with Windows**. That adds a shortcut named "Word Focus" to your personal Startup folder, and unticking removes it. If you move the Word Focus folder later, untick and re-tick the box so the shortcut points to the new location.
- While a Paragraph / Page Focus area is showing, it watches mouse clicks, only to notice Ctrl+click inside that area for the Reader. That click is kept from reaching the app, so links don't open. While the Reader runs, it borrows your keyboard's Play/Pause key, and gives it back when the Reader stops. Nothing is recorded.
- **Read aloud** uses the voices built into Windows. Nothing is downloaded or sent anywhere.
- It never connects to the internet. The only files it writes are your settings file (`%APPDATA%\WordFocus\settings.txt`) and that optional shortcut.
- While a highlight is showing, it swaps your arrow, text, and link pointers for see-through copies, kept in memory only. It puts your normal pointers back when you stop hovering, turn it off, or exit.
  - If Word Focus is force-closed (for example from Task Manager) while the pointer is faded, start it and exit it once, or sign out and back in, to get your normal pointer back.
- `Start Word Focus.cmd` runs PowerShell with `-ExecutionPolicy Bypass`, which lets this one script run without changing your computer's script settings.

You can read all of the code in `WordFocus.ps1` before running it.

## Requirements

- Windows 10 or Windows 11 (both tested).

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

## Speed test passages: sources

The reading speed test's passages were written for Word Focus. They draw on:

- [Chemical analysis reveals first cheese-making in Northern Europe 7,000 years ago](https://www.sciencedaily.com/releases/2012/12/121212134044.htm) (ScienceDaily)
- [Pots bear oldest signs of cheese making](https://www.sciencenews.org/article/pots-bear-oldest-signs-cheese-making) (Science News)
- [Cheese found in an Egyptian tomb is at least 3,200 years old](https://www.sciencenews.org/article/cheese-ancient-egyptian-tomb-3200-years-old) (Science News)
- [Jesse Williams' early cheese factory in Rome, NY](https://www.newyorkalmanack.com/2019/09/jesse-williams-early-cheese-factory-in-rome-ny/) (New York Almanack)
- [Babylonian mathematics](https://en.wikipedia.org/wiki/Babylonian_mathematics) (Wikipedia)
- [Producers](https://education.nationalgeographic.org/resource/producers/) (National Geographic Education)
- [Consumers](https://education.nationalgeographic.org/resource/consumers/) (National Geographic Education)
- [Decomposers](https://education.nationalgeographic.org/resource/decomposers/) (National Geographic Education)
- [Food chain](https://education.nationalgeographic.org/resource/food-chain/) (National Geographic Education)
