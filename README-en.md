### Personal Translation Tool
A tool created for my personal translation and localization work. This tool operates based on context and is powered by artificial intelligence. """
## TransPal — Translation Assistant
========================================================================================
### Installation (once, open CMD):
pip install keyboard pyperclip pyautogui google-genai deep-translator python-docx pypdf
Set API key (once initially):
setx GEMINI_API_KEY "your_key"
(close and reopen the application after setting)
Run: double-click the transpal_gui.pyw file, no terminal needed.
"""
### ===== Instructions =========================================================================
## Translate text:

Open any application (browser, Word, chat software...), highlight the text to be translated
Press Alt+Shift+T and release
At this point, there are 3 options in the pop-up window:
Edit the translation directly in the box if it doesn't match, then press
"Replace Press "like" (or Enter) → the highlighted text will be pasted with the translation
"Copy only" → the translation is in the clipboard, it will paste automatically
"Close" (or Esc) → skip, do nothing

##Other shortcuts:

Ctrl+Alt+M: quickly switch between AI mode (slower but understands context) and basic Google Translate (fast, no key needed)
Ctrl+Alt+X: close the script

Regarding the context file — this is the most important part for the AI ​​to translate correctly
When run for the first time, the script automatically creates a context.md file in the same directory as transpal.py. Open it with a modified Notepad (or VS Code), for example you can write:
markdown# Algorithm instructions
- Natural translation, Vietnamese style, using "you".

- This is documentation about an MMORPG game, containing words: buff, debuff, raid, dungeon.
- "Guild" translates to "guild," and "quest" translates to "mission."

- Character names are unique; there are no audio versions.

- The writing style is youthful and not overly formal.

Some points to note:

Editing context.md doesn't require restarting the script — the file can be read again each time you press the translate key, so after editing, saving (Ctrl+S) allows you to use it immediately in the next translation.

You can write anything: translation rules, glossaries, document context descriptions ("this is a legal contract," "this is a fantasy novel")... AI will read and follow.

If you have a pre-written .txt file, copy the style content into context.md or import it. If it's a PDF/Word document, you currently need to copy the text manually.

Note that AI mode only applies context; basic translation mode (Google Translate) cannot read this file because it's a static machine translation.

*For AI models, code adjustments can be made to regain maximum speed and accuracy; this tool is using free, speed-priority models with high RPM/RPD.*
