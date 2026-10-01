# CLAUDE.md

## Who does what

Project agents live in `.claude/agents/`. Route work like this:

| Agent | Model | Use it for |
|---|---|---|
| `quick-helper` | Haiku 4.5 | Reading the diary, finding things in files. Read-only. |
| `builder` | Sonnet 5.5 | Normal code changes: features, fixes, small refactors. |
| `architect` | Opus 5.5 | Big designs, hard bugs, and the Revit add-in. |

Rule of thumb: look-ups go to `quick-helper`, everyday edits go to `builder`, and anything big, tricky, or Revit-related goes to `architect`.
