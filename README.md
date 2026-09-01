# pptNoob

A tiny PowerPoint add-in for macOS users who are stuck working with PowerPoint on a Mac and need a few missing features that the native app still does not offer.

This project is intentionally small and practical: it adds a few quick actions to the PowerPoint ribbon so everyday annoyances are easier to handle without having to fight the default UI.

## What it does

- Opens the containing folder for the current presentation in the browser
- Resets selected text margins to 0.1 cm
- Resets selected text margins to 0 cm

## Why this exists

PowerPoint on Mac can be surprisingly limited compared with the Windows experience. Sometimes the missing option is small, but it keeps showing up every time you work with a deck.

This add-in is just a gentle workaround for that reality: a few helpful shortcuts for people who are forced to use PowerPoint on Mac and just want the job done without too much friction.

## Run locally

```bash
npm install
npm start
```

Then sideload the add-in in PowerPoint on macOS using the generated manifest.

## Notes

This project is meant to stay lightweight and focused on small, useful fixes rather than trying to recreate a full PowerPoint productivity suite.
