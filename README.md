# Workdesk Overlay

A minimalist workday dashboard overlay for Windows built with PowerShell + WPF.

It provides a centered overlay window with:
- A small monthly calendar
- Quick Notes panel (auto-saved notes bound to calendar days)
- Focus / Pomodoro timer
- System status (CPU, RAM, battery, network)
- Mini to‑do list
- Unread Outlook mails preview
- Quick launch shortcuts (Outlook, Teams, Explorer, Edge, VS Code, Excel)

Toggled globally with **Ctrl + Alt + D**, so you can open/close it on top of your current work.

## Requirements
- Windows (WPF desktop, .NET assemblies loaded via `PresentationFramework` etc.)
- PowerShell (Windows PowerShell recommended)
- Outlook desktop client for the "Unread mails" panel (uses Outlook COM; may be blocked by company policy)

## Run
```powershell
powershell -STA -NoProfile -ExecutionPolicy Bypass -File ".\command-center.ps1"
```

The script starts a hidden dashboard window and registers a global hotkey for toggling.

## Hotkey
- Toggle overlay: **Ctrl + Alt + D**

## Notes
- Notes are auto-saved into `notes.json`.
- Notes are associated with calendar days; you can:
	- Use **Prev/Next** arrows to change month
	- Click a day to see that day's notes
	- Use **+** to create a new note for today
	- Use the calendar-note button to create a note for the selected day

`notes.json` may contain personal data and is **never** committed to Git (see `.gitignore`).
When cloning the repository:
- Copy `notes.example.json` to `notes.json` to start with an empty, safe file.

## Todos
- Tasks are stored in `todos.json` and shown in the **TODAY'S TASKS** card.
- Click **+** to open the input, type a task and press **Enter** to save.
- Press **Escape** in the task input to cancel/close the input.
- Click a task text to edit it; clearing the text and confirming removes the task.
- Check/uncheck the box to mark tasks done/undone (with strikethrough style).

`todos.json` also stays local and is ignored by Git.
When cloning:
- Copy `todos.example.json` to `todos.json`.

## Pomodoro / Focus
- Default focus session: **25 minutes**.
- Break session: **5 minutes**.
- Buttons:
	- **Start Focus** / **Pause / Resume**: start or pause a focus sprint.
	- **Break**: start a 5‑minute break.
	- **Reset**: go back to idle and reset the timer.
- A progress bar and label show remaining time and current mode.

## System Status
- Shows CPU usage (%), memory usage (used/total GB and %), and a small meta line:
	- Battery level (if available)
	- Network online/offline state
- Values refresh automatically every few seconds.

## Unread Mails
- Shows up to **5 unread** emails from Outlook.
- Each row displays sender, subject and received time.
- Click a mail row to open it in Outlook.
- If Outlook is not installed or COM is blocked, a message is shown instead.

## Quick Launch
- Shortcuts panel to quickly start common work apps:
	- Outlook
	- Microsoft Teams (tries multiple launch methods; falls back to web)
	- File Explorer
	- Microsoft Edge
	- VS Code
	- Excel

## Data & Privacy
- The following local data files are **ignored by Git** and never pushed to GitHub:
	- `notes.json`
	- `todos.json`
	- `cc-errors.log`
- Example files committed to the repo:
	- `notes.example.json`
	- `todos.example.json`
- When setting up on a new machine, copy the example files to their non‑example counterparts before running the script.
