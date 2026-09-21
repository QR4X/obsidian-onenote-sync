# OneNote Vault Sync

High-speed incremental synchronization between Microsoft OneNote and Obsidian using native Windows OneNote COM automation.

## Overview

Unlike external binary converters or exporters, this solution interacts directly with the local Microsoft OneNote Desktop application through the official OneNote COM interface (`OneNote.Application`).

### Advantages
- **No Third-Party Binaries**: Works without standalone `.exe` tools, avoiding blocks by Windows Smart App Control, AppLocker, or Device Guard.
- **Fast Incremental Sync**: Evaluates page hierarchy and modification timestamps in under 1 second. Only new or modified pages are processed.
- **Accurate Markdown Conversion**: Converts outlines, multi-level lists, Markdown tables, and embedded images.
- **Dynamic Asset Paths**: Automatically calculates relative image links depending on folder nesting depth, ensuring images render properly in Obsidian.
- **Obsidian Integration**: Trigger synchronization directly via the Obsidian command palette (`Ctrl + P`).

## Architecture

The project consists of two components:
1. **PowerShell Engine (`scripts/sync-onenote.ps1`)**:
   - Communicates with OneNote via COM in Single-Threaded Apartment mode (`-STA`).
   - Maintains a local synchronization manifest (`.onenote-sync-manifest.json`).
   - Exports pages into `OneNote/<Notebook>/<Section>/<Page>.md`.
   - Extracts embedded base64 images into `OneNote/Assets/` and references them with correct relative Markdown links.
2. **Obsidian Plugin (`main.js`, `manifest.json`)**:
   - Registers the command `OneNote: Vault synchronisieren`.
   - Executes the sync script asynchronously in the background.
   - Shows status notices upon completion.

## Requirements
- Windows 10 or Windows 11
- Microsoft OneNote (Desktop version, 32-bit or 64-bit)
- Windows PowerShell 5.1 (standard built-in Windows component)
- Obsidian (desktop app)

## Installation

### Option 1: As an Obsidian Plugin
1. Create a folder named `obsidian-onenote-sync` inside your vault's `.obsidian/plugins/` directory.
2. Copy `manifest.json`, `main.js`, and the `scripts/` folder into that directory.
3. Reload Obsidian and enable **OneNote Vault Sync** under Community Plugins.
4. Press `Ctrl + P` and execute `OneNote: Vault synchronisieren`.

### Option 2: Standalone / CLI
Run the script from PowerShell or Command Prompt:

```powershell
& "scripts/sync-onenote.cmd"
```

To sync a specific notebook or section:

```powershell
& "scripts/sync-onenote.cmd" -Notebook "My Notebook" -Section "Work"
```

To force re-exporting all pages:

```powershell
& "scripts/sync-onenote.cmd" -Force
```

To list all pages and status without exporting:

```powershell
& "scripts/sync-onenote.cmd" -List
```

## License
MIT License
