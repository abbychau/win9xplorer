# Registry Settings Documentation

## Overview

The win9xplorer application now uses Windows Registry to persist all user settings, making them stateful across application sessions. This document explains how the registry-based settings work.

## Registry Location

All settings are stored under:
```
HKEY_CURRENT_USER\SOFTWARE\win9xplorer
```

## Settings Categories

### 1. Window Settings (`Window` subkey)
- **Width, Height**: Form dimensions
- **Left, Top**: Form position on screen
- **WindowState**: Normal, Minimized, or Maximized
- **SplitterDistance**: Position of the tree view/list view splitter
- **TreeViewVisible**: Whether the folder tree is visible
- **LastPath**: Last visited directory/location
- **ListViewMode**: Current view mode (Large Icons, Small Icons, List, Details)

### 2. Toolbar Settings (`ToolStrip` subkey)
- **ToolStripX, ToolStripY**: Position of the main toolbar
- **AddressStripX, AddressStripY**: Position of the address bar
- **AddressTextBoxWidth, AddressTextBoxHeight**: Size of the address text box

### 3. Font Settings (`Fonts` subkey)
- **TreeViewFontName, TreeViewFontSize, TreeViewFontStyle**: TreeView font configuration
- **ListViewFontName, ListViewFontSize, ListViewFontStyle**: ListView font configuration

### 4. Bookmarks/Favorites (`Bookmarks` subkey)
Each bookmark is stored as a numbered subkey (`Bookmark_0`, `Bookmark_1`, etc.):
- **Name**: Display name of the bookmark
- **Path**: File system path or special location ("My Computer")
- **DateAdded**: When the bookmark was created (as binary DateTime)

## Automatic Persistence

Settings are automatically saved when:
- **Window changes**: Form resize, move, or state change (with 500ms delay to avoid excessive saves)
- **View mode changes**: Switching between Large Icons, Small Icons, List, or Details view
- **Splitter moves**: Adjusting the tree view/list view splitter (with 300ms delay)
- **Toolbar moves**: Repositioning toolbars (with 500ms delay)
- **Tree view toggle**: Showing/hiding the folder tree
- **Font changes**: Applying new fonts through Options dialog
- **Navigation**: Changing directories or locations
- **Bookmarks**: Adding or removing favorites
- **Application exit**: All settings saved on close

## Settings Management Features

### Options Dialog
- Access via `Tools` ¡÷ `Options...`
- Configure TreeView and ListView fonts
- Changes are immediately saved to registry

### Export/Import Settings
- **Export**: `Tools` ¡÷ `Export Settings...` - Saves settings to a `.reg` file
- **Import**: `Tools` ¡÷ `Import Settings...` - Loads settings from a `.reg` file
- Application restarts automatically after import to apply settings

### Reset All Settings
- Access via `Tools` ¡÷ `Reset All Settings...`
- Clears all registry entries and restores defaults
- Application restarts automatically after reset

## Technical Implementation

### RegistrySettingsManager Class
The `RegistrySettingsManager` class handles all registry operations:
- **Load methods**: Read settings from registry with fallback defaults
- **Save methods**: Write current settings to registry
- **Apply methods**: Apply loaded settings to form controls
- **Export/Import**: Registry file operations
- **Clear**: Remove all application settings

### Automatic Save Mechanisms
- **Timers**: Used to batch saves and avoid excessive registry writes during dragging operations
- **Event handlers**: Attached to form and control events for automatic persistence
- **Validation**: Settings are validated before application (screen bounds, minimum sizes, etc.)

### Error Handling
- All registry operations are wrapped in try-catch blocks
- Failed operations are logged to Debug output
- Application continues with default settings if registry operations fail

## Settings Backup and Restoration

### Manual Backup
1. Use `Tools` ¡÷ `Export Settings...` to create a backup `.reg` file
2. Store the file in a safe location

### Manual Restoration
1. Use `Tools` ¡÷ `Import Settings...` to restore from a backup `.reg` file
2. Or double-click the `.reg` file in Windows Explorer (requires confirmation)

### Registry Editor Access
Advanced users can directly edit settings in Registry Editor:
1. Open `regedit.exe`
2. Navigate to `HKEY_CURRENT_USER\SOFTWARE\win9xplorer`
3. Edit values as needed (application must be closed first)

## Default Values

When no registry settings exist (first run), the application uses these defaults:
- **Window size**: 871¡Ñ384 pixels
- **Window position**: 100,100 from top-left corner
- **Window state**: Normal (not maximized/minimized)
- **Splitter distance**: 217 pixels
- **Tree view**: Visible
- **List view mode**: Details
- **Fonts**: System default font for both TreeView and ListView
- **Last path**: "My Computer"
- **Bookmarks**: None

## Troubleshooting

### Settings Not Persisting
1. Check if application has registry write permissions
2. Verify Windows Registry service is running
3. Try running as Administrator (one time) to establish registry keys
4. Use "Reset All Settings" to clear corrupted entries

### Application Won't Start
1. Delete the registry key: `HKEY_CURRENT_USER\SOFTWARE\win9xplorer`
2. Restart the application (will use defaults)
3. Or restore from a known good `.reg` backup file

### Corrupted Settings
1. Use `Tools` ¡÷ `Reset All Settings...` from within the application
2. Or manually delete registry entries and restart
3. Restore from backup if available

## Security Notes

- Settings are stored per-user (HKEY_CURRENT_USER) - not system-wide
- No sensitive information is stored in the registry
- Settings can be cleared at any time without affecting system operation
- Registry access follows Windows security model (user permissions only)