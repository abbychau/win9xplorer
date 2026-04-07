# Win9xplorer

A faithful recreation of the classic Windows 9x experience for modern Windows, featuring a retro file explorer and taskbar with authentic aesthetics and behavior, that AI knows very little.

## Screenshots

Single Bevel Taskbar

![](./screenshots/1.png)


Tray icon and menus

![](./screenshots/2.png)


Customized Menu Style

![](./screenshots/3.png)

Options

![](./screenshots/4.png)

File Explorer

![](./screenshots/5.png)

Windows Apps Menu for newer apps from Windows Store

![](./screenshots/6.png)

Classic Context Menus in Start Menu

![](./screenshots/7.png)

Double Bevel (Classic Style)

![](./screenshots/8.png)


## Design Philosophy

**Authenticity** – Win9xplorer aims to capture not just the visual appearance of Windows 95/98, but the *feel* of using it. Every bevel, shadow, and interaction pattern is crafted to evoke the genuine retro computing experience.

**Keystrike Matters** – Be sure that high frequency keystrikes(e.g. cursor navigations, searches) that happens in old and new start menus still work.

## Features

- **Classic File Explorer** – Two-pane folder browser with TreeView and ListView, supporting drag-and-drop file operations, context menus, and address bar navigation
- **Retro Taskbar** – Authentic Win9x-style taskbar with Start menu, Quick Launch, running programs, and system tray
- **Windows XP Icons** – Uses authentic Windows XP icon set with fallback to shell32.dll
- **Custom Rendering** – Win9xMenuRenderer for genuine 3D bevels, colors, and menu styling
- **Bookmarks & History** – Navigation history with back/forward, favorites sidebar
- **Persistent Settings** – Registry-based configuration saves window layout, toolbar positions, and preferences

## Usage

### Running the Application

```bash
dotnet run --project .\win9xplorer.csproj
```

### Building

```bash
dotnet build
```

The output executable will be in `bin\Debug\net10.0-windows\`.

## Requirements

- Windows 11 or Windows 10
- .NET 10.0 Windows SDK
- Visual Studio 2022 or later (for development)

See [PROJECT_STRUCTURE.md](PROJECT_STRUCTURE.md) for detailed architecture documentation.

## License

GNU GENERAL PUBLIC LICENSE

## Inspired by

[Retrobar](https://github.com/dremin/RetroBar)

[ManagedShell](https://github.com/cairoshell/ManagedShell)