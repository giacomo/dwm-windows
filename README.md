# 🪟 DWM Windows

> A modern Node.js/TypeScript library for accessing Windows Desktop Window Manager (DWM) APIs. Get live window thumbnails, icons, and manage virtual desktops with native C++ performance.

## ✨ Features

- 🖼️ **Live window thumbnails** - Real-time PNG previews of any window
- 🎨 **Window icons** - Extract application icons as base64 PNG data
- 🖥️ **Virtual desktop support** - Access windows across all virtual desktops
- 📝 **TypeScript-first** - Full type definitions and IntelliSense support
- ⚡ **Native performance** - C++ bindings for optimal speed
- 🎯 **Window management** - Focus, filter, and control windows programmatically
- 🔔 **Event hooks (no polling)** - Created/closed/focused callbacks via native WinEvent hooks

## 📦 Installation

```bash
yarn add dwm-windows
```

## 🚀 Quick Start

```typescript
import dwmWindows from 'dwm-windows';

// Get all visible windows on current desktop
const windows = dwmWindows.getVisibleWindows();

for (const window of windows) {
  console.log(`${window.title} (${window.executablePath})`);
  console.log(`Icon: ${window.icon.substring(0, 50)}...`);
  console.log(`Thumbnail: ${window.thumbnail.substring(0, 50)}...`);
}

// Focus a specific window
dwmWindows.openWindow(windows[0].id);
```

## 📖 API Reference

### Core Methods

#### `getVisibleWindows(options?): WindowInfo[]`
Returns visible windows. Use `{ includeAllDesktops: true }` to include all virtual desktops.

#### `getWindows(options?): WindowInfo[]`
Returns all windows (including hidden ones).

#### `openWindow(windowId: number): boolean`
Brings a window to the foreground and focuses it.

#### `updateThumbnail(windowId: number): string`
Refreshes and returns the PNG thumbnail as a base64 data URL.

### Async API (non-blocking)

Promise-based variants that run work off the main thread and won't block the event loop:

- `getWindowsAsync(options?): Promise<WindowInfo[]>`
- `getVisibleWindowsAsync(options?): Promise<WindowInfo[]>`
- `openWindowAsync(windowId: number): Promise<boolean>`
- `updateThumbnailAsync(windowId: number): Promise<string>`

Example:

```ts
import dwmWindows from 'dwm-windows';

const all = await dwmWindows.getWindowsAsync({ includeAllDesktops: true });
if (all.length) {
  await dwmWindows.openWindowAsync(all[0].id);
  const freshThumb = await dwmWindows.updateThumbnailAsync(all[0].id);
  console.log('Thumb length:', freshThumb.length);
}
```

### Event Hooks (no polling)

Subscribe to native window events without polling:

- `onWindowCreated(cb: (e) => void)`
- `onWindowClosed(cb: (e) => void)`
- `onWindowFocused(cb: (e) => void)`
- `onWindowChange(cb: (e) => void)` // unified: e.type in {created,closed,focused,minimized,restored}
- `stopWindowEvents()`

Each event payload contains: `{ id, hwnd, title, executablePath, isVisible, type }`.
For `closed`, `title`/`executablePath` may be empty because the window is already gone.

```ts
import dwmWindows from 'dwm-windows';

dwmWindows.onWindowCreated(e => console.log('Created:', e));
dwmWindows.onWindowClosed(e => console.log('Closed:', e));
dwmWindows.onWindowFocused(e => console.log('Focused:', e));

// Later, to clean up
dwmWindows.stopWindowEvents();
```

### Filter Methods

#### `getWindowsByTitle(titleFilter: string): WindowInfo[]`
Filters windows by title (case-insensitive substring match).

#### `getWindowsByExecutable(executableName: string): WindowInfo[]`
Filters windows by executable name (e.g., "notepad.exe").

### WindowInfo Interface

```typescript
interface WindowInfo {
  id: number;                // HWND as number (same as hwnd)
  title: string;             // Window title
  executablePath: string;    // Full path to executable
  isVisible: boolean;        // Visibility state
  hwnd: number;             // Windows handle (same as id)
  thumbnail: string;        // PNG thumbnail (base64 data URL)
  icon: string;             // App icon (base64 data URL)
}
```

## 💡 Usage Examples

### Virtual Desktop Comparison

```typescript
import dwmWindows from 'dwm-windows';

const currentDesktop = dwmWindows.getVisibleWindows();
const allDesktops = dwmWindows.getVisibleWindows({ includeAllDesktops: true });

// Find windows that exist on other desktops but not current one
// Note: id equals hwnd (native handle), which is stable for the window's lifetime
const currentWindowKeys = new Set(
  currentDesktop.map(w => `${w.title}|${w.executablePath}`)
);
const hiddenWindows = allDesktops.filter(w => 
  !currentWindowKeys.has(`${w.title}|${w.executablePath}`)
);

console.log(`Current desktop: ${currentDesktop.length} windows`);
console.log(`Hidden on other desktops: ${hiddenWindows.length} windows`);
```

### Window Management

```typescript
// Find all Chrome windows
const chromeWindows = dwmWindows.getWindowsByExecutable('chrome.exe');

// Focus the first Chrome window
if (chromeWindows.length > 0) {
  dwmWindows.openWindow(chromeWindows[0].id);
}

// Find windows by title
const codeWindows = dwmWindows.getWindowsByTitle('Visual Studio Code');
```

### Thumbnail Analysis

```typescript
const windows = dwmWindows.getVisibleWindows();

const getDataSize = (dataUrl: string) => {
  const base64 = dataUrl.split(',')[1] || '';
  return Math.round(base64.length * 0.75 / 1024); // KB estimate
};

const totalIconSize = windows.reduce((sum, w) => sum + getDataSize(w.icon), 0);
const totalThumbSize = windows.reduce((sum, w) => sum + getDataSize(w.thumbnail), 0);

console.log(`Total payload: ${totalIconSize + totalThumbSize}KB`);
console.log(`Icons: ${totalIconSize}KB, Thumbnails: ${totalThumbSize}KB`);
```

### Live Updates with Events

```ts
import dwmWindows from 'dwm-windows';

const windows = new Map<number, { title: string }>();
dwmWindows.getVisibleWindows().forEach(w => windows.set(w.id, { title: w.title }));

dwmWindows.onWindowCreated(e => {
  windows.set(e.id, { title: e.title });
  console.log('Added', e.id, e.title);
});

dwmWindows.onWindowClosed(e => {
  windows.delete(e.id);
  console.log('Removed', e.id);
});

dwmWindows.onWindowFocused(e => {
  console.log('Focused', e.id, e.title);
});
```

## 🏗️ Development

### Prerequisites

- Windows 10/11
- Node.js ≥ 16
- Visual Studio Build Tools or Visual Studio with C++ workload
- Python (for node-gyp)

### Building

```bash
# Install dependencies and build
yarn install

# Rebuild native module only
yarn rebuild

# Build TypeScript and native module
yarn build

# Run example
yarn example
```

Notes:
- Windows Graphics Capture is enabled by default for robust, flicker-free captures when supported.
- Minimized windows use DWM previews; if only a tiny title bar is available, a placeholder thumbnail (centered app icon) is shown instead of a low-quality image.

### Project Structure

```
dwm-windows/
├── src/
│   ├── index.ts      # Main API
│   ├── types.d.ts    # Type definitions
│   └── example.ts    # Usage examples
├── dwm_thumbnail.cc  # C++ native bindings
├── binding.gyp       # Node-gyp build configuration
└── build/            # Compiled binaries
```

## 🔧 Requirements

- **OS**: Windows 10/11 (x64)
- **Node.js**: ≥ 16.0.0
- **Architecture**: x64 only

## 📄 License

MIT © Giacomo Barbalinardo

## 🤝 Contributing

Contributions welcome! Please read our contributing guidelines and submit pull requests for any improvements.

---

<div align="center">
  <strong>Built with ❤️ for Windows and Node devs</strong>
</div>
