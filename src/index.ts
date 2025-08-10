import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const require = createRequire(import.meta.url);

// Load the native module
const nativeModule = require(join(__dirname, '../build/Release/dwm_windows.node'));

export interface WindowInfo {
  id: number; // equals hwnd (native window handle)
  title: string;
  executablePath: string;
  isVisible: boolean;
  hwnd: number; // same as id
  thumbnail: string; // data URL (PNG base64)
  icon: string; // data URL (PNG base64)
}

export class DwmWindows {
  /**
   * Get all windows with their thumbnails
   * @returns Array of window information including base64-encoded thumbnails
   */
  public getWindows(options?: { includeAllDesktops?: boolean } | boolean): WindowInfo[] {
    try {
      // Back-compat: no args -> current desktop only
      if (options === undefined) return nativeModule.getWindows();
      // Boolean shorthand
      if (typeof options === 'boolean') return nativeModule.getWindows(options);
      return nativeModule.getWindows({ includeAllDesktops: !!options.includeAllDesktops });
    } catch (error) {
      console.error('Error getting windows:', error);
      return [];
    }
  }

  /**
   * Async: Get all windows with their thumbnails without blocking the event loop
   */
  public async getWindowsAsync(options?: { includeAllDesktops?: boolean } | boolean): Promise<WindowInfo[]> {
    try {
      if (options === undefined) return await nativeModule.getWindowsAsync();
      if (typeof options === 'boolean') return await nativeModule.getWindowsAsync(options);
      return await nativeModule.getWindowsAsync({ includeAllDesktops: !!options.includeAllDesktops });
    } catch (error) {
      console.error('Error getting windows (async):', error);
      return [];
    }
  }

  /**
   * Update thumbnail for a specific window
   * @param windowId The window ID to update
   * @returns Updated base64-encoded thumbnail
   */
  public updateThumbnail(windowId: number): string {
    try {
      return nativeModule.updateThumbnail(windowId);
    } catch (error) {
      console.error('Error updating thumbnail:', error);
  return 'data:image/png;base64,';
    }
  }

  /**
   * Async: Update thumbnail for a specific window without blocking
   */
  public async updateThumbnailAsync(windowId: number): Promise<string> {
    try {
      return await nativeModule.updateThumbnailAsync(windowId);
    } catch (error) {
      console.error('Error updating thumbnail (async):', error);
      return 'data:image/png;base64,';
    }
  }

  /**
   * Bring a window to the foreground and focus it
   * @param windowId The window ID to open/focus
   * @returns True if successful
   */
  public openWindow(windowId: number): boolean {
    try {
      return nativeModule.openWindow(windowId);
    } catch (error) {
      console.error('Error opening window:', error);
      return false;
    }
  }

  /**
   * Async: Bring a window to the foreground without blocking
   */
  public async openWindowAsync(windowId: number): Promise<boolean> {
    try {
      return await nativeModule.openWindowAsync(windowId);
    } catch (error) {
      console.error('Error opening window (async):', error);
      return false;
    }
  }

  /**
   * Filter windows by title (case-insensitive)
   * @param titleFilter Filter string for window titles
   * @returns Filtered array of windows
   */
  public getWindowsByTitle(titleFilter: string): WindowInfo[] {
    const windows = this.getWindows();
    const filter = titleFilter.toLowerCase();
    return windows.filter(window => 
      window.title.toLowerCase().includes(filter)
    );
  }

  /**
   * Filter windows by executable name
   * @param executableName Name of the executable (e.g., "notepad.exe")
   * @returns Filtered array of windows
   */
  public getWindowsByExecutable(executableName: string): WindowInfo[] {
    const windows = this.getWindows();
    const filter = executableName.toLowerCase();
    return windows.filter(window => 
      window.executablePath.toLowerCase().includes(filter)
    );
  }

  /**
   * Get only visible windows
   * @returns Array of visible windows only
   */
  public getVisibleWindows(options?: { includeAllDesktops?: boolean } | boolean): WindowInfo[] {
    const windows = this.getWindows(options);
    return windows.filter(window => window.isVisible);
  }

  /**
   * Async: Get only visible windows (non-blocking)
   */
  public async getVisibleWindowsAsync(options?: { includeAllDesktops?: boolean } | boolean): Promise<WindowInfo[]> {
    const windows = await this.getWindowsAsync(options);
    return windows.filter(window => window.isVisible);
  }

  /**
   * Convenience: get windows from all virtual desktops
   */
  public getWindowsAllDesktops(): WindowInfo[] {
    return this.getWindows({ includeAllDesktops: true });
  }

  /**
   * Async convenience: get windows from all virtual desktops
   */
  public getWindowsAllDesktopsAsync(): Promise<WindowInfo[]> {
    return this.getWindowsAsync({ includeAllDesktops: true });
  }
}

// Create and export default instance
const dwmWindows = new DwmWindows();
export default dwmWindows;