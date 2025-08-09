export interface WindowInfo {
  id: number;
  title: string;
  executablePath: string;
  isVisible: boolean;
  hwnd: number;
  thumbnail: string; // data URL (PNG base64)
  icon: string; // data URL (PNG base64)
}

export interface DwmWindows {
  /**
   * Get all windows with their thumbnails
   * @returns Array of window information including base64-encoded thumbnails
   */
  getWindows(): WindowInfo[];
  getWindows(options: { includeAllDesktops?: boolean } | boolean): WindowInfo[];
  getWindowsAsync(): Promise<WindowInfo[]>;
  getWindowsAsync(options: { includeAllDesktops?: boolean } | boolean): Promise<WindowInfo[]>;

  /**
   * Update thumbnail for a specific window
   * @param windowId The window ID to update
   * @returns Updated base64-encoded thumbnail
   */
  updateThumbnail(windowId: number): string;
  updateThumbnailAsync(windowId: number): Promise<string>;

  /**
   * Bring a window to the foreground and focus it
   * @param windowId The window ID to open/focus
   * @returns True if successful
   */
  openWindow(windowId: number): boolean;
  openWindowAsync(windowId: number): Promise<boolean>;

  getVisibleWindows(options?: { includeAllDesktops?: boolean } | boolean): WindowInfo[];
  getVisibleWindowsAsync(options?: { includeAllDesktops?: boolean } | boolean): Promise<WindowInfo[]>;
  getWindowsAllDesktops(): WindowInfo[];
  getWindowsAllDesktopsAsync(): Promise<WindowInfo[]>;
}

declare const dwmWindows: DwmWindows;
export default dwmWindows;