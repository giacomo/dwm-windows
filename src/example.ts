import dwmWindows from './index.js';

const dataUrlBytes = (dataUrl?: string) => {
	if (!dataUrl || typeof dataUrl !== 'string') return 0;
	const idx = dataUrl.indexOf(',');
	if (idx < 0) return 0;
	const b64 = dataUrl.slice(idx + 1);
	try { return Buffer.from(b64, 'base64').byteLength; } catch { return 0; }
};

const fmtKB = (bytes: number) => `${(bytes / 1024).toFixed(1)}KB`;

// Aktueller Desktop (Task View-ähnlich)
const current = dwmWindows.getVisibleWindows();
console.log(`Current desktop windows: ${current.length}`);
for (const w of current) {
	const hasIcon = typeof w.icon === 'string' && w.icon.startsWith('data:image/png;base64,');
	const hasThumb = typeof w.thumbnail === 'string' && w.thumbnail.startsWith('data:image/png;base64,');
	const iconBytes = hasIcon ? dataUrlBytes(w.icon) : 0;
	const thumbBytes = hasThumb ? dataUrlBytes(w.thumbnail) : 0;
	const sizeInfo = `${hasIcon ? ' [icon ' + fmtKB(iconBytes) + ']' : ''}${hasThumb ? ' [thumb ' + fmtKB(thumbBytes) + ']' : ''}`;
	console.log(`- ${w.id} ${w.title} (${w.executablePath})${sizeInfo}`);
}

// Alle virtuellen Desktops
const all = dwmWindows.getVisibleWindows({ includeAllDesktops: true });
console.log(`\nAll desktops windows: ${all.length}`);
for (const w of all) {
	const hasIcon = typeof w.icon === 'string' && w.icon.startsWith('data:image/png;base64,');
	const hasThumb = typeof w.thumbnail === 'string' && w.thumbnail.startsWith('data:image/png;base64,');
	const iconBytes = hasIcon ? dataUrlBytes(w.icon) : 0;
	const thumbBytes = hasThumb ? dataUrlBytes(w.thumbnail) : 0;
	const sizeInfo = `${hasIcon ? ' [icon ' + fmtKB(iconBytes) + ']' : ''}${hasThumb ? ' [thumb ' + fmtKB(thumbBytes) + ']' : ''}`;
	console.log(`- ${w.id} ${w.title} (${w.executablePath})${sizeInfo}`);
}


// compare current and all and get the different
// id equals hwnd (native handle), stable while the window exists
const currentIds = new Set(current.map(w => w.id));
const different = all.filter(w => !currentIds.has(w.id));
console.log(`\nDifferent windows (not on current desktop): ${different.length}`);
for (const w of different) console.log(`- ${w.title} (${w.executablePath})`);

// Summary to visualize payload sizes
const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
const iconSizes = current.map(w => dataUrlBytes(w.icon));
const thumbSizes = current.map(w => dataUrlBytes(w.thumbnail));
console.log(`\nPNG sizes on current desktop:`);
console.log(`- Icons: total ${fmtKB(sum(iconSizes))}, avg ${fmtKB(sum(iconSizes)/Math.max(1,current.length))}`);
console.log(`- Thumbnails: total ${fmtKB(sum(thumbSizes))}, avg ${fmtKB(sum(thumbSizes)/Math.max(1,current.length))}`);

// --- Async API demo (non-blocking) ---
(async () => {
	console.log('\n[Async] Fetching visible windows...');
	const asyncVisible = await dwmWindows.getVisibleWindowsAsync();
	console.log(`[Async] Visible windows: ${asyncVisible.length}`);
	if (asyncVisible[0]) {
		const first = asyncVisible[0];
		console.log(`[Async] Opening first window: ${first.title}`);
		const ok = await dwmWindows.openWindowAsync(first.id);
		console.log(`[Async] openWindowAsync ok: ${ok}`);
		const thumb = await dwmWindows.updateThumbnailAsync(first.id);
		console.log(`[Async] New thumbnail size: ${fmtKB(dataUrlBytes(thumb))}`);
	}

	const asyncAll = await dwmWindows.getWindowsAsync({ includeAllDesktops: true });
	console.log(`[Async] All desktops windows: ${asyncAll.length}`);
})();