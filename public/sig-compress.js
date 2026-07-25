// Compress signature canvases for storage / LAN transfer.
// Scales down (default max 640×240) and picks JPEG vs PNG by size.
// Usage: window.aisCompressSignature(sourceCanvas) → data URL string
(function (global) {
  'use strict';

  const DEFAULTS = {
    maxWidth: 640,
    maxHeight: 240,
    jpegQuality: 0.82,
    // Prefer JPEG unless PNG is meaningfully smaller (line-art edge cases).
    preferJpegIfSmallerBy: 0.92
  };

  function compressSignatureCanvas(sourceCanvas, options) {
    if (!sourceCanvas || !sourceCanvas.width || !sourceCanvas.height) return null;
    const opts = Object.assign({}, DEFAULTS, options || {});

    const srcW = sourceCanvas.width;
    const srcH = sourceCanvas.height;
    const scale = Math.min(1, opts.maxWidth / srcW, opts.maxHeight / srcH);
    const outW = Math.max(1, Math.round(srcW * scale));
    const outH = Math.max(1, Math.round(srcH * scale));

    const out = document.createElement('canvas');
    out.width = outW;
    out.height = outH;
    const ctx = out.getContext('2d');
    // White background (no transparency) — better for print + JPEG.
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, outW, outH);
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(sourceCanvas, 0, 0, outW, outH);

    let jpeg = '';
    let png = '';
    try {
      jpeg = out.toDataURL('image/jpeg', opts.jpegQuality);
    } catch { /* old browser */ }
    try {
      png = out.toDataURL('image/png');
    } catch { /* ignore */ }

    if (jpeg && png) {
      return jpeg.length <= png.length * opts.preferJpegIfSmallerBy ? jpeg : png;
    }
    return jpeg || png || null;
  }

  global.aisCompressSignature = compressSignatureCanvas;
})(typeof window !== 'undefined' ? window : globalThis);
