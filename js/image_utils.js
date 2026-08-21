window.ImageUtils = {
  loadImage: function(fileOrBlob) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(fileOrBlob);
      const img = new Image();
      img.onload = () => {
        URL.revokeObjectURL(url);
        resolve(img);
      };
      img.onerror = () => {
        URL.revokeObjectURL(url);
        reject(new Error("Image load failed"));
      };
      img.src = url;
    });
  },
  
  ensureMinShortEdge: function(img, minEdge = 1000) {
    return new Promise((resolve) => {
      const w = img.width || 0;
      const h = img.height || 0;
      if (w <= 0 || h <= 0) {
        resolve(null);
        return;
      }
      const shortEdge = Math.min(w, h);
      
      if (shortEdge >= minEdge) {
        // Just return original blob via canvas if already large enough
        const canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0);
        canvas.toBlob(b => resolve(b), 'image/jpeg', 0.95);
        return;
      }
      
      const ratio = minEdge / shortEdge;
      const newW = Math.round(w * ratio);
      const newH = Math.round(h * ratio);
      
      const canvas = document.createElement('canvas');
      canvas.width = newW;
      canvas.height = newH;
      const ctx = canvas.getContext('2d');
      // basic smoothing
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(img, 0, 0, newW, newH);
      
      canvas.toBlob(b => resolve(b), 'image/jpeg', 0.95);
    });
  },

  resizeAndPad: function(img, targetW = 1000, targetH = 1000, padColor = '#FFFFFF') {
    return new Promise((resolve) => {
      const w = img.width || 0;
      const h = img.height || 0;
      if (w <= 0 || h <= 0) {
        resolve(null);
        return;
      }

      // Calculate scale ratio to fit within targetW and targetH
      const ratio = Math.min(targetW / w, targetH / h);
      const newW = Math.round(w * ratio);
      const newH = Math.round(h * ratio);

      const canvas = document.createElement('canvas');
      canvas.width = targetW;
      canvas.height = targetH;
      const ctx = canvas.getContext('2d');

      // Fill with padColor
      ctx.fillStyle = padColor;
      ctx.fillRect(0, 0, targetW, targetH);

      // Draw image centered
      const x = (targetW - newW) / 2;
      const y = (targetH - newH) / 2;

      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(img, x, y, newW, newH);

      canvas.toBlob(b => resolve(b), 'image/jpeg', 0.95);
    });
  }
};
