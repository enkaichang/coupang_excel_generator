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
  
  ensureMinWidth: function(img, minW = 600) {
    return new Promise((resolve) => {
      const w = img.width;
      const h = img.height;
      if (w > 500) {
        // Just return original blob via canvas if already large enough
        const canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0);
        canvas.toBlob(b => resolve(b), 'image/jpeg', 0.95);
        return;
      }
      
      const targetW = Math.max(minW, 600);
      const ratio = targetW / w;
      const newW = targetW;
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
  }
};
