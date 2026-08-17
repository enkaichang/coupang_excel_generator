/**
 * Coupang Excel Generator - Storage Utilities (IndexedDB & Config Management)
 * Handles template profiles, custom Excel template binary storage, and full backup export/import.
 */
(function() {
  const DB_NAME = 'CoupangGeneratorDB';
  const DB_VERSION = 1;
  const STORE_PROFILES = 'template_profiles';
  const STORE_TEMPLATES = 'templates_data';

  let dbInstance = null;

  function openDB() {
    if (dbInstance) return Promise.resolve(dbInstance);
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(STORE_PROFILES)) {
          db.createObjectStore(STORE_PROFILES, { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains(STORE_TEMPLATES)) {
          db.createObjectStore(STORE_TEMPLATES, { keyPath: 'id' });
        }
      };
      request.onsuccess = (e) => {
        dbInstance = e.target.result;
        resolve(dbInstance);
      };
      request.onerror = (e) => {
        console.error('IndexedDB open error:', e);
        reject(e.target.error);
      };
    });
  }

  async function getStore(storeName, mode = 'readonly') {
    const db = await openDB();
    const tx = db.transaction(storeName, mode);
    return tx.objectStore(storeName);
  }

  window.StorageUtils = {
    /**
     * Get all saved template profiles
     */
    async getAllProfiles() {
      try {
        const store = await getStore(STORE_PROFILES, 'readonly');
        return new Promise((resolve, reject) => {
          const req = store.getAll();
          req.onsuccess = () => resolve(req.result || []);
          req.onerror = () => reject(req.error);
        });
      } catch (err) {
        console.warn('StorageUtils.getAllProfiles error, fallback to localStorage:', err);
        const saved = localStorage.getItem('coupang_template_profiles');
        return saved ? JSON.parse(saved) : [];
      }
    },

    /**
     * Save or update a single template profile
     */
    async saveProfile(profile) {
      if (!profile || !profile.id) throw new Error('Invalid profile object');
      try {
        const store = await getStore(STORE_PROFILES, 'readwrite');
        return new Promise((resolve, reject) => {
          const req = store.put(profile);
          req.onsuccess = () => resolve(profile);
          req.onerror = () => reject(req.error);
        });
      } catch (err) {
        console.warn('StorageUtils.saveProfile error, fallback to localStorage:', err);
        const all = await this.getAllProfiles();
        const idx = all.findIndex(p => p.id === profile.id);
        if (idx >= 0) all[idx] = profile;
        else all.push(profile);
        localStorage.setItem('coupang_template_profiles', JSON.stringify(all));
        return profile;
      }
    },

    /**
     * Delete a template profile by ID
     */
    async deleteProfile(profileId) {
      try {
        const store = await getStore(STORE_PROFILES, 'readwrite');
        await new Promise((resolve, reject) => {
          const req = store.delete(profileId);
          req.onsuccess = () => resolve();
          req.onerror = () => reject(req.error);
        });
        await this.deleteTemplateData(profileId);
      } catch (err) {
        console.warn('StorageUtils.deleteProfile error:', err);
        const all = await this.getAllProfiles();
        const filtered = all.filter(p => p.id !== profileId);
        localStorage.setItem('coupang_template_profiles', JSON.stringify(filtered));
      }
    },

    /**
     * Get binary/base64 template data for a profile
     */
    async getTemplateData(profileId) {
      try {
        const store = await getStore(STORE_TEMPLATES, 'readonly');
        return new Promise((resolve, reject) => {
          const req = store.get(profileId);
          req.onsuccess = () => resolve(req.result ? req.result.data : null);
          req.onerror = () => reject(req.error);
        });
      } catch (err) {
        console.warn('StorageUtils.getTemplateData error:', err);
        const customTemplates = JSON.parse(localStorage.getItem('coupang_templates') || '{}');
        return customTemplates[profileId] || null;
      }
    },

    /**
     * Save binary/base64 template data for a profile
     */
    async saveTemplateData(profileId, dataBase64) {
      try {
        const store = await getStore(STORE_TEMPLATES, 'readwrite');
        return new Promise((resolve, reject) => {
          const req = store.put({ id: profileId, data: dataBase64, updatedAt: Date.now() });
          req.onsuccess = () => resolve();
          req.onerror = () => reject(req.error);
        });
      } catch (err) {
        console.warn('StorageUtils.saveTemplateData error:', err);
        const customTemplates = JSON.parse(localStorage.getItem('coupang_templates') || '{}');
        customTemplates[profileId] = dataBase64;
        try {
          localStorage.setItem('coupang_templates', JSON.stringify(customTemplates));
        } catch (e) {
          console.error('LocalStorage quota exceeded for template data:', e);
        }
      }
    },

    /**
     * Delete template binary data for a profile
     */
    async deleteTemplateData(profileId) {
      try {
        const store = await getStore(STORE_TEMPLATES, 'readwrite');
        return new Promise((resolve, reject) => {
          const req = store.delete(profileId);
          req.onsuccess = () => resolve();
          req.onerror = () => reject(req.error);
        });
      } catch (err) {
        console.warn('StorageUtils.deleteTemplateData error:', err);
      }
    },

    /**
     * Clear all template profiles
     */
    async clearAllProfiles() {
      try {
        const store = await getStore(STORE_PROFILES, 'readwrite');
        await new Promise((resolve, reject) => {
          const req = store.clear();
          req.onsuccess = () => resolve();
          req.onerror = () => reject(req.error);
        });
      } catch (err) {
        console.warn('StorageUtils.clearAllProfiles error:', err);
      }
      localStorage.removeItem('coupang_template_profiles');
    },

    /**
     * Clear all template binary data
     */
    async clearAllTemplates() {
      try {
        const store = await getStore(STORE_TEMPLATES, 'readwrite');
        await new Promise((resolve, reject) => {
          const req = store.clear();
          req.onsuccess = () => resolve();
          req.onerror = () => reject(req.error);
        });
      } catch (err) {
        console.warn('StorageUtils.clearAllTemplates error:', err);
      }
      localStorage.removeItem('coupang_templates');
    },

    /**
     * Export all configuration (profiles, category rules, aliases, and custom templates) to a single JSON package
     */
    async exportConfigPackage() {
      const profiles = await this.getAllProfiles();
      const templatesData = {};
      for (const p of profiles) {
        if (!p.is_builtin) {
          const data = await this.getTemplateData(p.id);
          if (data) templatesData[p.id] = data;
        }
      }

      const customTemplates = JSON.parse(localStorage.getItem('coupang_templates') || '{}');
      if (customTemplates.HARNESS) templatesData['HARNESS'] = customTemplates.HARNESS;
      if (customTemplates.LEASH) templatesData['LEASH'] = customTemplates.LEASH;

      const sourceConfig = window.AppConfig ? window.AppConfig.get().source : {};

      return {
        version: '2.0.0',
        exportedAt: new Date().toISOString(),
        sourceConfig,
        profiles,
        templatesData
      };
    },

    /**
     * Import full configuration package from JSON object
     */
    async importConfigPackage(pkg) {
      if (!pkg || typeof pkg !== 'object') throw new Error('無效的設定檔格式');
      
      if (pkg.sourceConfig) {
        const current = window.AppConfig.get();
        current.source = { ...current.source, ...pkg.sourceConfig };
        localStorage.setItem('coupang_config', JSON.stringify(current));
      }

      if (Array.isArray(pkg.categoryRules)) {
        localStorage.setItem('coupang_category_rules', JSON.stringify(pkg.categoryRules));
      }

      if (Array.isArray(pkg.profiles)) {
        await this.clearAllProfiles();
        for (const p of pkg.profiles) {
          await this.saveProfile(p);
        }
      }

      if (pkg.templatesData && typeof pkg.templatesData === 'object') {
        await this.clearAllTemplates();
        for (const [id, data] of Object.entries(pkg.templatesData)) {
          if (['HARNESS', 'LEASH'].includes(id)) {
            const customTemplates = JSON.parse(localStorage.getItem('coupang_templates') || '{}');
            customTemplates[id] = data;
            localStorage.setItem('coupang_templates', JSON.stringify(customTemplates));
          } else {
            await this.saveTemplateData(id, data);
          }
        }
      }

      return true;
    }
  };
})();
