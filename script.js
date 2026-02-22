// Tree-based Version Control System
class VersionControl {
    constructor(maxVersions = 100) {
        this.versions = new Map();      // Version ID -> Version node mapping
        this.rootId = null;             // Root version ID
        this.currentId = null;          // Current version ID
        this.maxVersions = maxVersions;
        this.loadHistory();
    }

    // Generate unique ID
    generateId() {
        return Date.now().toString(36) + Math.random().toString(36).slice(2);
    }

    // Load version history from localStorage
    loadHistory() {
        try {
            const savedVersions = localStorage.getItem('wordMemoryVersions');
            const savedMeta = localStorage.getItem('wordMemoryVersionMeta');
            const savedSettings = localStorage.getItem('wordMemorySettings');

            // Check if it's old linear format (array) and migrate
            if (savedVersions) {
                const parsed = JSON.parse(savedVersions);

                if (Array.isArray(parsed)) {
                    // Old linear format - migrate to tree structure
                    this.migrateFromLinear(parsed);
                } else {
                    // New tree format
                    this.versions = new Map(Object.entries(parsed));
                }
            }

            if (savedMeta) {
                const meta = JSON.parse(savedMeta);
                this.rootId = meta.rootId;
                this.currentId = meta.currentId;
            }

            if (savedSettings) {
                const settings = JSON.parse(savedSettings);
                this.maxVersions = settings.maxVersions || 100;
            }
        } catch (error) {
            this.versions = new Map();
            this.rootId = null;
            this.currentId = null;
        }
    }

    // Migrate from old linear format to tree structure
    migrateFromLinear(linearVersions) {
        if (linearVersions.length === 0) return;

        // Convert linear array to tree (single branch)
        let previousId = null;

        for (let i = 0; i < linearVersions.length; i++) {
            const oldVersion = linearVersions[i];
            const newId = oldVersion.id || this.generateId();

            const newVersion = {
                id: newId,
                parentId: previousId,
                children: [],
                timestamp: oldVersion.timestamp,
                data: oldVersion.data,
                description: oldVersion.description || 'Version',
                wordCount: oldVersion.wordCount || (oldVersion.data ? oldVersion.data.length : 0),
                lastAccessed: oldVersion.timestamp
            };

            // Add as child of previous version
            if (previousId) {
                const parent = this.versions.get(previousId);
                if (parent) {
                    parent.children.push(newId);
                }
            } else {
                // First version is root
                this.rootId = newId;
            }

            this.versions.set(newId, newVersion);
            previousId = newId;
        }

        // Set current to last version
        this.currentId = previousId;

        // Save migrated data
        this.saveHistory();
    }

    // Save version history to localStorage
    saveHistory() {
        try {
            // Save versions as object (will be converted to Map on load)
            const versionsObj = Object.fromEntries(this.versions);
            localStorage.setItem('wordMemoryVersions', JSON.stringify(versionsObj));

            // Save metadata
            const meta = {
                rootId: this.rootId,
                currentId: this.currentId
            };
            localStorage.setItem('wordMemoryVersionMeta', JSON.stringify(meta));
        } catch (error) {
            // If localStorage is full, try to remove old versions
            if (error.name === 'QuotaExceededError') {
                const toRemove = Math.max(10, Math.floor(this.maxVersions / 4));
                this.pruneOldVersions(toRemove);
                try {
                    const versionsObj = Object.fromEntries(this.versions);
                    localStorage.setItem('wordMemoryVersions', JSON.stringify(versionsObj));
                    const meta = { rootId: this.rootId, currentId: this.currentId };
                    localStorage.setItem('wordMemoryVersionMeta', JSON.stringify(meta));
                } catch (e) {
                }
            }
        }
    }

    // --- Delta compression helpers ---

    // Compute delta: mixed array of parent indices (number) and new items (object)
    computeDelta(parentData, newData) {
        const fingerprints = new Map(); // JSON string -> [indices]
        parentData.forEach((item, i) => {
            const fp = JSON.stringify(item);
            if (!fingerprints.has(fp)) fingerprints.set(fp, []);
            fingerprints.get(fp).push(i);
        });

        return newData.map(item => {
            const fp = JSON.stringify(item);
            const indices = fingerprints.get(fp);
            if (indices && indices.length > 0) {
                return indices.shift(); // Reference parent by index
            }
            return JSON.parse(JSON.stringify(item)); // New/modified item
        });
    }

    // Apply delta to reconstruct full data
    applyDelta(parentData, delta) {
        return delta.map(entry =>
            typeof entry === 'number' ? parentData[entry] : entry
        );
    }

    // Count consecutive delta depth from a version up to nearest snapshot
    _deltaDepth(versionId) {
        let depth = 0;
        let id = versionId;
        while (id) {
            const v = this.versions.get(id);
            if (!v || v.data) break; // snapshot or missing → stop
            depth++;
            id = v.parentId;
        }
        return depth;
    }

    // Resolve full data for any version (handles delta chains + caching)
    resolveData(versionId) {
        const version = this.versions.get(versionId);
        if (!version) return null;

        // Check cache
        if (!this._dataCache) this._dataCache = new Map();
        if (this._dataCache.has(versionId)) return this._dataCache.get(versionId);

        let result;
        if (version.data) {
            // Snapshot (or legacy version with data field)
            result = version.data;
        } else if (version.delta && version.parentId) {
            // Delta version — resolve parent first, then apply
            const parentData = this.resolveData(version.parentId);
            if (!parentData) return null;
            result = this.applyDelta(parentData, version.delta);
        } else {
            return null;
        }

        // Cache (LRU, keep max 5)
        this._dataCache.set(versionId, result);
        if (this._dataCache.size > 5) {
            const oldest = this._dataCache.keys().next().value;
            this._dataCache.delete(oldest);
        }

        return result;
    }

    // Invalidate data cache
    _clearCache() {
        if (this._dataCache) this._dataCache.clear();
    }

    // Create a new version (creates child version under current version)
    createVersion(data, description = 'Manual save') {
        const now = new Date().toISOString();
        const newVersion = {
            id: this.generateId(),
            parentId: this.currentId,
            children: [],
            timestamp: now,
            description: description,
            wordCount: data.length,
            lastAccessed: now
        };

        // Decide: snapshot or delta
        let useSnapshot = true;
        if (this.currentId) {
            const parentData = this.resolveData(this.currentId);
            if (parentData) {
                const delta = this.computeDelta(parentData, data);
                const deltaSize = JSON.stringify(delta).length;
                const fullSize = JSON.stringify(data).length;
                const depth = this._deltaDepth(this.currentId);

                // Use delta if: not too deep, and delta is meaningfully smaller
                if (depth < 9 && deltaSize < fullSize * 0.6) {
                    newVersion.delta = delta;
                    useSnapshot = false;
                }
            }
        }

        if (useSnapshot) {
            newVersion.data = JSON.parse(JSON.stringify(data)); // Deep copy
        }

        // If current version exists, add as its child
        if (this.currentId) {
            const parent = this.versions.get(this.currentId);
            if (parent) {
                parent.children.push(newVersion.id);
            }
        } else {
            // First version, set as root
            this.rootId = newVersion.id;
        }

        this.versions.set(newVersion.id, newVersion);
        this.currentId = newVersion.id;
        this._clearCache();

        // Limit number of versions (keep recently accessed)
        if (this.versions.size > this.maxVersions) {
            this.pruneOldVersions();
        }

        this.saveHistory();
        updateHistoryButtonLabel(this.versions.size);
        return newVersion;
    }

    // Smart Undo (prefer current branch, choose most recent at forks)
    undo() {
        if (!this.canUndo()) return null;

        const current = this.versions.get(this.currentId);
        if (!current || !current.parentId) return null; // Already at root

        this.currentId = current.parentId;
        const parent = this.versions.get(this.currentId);
        if (parent) {
            parent.lastAccessed = new Date().toISOString();
            this.saveHistory();
        }
        return parent;
    }

    // Smart Redo (choose most recently accessed child branch)
    redo() {
        if (!this.canRedo()) return null;

        const current = this.versions.get(this.currentId);
        if (!current || current.children.length === 0) return null; // No children

        // Choose most recently accessed child
        const childId = this.getMostRecentChild(current.children);
        this.currentId = childId;
        const child = this.versions.get(childId);
        if (child) {
            child.lastAccessed = new Date().toISOString();
            this.saveHistory();
        }
        return child;
    }

    // Get most recently accessed child version
    getMostRecentChild(childIds) {
        if (childIds.length === 0) return null;
        if (childIds.length === 1) return childIds[0];

        let mostRecent = childIds[0];
        let mostRecentTime = this.versions.get(mostRecent).lastAccessed;

        for (let i = 1; i < childIds.length; i++) {
            const childId = childIds[i];
            const child = this.versions.get(childId);
            if (child && child.lastAccessed > mostRecentTime) {
                mostRecent = childId;
                mostRecentTime = child.lastAccessed;
            }
        }
        return mostRecent;
    }

    // Check if can undo
    canUndo() {
        if (!this.currentId) return false;
        const current = this.versions.get(this.currentId);
        return current && current.parentId !== null;
    }

    // Check if can redo
    canRedo() {
        if (!this.currentId) return false;
        const current = this.versions.get(this.currentId);
        return current && current.children.length > 0;
    }

    // Get current version
    getCurrentVersion() {
        if (!this.currentId) return null;
        return this.versions.get(this.currentId);
    }

    // Go to specific version by ID
    goToVersion(versionId) {
        if (!this.versions.has(versionId)) return null;

        this.currentId = versionId;
        const version = this.versions.get(versionId);
        if (version) {
            version.lastAccessed = new Date().toISOString();
            this.saveHistory();
        }
        return version;
    }

    // Graft an imported version tree as a fork under the local root
    graftTree(importedVersionsObj, importedRootId, importedCurrentId) {
        const importedVersions = new Map(Object.entries(importedVersionsObj));

        // Build a temporary VersionControl to resolve deltas within the imported tree
        const tempVC = new VersionControl();
        tempVC.versions = importedVersions;
        tempVC.rootId = importedRootId;
        tempVC.currentId = importedCurrentId;

        // Build ID remap: old ID -> new ID
        const idMap = new Map();
        for (const oldId of importedVersions.keys()) {
            idMap.set(oldId, this.generateId());
        }

        // Re-key and insert all imported versions
        for (const [oldId, version] of importedVersions) {
            const newId = idMap.get(oldId);
            const newVersion = {
                ...version,
                id: newId,
                parentId: oldId === importedRootId
                    ? this.rootId  // graft point: local root
                    : (idMap.get(version.parentId) || null),
                children: (version.children || []).map(cid => idMap.get(cid)).filter(Boolean)
            };

            // Imported root's parent changes, so if it's a delta, convert to snapshot
            if (oldId === importedRootId && newVersion.delta && !newVersion.data) {
                const resolved = tempVC.resolveData(oldId);
                if (resolved) {
                    newVersion.data = JSON.parse(JSON.stringify(resolved));
                    delete newVersion.delta;
                }
            }

            this.versions.set(newId, newVersion);
        }

        // Add imported root as child of local root
        const localRoot = this.versions.get(this.rootId);
        if (localRoot) {
            localRoot.children.push(idMap.get(importedRootId));
        }

        // Set current to imported current
        this.currentId = idMap.get(importedCurrentId) || this.currentId;
        this._clearCache();

        // Prune if needed
        if (this.versions.size > this.maxVersions) {
            this.pruneOldVersions();
        }

        this.saveHistory();
        return idMap.get(importedCurrentId);
    }

    // Check if version is in current path (from current to root)
    isInCurrentPath(versionId) {
        let current = this.currentId;
        while (current) {
            if (current === versionId) return true;
            const version = this.versions.get(current);
            current = version ? version.parentId : null;
        }
        return false;
    }

    // Delete version and all its children (recursive)
    deleteVersionAndChildren(versionId) {
        const version = this.versions.get(versionId);
        if (!version) return;

        // Recursively delete all children first
        for (const childId of version.children) {
            this.deleteVersionAndChildren(childId);
        }

        // Delete this version
        this.versions.delete(versionId);
        this._clearCache();
    }

    // Delete only this version and reparent its children to its parent
    deleteVersionAndReparentChildren(versionId) {
        const version = this.versions.get(versionId);
        if (!version) return;
        if (versionId === this.rootId) return;

        const parentId = version.parentId;
        const parent = parentId ? this.versions.get(parentId) : null;
        const versionChildren = Array.isArray(version.children) ? version.children : [];
        const childrenToReparent = [...versionChildren];

        // Delta children reference this version as parent.
        // Convert them to snapshots before reparenting so they don't break.
        for (const childId of childrenToReparent) {
            const child = this.versions.get(childId);
            if (child && child.delta && !child.data) {
                const resolved = this.resolveData(childId);
                if (resolved) {
                    child.data = JSON.parse(JSON.stringify(resolved));
                    delete child.delta;
                }
            }
        }
        this._clearCache();

        // Reattach children in place of the deleted node to preserve branch order.
        if (parent) {
            const parentChildren = Array.isArray(parent.children) ? parent.children : [];
            const at = parentChildren.indexOf(versionId);
            parent.children = parentChildren.filter(id => id !== versionId);
            const insertAt = at >= 0 ? at : parent.children.length;
            parent.children.splice(insertAt, 0, ...childrenToReparent);
        }

        for (const childId of childrenToReparent) {
            const child = this.versions.get(childId);
            if (child) child.parentId = parentId || null;
        }

        version.children = [];
        this.versions.delete(versionId);
    }

    // Prune old versions (keep recently accessed, protect current path)
    pruneOldVersions(forceRemove = 0) {
        const targetSize = forceRemove > 0 ? this.versions.size - forceRemove : this.maxVersions;
        if (this.versions.size <= targetSize) return;

        const toRemove = this.versions.size - targetSize;

        // Get all versions sorted by last accessed time
        const allVersions = Array.from(this.versions.values())
            .sort((a, b) => new Date(a.lastAccessed) - new Date(b.lastAccessed));

        let removed = 0;
        for (const version of allVersions) {
            if (removed >= toRemove) break;

            // Don't delete versions in current path
            if (this.isInCurrentPath(version.id)) continue;

            // Don't delete root node
            if (version.id === this.rootId) continue;

            // Remove from parent's children array
            if (version.parentId) {
                const parent = this.versions.get(version.parentId);
                if (parent) {
                    parent.children = parent.children.filter(id => id !== version.id);
                }
            }

            // Delete this version and all its children
            this.deleteVersionAndChildren(version.id);
            removed++;
        }
    }

    // Get version tree for display
    getVersionTree() {
        if (!this.rootId) return [];

        const buildTree = (versionId, depth = 0) => {
            const version = this.versions.get(versionId);
            if (!version) return null;

            const node = {
                ...version,
                depth: depth,
                isCurrent: versionId === this.currentId,
                hasChildren: version.children.length > 0
            };

            const result = [node];

            // Recursively add children (sorted by timestamp)
            const sortedChildren = [...version.children].sort((a, b) => {
                const vA = this.versions.get(a);
                const vB = this.versions.get(b);
                return new Date(vA.timestamp) - new Date(vB.timestamp);
            });

            for (const childId of sortedChildren) {
                const childTree = buildTree(childId, depth + 1);
                if (childTree) {
                    result.push(...childTree);
                }
            }

            return result;
        };

        return buildTree(this.rootId);
    }

    // Get version history (for compatibility)
    getHistory() {
        return {
            versions: Array.from(this.versions.values()),
            currentIndex: -1, // Deprecated in tree structure
            canUndo: this.canUndo(),
            canRedo: this.canRedo(),
            tree: this.getVersionTree()
        };
    }

    // Clear all history
    clearHistory() {
        this.versions = new Map();
        this.rootId = null;
        this.currentId = null;
        this._clearCache();
        this.saveHistory();
    }

    // Update settings
    updateSettings(settings) {
        if (settings.maxVersions) {
            this.maxVersions = settings.maxVersions;
            localStorage.setItem('wordMemorySettings', JSON.stringify({ maxVersions: this.maxVersions }));

            // Trim versions if needed
            if (this.versions.size > this.maxVersions) {
                const toRemove = this.versions.size - this.maxVersions;
                this.pruneOldVersions(toRemove);
            }
        }
    }

    // Get settings
    getSettings() {
        return {
            maxVersions: this.maxVersions
        };
    }
}

// Project ID
const PROJECT_ID_KEY = 'wordMemoryProjectId';
const APP_SETTINGS_KEY = 'wordMemoryAppSettings';
const DEFAULT_APP_SETTINGS = {
    actionSoundEnabled: true,
    audioPreloadEnabled: true,
    pronounceVolume: 1,
    quizAutoPlay: false
};

function createProjectId() {
    const adj  = ['bold','calm','cool','dark','fast','gold','keen','lime','mint','neat','pink','pure','sage','slim','soft','warm','wild','grim','rosy','teal'];
    const noun = ['axe','bay','bee','cat','elk','fen','fox','gem','ivy','jay','oak','owl','pea','ray','rye','sea','sky','tern','dew','cod'];
    const a = adj[Math.floor(Math.random() * adj.length)];
    const n = noun[Math.floor(Math.random() * noun.length)];
    const num = String(Math.floor(Math.random() * 90) + 10);
    return `${a}-${n}-${num}`;
}

function setProjectId(id) {
    const normalizedId = String(id || '').trim();
    localStorage.setItem(PROJECT_ID_KEY, normalizedId);
    return normalizedId;
}

function getProjectId() {
    let id = localStorage.getItem(PROJECT_ID_KEY);
    if (!id) {
        id = setProjectId(createProjectId());
    }
    return id;
}

function toSafeFilenamePart(value) {
    const normalized = String(value || '').trim();
    const safe = normalized
        .replace(/[\\/:*?"<>|]/g, '-')
        .replace(/\s+/g, '-')
        .replace(/-+/g, '-')
        .replace(/^\.+|\.+$/g, '')
        .replace(/^-+|-+$/g, '');
    return safe || 'unknown';
}

function sanitizeAppSettings(rawSettings) {
    const parsed = rawSettings && typeof rawSettings === 'object' ? rawSettings : {};
    const actionSoundEnabled = typeof parsed.actionSoundEnabled === 'boolean'
        ? parsed.actionSoundEnabled
        : DEFAULT_APP_SETTINGS.actionSoundEnabled;
    const audioPreloadEnabled = typeof parsed.audioPreloadEnabled === 'boolean'
        ? parsed.audioPreloadEnabled
        : DEFAULT_APP_SETTINGS.audioPreloadEnabled;
    const parsedVolume = Number(parsed.pronounceVolume);
    const pronounceVolume = Number.isFinite(parsedVolume)
        ? Math.max(0, Math.min(1, parsedVolume))
        : DEFAULT_APP_SETTINGS.pronounceVolume;

    const quizAutoPlay = typeof parsed.quizAutoPlay === 'boolean'
        ? parsed.quizAutoPlay
        : DEFAULT_APP_SETTINGS.quizAutoPlay;

    return {
        actionSoundEnabled,
        audioPreloadEnabled,
        pronounceVolume,
        quizAutoPlay
    };
}

function getPronounceVolumePercent() {
    return Math.round((appSettings.pronounceVolume || 0) * 100);
}

function updatePronounceVolumePreview(value = null) {
    const input = document.getElementById('pronounceVolumeInput');
    const valueEl = document.getElementById('pronounceVolumeValue');
    if (!input || !valueEl) return;
    const rawValue = value === null ? input.value : value;
    const percent = Number.isFinite(Number(rawValue)) ? Math.max(0, Math.min(100, Math.round(Number(rawValue)))) : 100;
    valueEl.textContent = `${percent}%`;
}

function loadAppSettings() {
    try {
        const raw = localStorage.getItem(APP_SETTINGS_KEY);
        if (!raw) {
            return sanitizeAppSettings(null);
        }
        return sanitizeAppSettings(JSON.parse(raw));
    } catch (error) {
        return sanitizeAppSettings(null);
    }
}

function saveAppSettings() {
    try {
        localStorage.setItem(APP_SETTINGS_KEY, JSON.stringify(appSettings));
    } catch (error) {
    }
}

// Data storage
let words = [];
let isFullMode = false;
let hideMeaning = false;
let editingIndex = -1;
let selectedWords = new Set();
let isSelectMode = false;
let versionControl = null;
let wordSortMode = 'alpha'; // 'alpha', 'alpha-desc', 'chrono', 'chrono-desc'
let wordGroupMode = 'weight'; // 'weight', 'tag'
let activeTagFilters = new Set(); // tags selected for filtering display
let excludedTagFilters = new Set(); // tags excluded from display
let cachedVoices = [];
let appSettings = sanitizeAppSettings(null);

const CORE_ASSET_PATHS = [
    'assets/ding.mp3',
    'assets/delete.mp3',
    'assets/put.mp3',
    'assets/icon.png',
    'assets/title.png',
    'assets/webapp.png'
];
const CORE_ASSET_CACHE_NAME = 'word-memory-core-assets-v1';
const WORD_AUDIO_PRELOAD_CONCURRENCY = 3;

// Audio preload cache: word -> { status, url, blobUrl, buffer, promise }
const audioCache = new Map();
// status: 'idle' | 'fetching' | 'ready' | 'failed'
let audioPreloadObserver = null;
let deleteActionSoundAudio = null;
let putActionSoundAudio = null;
const preloadedAssetBlobUrls = new Map();
let coreAssetsWarmupPromise = null;
const wordAudioPreloadQueue = [];
const queuedWordAudio = new Set();
let wordAudioPreloadInFlight = 0;

function normalizeAudioWord(word) {
    return String(word || '').trim();
}

function getPreloadedAssetURL(path) {
    return preloadedAssetBlobUrls.get(path) || path;
}

async function _cacheAndFetchAsset(path) {
    if (!path) return null;
    let response = null;

    if ('caches' in window) {
        try {
            const cache = await caches.open(CORE_ASSET_CACHE_NAME);
            response = await cache.match(path);
            if (!response) {
                const fetched = await fetch(path, { cache: 'force-cache' });
                if (!fetched || !fetched.ok) return null;
                await cache.put(path, fetched.clone());
                response = fetched;
            }
        } catch (_) {}
    }

    if (!response) {
        try {
            response = await fetch(path, { cache: 'force-cache' });
        } catch (_) {
            return null;
        }
    }

    if (!response || !response.ok) return null;

    if (/\.(mp3|wav|ogg|m4a)$/i.test(path)) {
        try {
            const blob = await response.clone().blob();
            const blobUrl = URL.createObjectURL(blob);
            preloadedAssetBlobUrls.set(path, blobUrl);
        } catch (_) {}
    }

    return response;
}

function warmCoreAssets() {
    if (coreAssetsWarmupPromise) return coreAssetsWarmupPromise;
    coreAssetsWarmupPromise = Promise.allSettled(CORE_ASSET_PATHS.map(path => _cacheAndFetchAsset(path)));
    return coreAssetsWarmupPromise;
}

function _drainWordAudioPreloadQueue() {
    while (wordAudioPreloadInFlight < WORD_AUDIO_PRELOAD_CONCURRENCY && wordAudioPreloadQueue.length > 0) {
        const word = wordAudioPreloadQueue.shift();
        if (!word) continue;
        wordAudioPreloadInFlight++;
        preloadAudioForWord(word).finally(() => {
            queuedWordAudio.delete(word);
            wordAudioPreloadInFlight = Math.max(0, wordAudioPreloadInFlight - 1);
            if (wordAudioPreloadQueue.length > 0) {
                setTimeout(_drainWordAudioPreloadQueue, 0);
            }
        });
    }
}

function queueWordAudioPreload(word) {
    const normalizedWord = normalizeAudioWord(word);
    if (!normalizedWord) return;
    const existing = audioCache.get(normalizedWord);
    if (existing && (existing.status === 'ready' || existing.status === 'fetching')) return;
    if (queuedWordAudio.has(normalizedWord)) return;
    queuedWordAudio.add(normalizedWord);
    wordAudioPreloadQueue.push(normalizedWord);
    _drainWordAudioPreloadQueue();
}

function scheduleGlobalWordAudioPreload() {
    if (!appSettings.audioPreloadEnabled) return;
    words.forEach(w => {
        if (!w) return;
        if (!w.word || !(w.meaning || '').trim()) return;
        queueWordAudioPreload(w.word);
    });
}

function preloadAudioForWord(word) {
    const normalizedWord = normalizeAudioWord(word);
    if (!normalizedWord) return Promise.resolve(null);

    let entry = audioCache.get(normalizedWord);
    if (entry && entry.status === 'ready') return Promise.resolve(entry);
    if (entry && entry.status === 'fetching' && entry.promise) return entry.promise;

    if (!entry) {
        entry = {
            status: 'idle',
            url: null,
            blobUrl: null,
            buffer: null,
            promise: null
        };
        audioCache.set(normalizedWord, entry);
    }

    entry.status = 'fetching';
    entry.promise = fetch(`https://api.dictionaryapi.dev/api/v2/entries/en_GB/${encodeURIComponent(normalizedWord)}`, { cache: 'force-cache' })
        .then(r => {
            if (!r.ok) throw new Error('API failed');
            return r.json();
        })
        .then(data => {
            const audioUrl = data[0]?.phonetics?.find(p => p.audio)?.audio;
            if (!audioUrl) throw new Error('No audio URL');
            entry.url = audioUrl;
            return fetch(audioUrl, { cache: 'force-cache' })
                .then(r => {
                    if (!r.ok) throw new Error('Audio fetch failed');
                    return r.arrayBuffer();
                })
                .then(buf => {
                    try {
                        const blob = new Blob([buf], { type: 'audio/mpeg' });
                        if (entry.blobUrl) URL.revokeObjectURL(entry.blobUrl);
                        entry.blobUrl = URL.createObjectURL(blob);
                    } catch (_) {}

                    const AudioCtx = window.AudioContext || window.webkitAudioContext;
                    if (!AudioCtx) {
                        entry.status = 'ready';
                        return entry;
                    }
                    const ctx = new AudioCtx();
                    return ctx.decodeAudioData(buf.slice(0)).then(decoded => {
                        entry.buffer = decoded;
                        entry.status = 'ready';
                        return ctx.close().catch(() => {});
                    }).catch(() => {
                        entry.status = 'ready';
                        return null;
                    }).finally(() => {
                        ctx.close().catch(() => {});
                    });
                });
        })
        .then(() => {
            entry.status = 'ready';
            return entry;
        })
        .catch(() => {
            entry.status = 'failed';
            return entry;
        })
        .finally(() => {
            entry.promise = null;
        });

    return entry.promise;
}

function setupAudioPreloadObserver() {
    if (!appSettings.audioPreloadEnabled) {
        if (audioPreloadObserver) {
            audioPreloadObserver.disconnect();
            audioPreloadObserver = null;
        }
        return;
    }

    if (audioPreloadObserver) audioPreloadObserver.disconnect();

    audioPreloadObserver = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const word = entry.target.dataset.word;
                if (word) queueWordAudioPreload(word);
            }
        });
    }, {
        root: null,
        rootMargin: '200px 0px', // preload slightly before scrolling into view
        threshold: 0
    });

    document.querySelectorAll('.word-item[data-word]').forEach(el => {
        audioPreloadObserver.observe(el);
    });
}

function applyAudioPreloadSetting() {
    if (appSettings.audioPreloadEnabled) {
        setupAudioPreloadObserver();
        scheduleGlobalWordAudioPreload();
        return;
    }

    if (audioPreloadObserver) {
        audioPreloadObserver.disconnect();
        audioPreloadObserver = null;
    }
}

function playActionSound(actionType) {
    if (!appSettings.actionSoundEnabled) return;
    if (actionType !== 'put' && actionType !== 'delete') return;

    const isPut = actionType === 'put';
    let sound = isPut ? putActionSoundAudio : deleteActionSoundAudio;
    if (!sound) {
        const audioPath = isPut ? 'assets/put.mp3' : 'assets/delete.mp3';
        sound = new Audio(getPreloadedAssetURL(audioPath));
        sound.preload = 'auto';
        if (isPut) {
            putActionSoundAudio = sound;
        } else {
            deleteActionSoundAudio = sound;
        }
    }

    try {
        sound.currentTime = 0;
        const playPromise = sound.play();
        if (playPromise && typeof playPromise.catch === 'function') {
            playPromise.catch(() => {});
        }
    } catch (error) {
    }
}

const SPEAKER_ICON_SVG = '<svg class="icon-speaker" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path fill="currentColor" d="M3 10v4c0 1.1.9 2 2 2h2.35l3.38 2.7c.93.74 2.27.08 2.27-1.1V6.4c0-1.18-1.34-1.84-2.27-1.1L7.35 8H5c-1.1 0-2 .9-2 2Zm14.5 2c0-1.77-1.02-3.29-2.5-4.03v8.06c1.48-.74 2.5-2.26 2.5-4.03ZM15 3.23v2.06c2.89.86 5 3.54 5 6.71s-2.11 5.85-5 6.71v2.06c4.01-.91 7-4.49 7-8.77s-2.99-7.86-7-8.77Z"/></svg>';
const DATE_INPUT_ICON_ONLY_THRESHOLD = 120;
let dateInputResizeObserver = null;
const LANDSCAPE_LAYOUT_MIN_WIDTH = 768;
const PORTRAIT_LAYOUT_MAX_WIDTH = LANDSCAPE_LAYOUT_MIN_WIDTH - 1;
const LAYOUT_BREAKPOINT_QUERY = `(min-width: ${LANDSCAPE_LAYOUT_MIN_WIDTH}px)`;
const PORTRAIT_MOBILE_QUERY = `(max-width: ${PORTRAIT_LAYOUT_MAX_WIDTH}px) and (orientation: portrait)`;
const LAYOUT_SPLIT_STORAGE_KEY = 'wordMemoryLayoutSplitRatio';
const LAYOUT_MIN_PANEL_WIDTH = 320;
const LAYOUT_KEYBOARD_STEP = 0.02;
let layoutSplitRatio = 0.5;

// Dropdown Base Class
class Dropdown {
    constructor(containerId, selectedId, dropdownId) {
        this.container = document.getElementById(containerId);
        this.selected = document.getElementById(selectedId);
        this.dropdown = document.getElementById(dropdownId);
        this.isOpen = false;

        // Check if all elements exist
        if (!this.container || !this.selected || !this.dropdown) {
            return;
        }

        if (this.selected) {
            this.selected.addEventListener('click', (e) => {
                e.stopPropagation();
                this.toggle();
            });
        }
    }

    toggle() {
        if (!this.dropdown || !this.container) return;

        this.isOpen = !this.isOpen;
        this.dropdown.classList.toggle('active', this.isOpen);
        this.container.classList.toggle('open', this.isOpen);

        if (this.isOpen) {
            Dropdown.closeAll(this);
        }
    }

    close() {
        if (!this.dropdown || !this.container) return;

        if (this.isOpen) {
            this.isOpen = false;
            this.dropdown.classList.remove('active');
            this.container.classList.remove('open');
        }
    }

    static instances = [];

    static register(instance) {
        if (instance.container && instance.selected && instance.dropdown) {
            Dropdown.instances.push(instance);
        }
    }

    static closeAll(except = null) {
        Dropdown.instances.forEach(instance => {
            if (instance !== except) {
                instance.close();
            }
        });
    }
}

function getWeightLabel(weight) {
    const name = weight === -3 ? 'Invalid' :
                 weight >= 5 ? 'Hardest' :
                 weight === 4 ? 'Hard' :
                 weight === 3 ? 'Memorise' :
                 weight === 2 ? 'Normal' :
                 weight === 1 ? 'Recognise' :
                 weight === 0 ? 'Basic' :
                 weight === -1 ? 'Mastered' :
                 'Easy';

    return name;
}

// Status indicator
function showStatus(message, type = 'info') {
    const indicator = document.getElementById('statusIndicator');
    indicator.textContent = message;
    indicator.className = 'status-indicator';
    if (type === 'success') indicator.classList.add('success');
    if (type === 'error') indicator.classList.add('error');

    setTimeout(() => {
        indicator.textContent = '';
        indicator.className = 'status-indicator';
    }, 3000);
}

function ensureDateCompactLabel(input) {
    if (!input || !input.parentElement) return null;

    let wrapper = input.parentElement;
    if (!wrapper.classList.contains('range-input-date-wrap')) {
        wrapper = document.createElement('div');
        wrapper.className = 'range-input-date-wrap';
        input.parentElement.insertBefore(wrapper, input);
        wrapper.appendChild(input);
    }

    let label = wrapper.querySelector('.range-input-date-compact-label');
    if (!label) {
        label = document.createElement('span');
        label.className = 'range-input-date-compact-label';
        label.setAttribute('aria-hidden', 'true');
        wrapper.appendChild(label);
    }

    return label;
}

function parseIsoDateParts(isoDate) {
    if (!isoDate) return null;

    const parts = isoDate.split('-');
    if (parts.length !== 3) return null;

    const year = Number(parts[0]);
    const month = Number(parts[1]);
    const day = Number(parts[2]);
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
        return null;
    }

    return { year, month, day };
}

function getRelativeDayOffset(isoDate) {
    const parsed = parseIsoDateParts(isoDate);
    if (!parsed) return null;

    const targetDate = new Date(parsed.year, parsed.month - 1, parsed.day);
    const today = new Date();
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const targetStart = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
    return Math.round((targetStart.getTime() - todayStart.getTime()) / 86400000);
}

function getSpecialDateLabel(isoDate, shortLabel = false) {
    const diffDays = getRelativeDayOffset(isoDate);
    if (diffDays === null) return '';

    if (shortLabel) {
        if (diffDays === 0) return 'TOD';
        if (diffDays === 1) return 'TMR';
        if (diffDays === -1) return 'YDA';
        return '';
    }

    if (diffDays === 0) return 'Today';
    if (diffDays === 1) return 'Tomorrow';
    if (diffDays === -1) return 'Yesterday';
    return '';
}

function formatDateByRules(isoDate) {
    const parsed = parseIsoDateParts(isoDate);
    if (!parsed) return isoDate || '';

    const currentYear = new Date().getFullYear();
    if (parsed.year === currentYear) {
        return `${parsed.month}-${parsed.day}`;
    }

    const sameCentury = Math.floor(parsed.year / 100) === Math.floor(currentYear / 100);
    if (sameCentury) {
        return `${parsed.year % 100}-${parsed.month}-${parsed.day}`;
    }

    return `${parsed.year}-${parsed.month}-${parsed.day}`;
}

function formatCompactDateLabel(isoDate) {
    if (!isoDate) return '';

    const specialShort = getSpecialDateLabel(isoDate, true);
    if (specialShort) return specialShort;

    return formatDateByRules(isoDate);
}

function formatAddedDateLabel(isoDate) {
    if (!isoDate) return '';

    const specialLabel = getSpecialDateLabel(isoDate, false);
    if (specialLabel) return specialLabel;

    return formatDateByRules(isoDate);
}

function syncDateInputDisplayMode() {
    const dateInputs = document.querySelectorAll('.range-input-date');

    dateInputs.forEach((input) => {
        const label = ensureDateCompactLabel(input);
        const isNarrow = input.getBoundingClientRect().width <= DATE_INPUT_ICON_ONLY_THRESHOLD;
        const hasValue = Boolean(input.value);
        const isIconOnly = isNarrow && !hasValue;
        const isCompactValue = isNarrow && hasValue;
        const specialFullLabel = !isNarrow && hasValue ? getSpecialDateLabel(input.value, false) : '';
        const isSpecialFull = Boolean(specialFullLabel);

        input.classList.toggle('date-icon-only', isIconOnly);
        input.classList.toggle('date-compact-value', isCompactValue);
        input.classList.toggle('date-special-full', isSpecialFull);

        if (label) {
            label.textContent = isCompactValue ? formatCompactDateLabel(input.value) : (isSpecialFull ? specialFullLabel : '');
            label.classList.toggle('active', isCompactValue || isSpecialFull);
        }
    });
}

function initResponsiveDateInputs() {
    syncDateInputDisplayMode();

    const dateInputs = document.querySelectorAll('.range-input-date');
    if (!dateInputs.length) return;

    dateInputs.forEach((input) => {
        const label = ensureDateCompactLabel(input);

        if (input.dataset.iconOnlyPickerBound === 'true') return;

        // Make overlay label click through to the input on Safari
        if (label) {
            label.addEventListener('click', () => {
                input.focus();
                if (typeof input.showPicker === 'function') {
                    try { input.showPicker(); } catch(e) {}
                }
            });
            label.style.pointerEvents = 'auto';
            label.style.cursor = 'pointer';
        }

        input.addEventListener('click', (event) => {
            const isCustomDisplayMode = input.classList.contains('date-icon-only') ||
                input.classList.contains('date-compact-value') ||
                input.classList.contains('date-special-full');
            if (!isCustomDisplayMode) return;

            if (typeof input.showPicker === 'function') {
                try {
                    event.preventDefault();
                    input.showPicker();
                } catch(e) {
                    input.focus();
                }
            } else {
                input.focus();
            }
        });

        input.addEventListener('change', syncDateInputDisplayMode);
        input.addEventListener('input', syncDateInputDisplayMode);
        input.dataset.iconOnlyPickerBound = 'true';
    });

    if (typeof ResizeObserver !== 'undefined') {
        if (dateInputResizeObserver) {
            dateInputResizeObserver.disconnect();
        }

        dateInputResizeObserver = new ResizeObserver(() => {
            syncDateInputDisplayMode();
        });
        dateInputs.forEach((input) => dateInputResizeObserver.observe(input));
    } else {
        window.addEventListener('resize', syncDateInputDisplayMode);
    }
}

function clampLayoutValue(value, min, max) {
    return Math.min(Math.max(value, min), max);
}

function loadLayoutSplitRatio() {
    try {
        const raw = localStorage.getItem(LAYOUT_SPLIT_STORAGE_KEY);
        const parsed = raw ? Number(raw) : NaN;
        if (!Number.isFinite(parsed)) {
            return 0.5;
        }
        return clampLayoutValue(parsed, 0, 1);
    } catch (error) {
        return 0.5;
    }
}

function saveLayoutSplitRatio(ratio) {
    try {
        localStorage.setItem(LAYOUT_SPLIT_STORAGE_KEY, String(clampLayoutValue(ratio, 0, 1)));
    } catch (error) {
    }
}

function setLayoutSplitRatio(layout, divider, ratio) {
    if (!layout || !divider) return;

    const totalWidth = layout.getBoundingClientRect().width;
    const dividerWidth = divider.getBoundingClientRect().width || 12;
    const availableWidth = Math.max(0, totalWidth - dividerWidth);
    if (availableWidth <= 0) return;

    let minLeft = LAYOUT_MIN_PANEL_WIDTH;
    let maxLeft = availableWidth - LAYOUT_MIN_PANEL_WIDTH;

    // If viewport is too narrow to keep both panes at min width, fall back to 50/50.
    if (maxLeft < minLeft) {
        minLeft = availableWidth / 2;
        maxLeft = availableWidth / 2;
    }

    const clampedRatio = clampLayoutValue(ratio, 0, 1);
    const desiredLeft = availableWidth * clampedRatio;
    const leftWidth = clampLayoutValue(desiredLeft, minLeft, maxLeft);
    const normalizedRatio = availableWidth > 0 ? (leftWidth / availableWidth) : 0.5;

    layout.style.setProperty('--workspace-left-width', `${leftWidth}px`);
    layoutSplitRatio = normalizedRatio;
    divider.setAttribute('aria-valuenow', String(Math.round(normalizedRatio * 100)));
}

function initResizableLayout() {
    const layout = document.getElementById('workspaceLayout');
    const divider = document.getElementById('layoutDivider');
    if (!layout || !divider) return;

    layoutSplitRatio = loadLayoutSplitRatio();
    const mediaQuery = window.matchMedia(LAYOUT_BREAKPOINT_QUERY);
    let isResizing = false;

    const applyLayout = () => {
        if (!mediaQuery.matches) {
            layout.style.removeProperty('--workspace-left-width');
            divider.setAttribute('aria-valuenow', String(Math.round(layoutSplitRatio * 100)));
            return;
        }

        setLayoutSplitRatio(layout, divider, layoutSplitRatio);
        syncDateInputDisplayMode();
    };

    const handlePointerMove = (event) => {
        if (!isResizing) return;
        event.preventDefault();

        const rect = layout.getBoundingClientRect();
        const dividerWidth = divider.getBoundingClientRect().width || 12;
        const availableWidth = Math.max(0, rect.width - dividerWidth);
        if (availableWidth <= 0) return;

        const rawLeft = event.clientX - rect.left;
        const nextRatio = rawLeft / availableWidth;
        setLayoutSplitRatio(layout, divider, nextRatio);
        syncDateInputDisplayMode();
    };

    const stopResize = () => {
        if (!isResizing) return;
        isResizing = false;
        divider.classList.remove('is-dragging');
        document.body.classList.remove('is-resizing-layout');
        saveLayoutSplitRatio(layoutSplitRatio);
        window.removeEventListener('pointermove', handlePointerMove);
        window.removeEventListener('pointerup', stopResize);
        window.removeEventListener('pointercancel', stopResize);
    };

    divider.addEventListener('pointerdown', (event) => {
        if (!mediaQuery.matches) return;

        event.preventDefault();
        isResizing = true;
        divider.classList.add('is-dragging');
        document.body.classList.add('is-resizing-layout');

        if (typeof divider.setPointerCapture === 'function') {
            try {
                divider.setPointerCapture(event.pointerId);
            } catch (error) {
                // Some environments can reject pointer capture; dragging still works via window listeners.
            }
        }

        window.addEventListener('pointermove', handlePointerMove);
        window.addEventListener('pointerup', stopResize);
        window.addEventListener('pointercancel', stopResize);
    });

    divider.addEventListener('keydown', (event) => {
        if (!mediaQuery.matches) return;

        if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') {
            return;
        }

        event.preventDefault();
        const delta = event.key === 'ArrowLeft' ? -LAYOUT_KEYBOARD_STEP : LAYOUT_KEYBOARD_STEP;
        setLayoutSplitRatio(layout, divider, layoutSplitRatio + delta);
        saveLayoutSplitRatio(layoutSplitRatio);
        syncDateInputDisplayMode();
    });

    divider.addEventListener('dblclick', (event) => {
        event.preventDefault();
        layoutSplitRatio = 0.5;
        setLayoutSplitRatio(layout, divider, 0.5);
        saveLayoutSplitRatio(0.5);
        syncDateInputDisplayMode();
    });

    window.addEventListener('resize', applyLayout);
    if (typeof mediaQuery.addEventListener === 'function') {
        mediaQuery.addEventListener('change', applyLayout);
    } else if (typeof mediaQuery.addListener === 'function') {
        mediaQuery.addListener(applyLayout);
    }

    applyLayout();
}

function selectRegistryPreviewContent() {
    const preview = document.getElementById('registryPreview');
    if (!preview) return false;

    const selectableNodes = preview.querySelectorAll('.json-line, .diff-header, .registry-preview-empty');
    if (!selectableNodes.length) return false;

    const selection = window.getSelection();
    if (!selection) return false;

    const firstNode = selectableNodes[0];
    const lastNode = selectableNodes[selectableNodes.length - 1];
    const range = document.createRange();
    range.setStart(firstNode, 0);
    range.setEnd(lastNode, lastNode.childNodes.length);
    selection.removeAllRanges();
    selection.addRange(range);
    return true;
}

function shouldHandleRegistryPreviewSelectAll(event, isInputFocused) {
    if (isInputFocused) return false;
    if (!(event.metaKey || event.ctrlKey)) return false;
    if (event.altKey || event.shiftKey) return false;
    if (event.key.toLowerCase() !== 'a') return false;

    const historyModal = document.getElementById('historyModal');
    const preview = document.getElementById('registryPreview');
    if (!historyModal || !preview || !historyModal.classList.contains('active')) {
        return false;
    }

    const targetNode = event.target instanceof Node ? event.target : null;
    if (targetNode && preview.contains(targetNode)) {
        return true;
    }

    const selection = window.getSelection();
    const anchorNode = selection ? selection.anchorNode : null;
    return Boolean(anchorNode && preview.contains(anchorNode));
}

function resetAddAndBatchToolbarInputs() {
    const today = new Date().toISOString().split('T')[0];

    const wordInput = document.getElementById('wordInput');
    if (wordInput) wordInput.value = '';

    const meaningInput = document.getElementById('meaningInput');
    if (meaningInput) meaningInput.value = '';

    tagSelectorState.add = [];
    renderTagSelector('add');

    selectedPos = [];
    updatePosSelection();

    selectedWeight = 3;
    updateWeightSelection();

    const dateInput = document.getElementById('dateInput');
    if (dateInput) dateInput.value = today;

    const weightMinInput = document.getElementById('weightMinInput');
    if (weightMinInput) weightMinInput.value = '-3';

    const weightMaxInput = document.getElementById('weightMaxInput');
    if (weightMaxInput) weightMaxInput.value = '';

    const dateMinInput = document.getElementById('dateMinInput');
    if (dateMinInput) dateMinInput.value = '';

    const dateMaxInput = document.getElementById('dateMaxInput');
    if (dateMaxInput) dateMaxInput.value = '';

    const regexFilterInput = document.getElementById('regexFilterInput');
    if (regexFilterInput) regexFilterInput.value = '';

    const batchDateInput = document.getElementById('batchDateInput');
    if (batchDateInput) batchDateInput.value = today;
}

// Initialize
window.onload = function() {
    loadData();
    loadTagRegistry();
    migrateStringTagsToRegistry();
    appSettings = loadAppSettings();
    warmCoreAssets().finally(() => {
        _initActionSounds();
        if (_dingAudio) {
            _dingAudio.src = getPreloadedAssetURL('assets/ding.mp3');
            _dingAudio.preload = 'auto';
        }
    });

    // Initialize version control
    versionControl = new VersionControl(50);

    // Create initial version if no versions exist
    if (versionControl.versions.size === 0 && words.length > 0) {
        versionControl.createVersion(words, 'Initial version');
    } else if (versionControl.versions.size === 0 && words.length === 0) {
        versionControl.createVersion([], 'Initial empty state');
    }
    updateHistoryButtonLabel();

    _initActionSounds();

    renderWords();
    resetAddAndBatchToolbarInputs();
    updateEditWeightSelection();

    // Load voices for speech synthesis — cache for iOS which loads them async
    if ('speechSynthesis' in window) {
        cachedVoices = speechSynthesis.getVoices();
        speechSynthesis.onvoiceschanged = () => {
            cachedVoices = speechSynthesis.getVoices();
        };
    }

    // Initialize dropdown instances
    posDropdown = new Dropdown('posSelector', 'posSelected', 'posDropdown');
    weightDropdown = new Dropdown('weightSelector', 'weightSelected', 'weightDropdown');
    editPosDropdown = new Dropdown('editPosSelector', 'editPosSelected', 'editPosDropdown');
    editWeightDropdown = new Dropdown('editWeightSelector', 'editWeightSelected', 'editWeightDropdown');

    tagDropdownInstance = new Dropdown('tagSelector', 'tagSelected', 'tagDropdown');
    editTagDropdownInstance = new Dropdown('editTagSelector', 'editTagSelected', 'editTagDropdown');
    batchTagFilterDropdownInstance = new Dropdown('batchTagFilterSelector', 'batchTagFilterSelected', 'batchTagFilterDropdown');
    batchTagActionDropdownInstance = new Dropdown('batchTagActionSelector', 'batchTagActionSelected', 'batchTagActionDropdown');

    // Refresh tag options when dropdown opens
    const tagSelectedEl = document.getElementById('tagSelected');
    if (tagSelectedEl) {
        tagSelectedEl.addEventListener('click', () => renderTagSelector('add'));
    }
    const editTagSelectedEl = document.getElementById('editTagSelected');
    if (editTagSelectedEl) {
        editTagSelectedEl.addEventListener('click', () => renderTagSelector('edit'));
    }
    const batchTagFilterSelectedEl = document.getElementById('batchTagFilterSelected');
    if (batchTagFilterSelectedEl) {
        batchTagFilterSelectedEl.addEventListener('click', () => renderTagSelector('batchFilter'));
    }
    const batchTagActionSelectedEl = document.getElementById('batchTagActionSelected');
    if (batchTagActionSelectedEl) {
        batchTagActionSelectedEl.addEventListener('click', () => renderTagSelector('batchAction'));
    }

    Dropdown.register(posDropdown);
    Dropdown.register(weightDropdown);
    Dropdown.register(editPosDropdown);
    Dropdown.register(editWeightDropdown);
    Dropdown.register(tagDropdownInstance);
    Dropdown.register(editTagDropdownInstance);
    Dropdown.register(batchTagFilterDropdownInstance);
    Dropdown.register(batchTagActionDropdownInstance);

    // Initialize batch toolbar
    updateBatchToolbar();
    initResponsiveDateInputs();
    initResizableLayout();
    initDragAndDropImport();
    updatePortraitModeSwitchTop();
};

// Mode toggle
const modeToggle = document.getElementById('modeToggle');
const modeSwitch = document.querySelector('.mode-switch');
const pageTitle = document.querySelector('h1');

let modeSwitchPlaceholder = null;
let modeSwitchNaturalTop = null;

function measureModeSwitchNaturalTop() {
    if (!modeSwitch) return;
    const wasFixed = modeSwitch.classList.contains('is-fixed');
    if (wasFixed) {
        modeSwitch.classList.remove('is-fixed');
        if (modeSwitchPlaceholder) modeSwitchPlaceholder.style.display = 'none';
    }
    modeSwitchNaturalTop = modeSwitch.getBoundingClientRect().top + window.scrollY;
    if (wasFixed) {
        modeSwitch.classList.add('is-fixed');
        if (modeSwitchPlaceholder) modeSwitchPlaceholder.style.display = '';
    }
}

function updatePortraitModeSwitchTop() {
    if (!modeSwitch || !pageTitle) return;

    if (!window.matchMedia(PORTRAIT_MOBILE_QUERY).matches) {
        modeSwitch.classList.remove('is-fixed');
        if (modeSwitchPlaceholder) modeSwitchPlaceholder.style.display = 'none';
        modeSwitchNaturalTop = null;
        return;
    }

    if (modeSwitchNaturalTop === null) measureModeSwitchNaturalTop();

    if (window.scrollY >= modeSwitchNaturalTop) {
        if (!modeSwitch.classList.contains('is-fixed')) {
            // Create placeholder to prevent layout shift
            if (!modeSwitchPlaceholder) {
                modeSwitchPlaceholder = document.createElement('div');
                modeSwitch.parentNode.insertBefore(modeSwitchPlaceholder, modeSwitch.nextSibling);
            }
            modeSwitchPlaceholder.style.height = modeSwitch.offsetHeight + 'px';
            modeSwitchPlaceholder.style.display = '';
            modeSwitch.classList.add('is-fixed');
        }
    } else {
        if (modeSwitch.classList.contains('is-fixed')) {
            modeSwitch.classList.remove('is-fixed');
            if (modeSwitchPlaceholder) modeSwitchPlaceholder.style.display = 'none';
        }
    }
}

window.addEventListener('scroll', updatePortraitModeSwitchTop, { passive: true });
window.addEventListener('resize', () => { modeSwitchNaturalTop = null; updatePortraitModeSwitchTop(); });

// Run fn() while preserving the scroll anchor of the first visible word item
function withScrollAnchor(fn) {
    const sc = document.querySelector('.workspace-right');
    let anchorWord = null, anchorOffset = 0;
    if (sc) {
        const top = sc.getBoundingClientRect().top;
        for (const item of sc.querySelectorAll('.word-item')) {
            const rect = item.getBoundingClientRect();
            if (rect.bottom > top) {
                anchorWord = item.getAttribute('data-word');
                anchorOffset = rect.top - top;
                break;
            }
        }
    }
    fn();
    if (sc && anchorWord) {
        const target = sc.querySelector(`.word-item[data-word="${CSS.escape(anchorWord)}"]`);
        if (target) {
            sc.scrollTop += target.getBoundingClientRect().top - sc.getBoundingClientRect().top - anchorOffset;
        }
    }
}

function toggleMode() {
    modeToggle.classList.toggle('active');
    hideMeaning = !hideMeaning;
    withScrollAnchor(() => renderWords());
}

modeToggle.addEventListener('click', toggleMode);

modeToggle.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        toggleMode();
    }
});

// POS selection for add form
let selectedPos = [];
let selectedWeight = 3;
let editSelectedWeight = 3;

// Dropdown instances
let posDropdown = null;
let weightDropdown = null;
let editPosDropdown = null;
let editWeightDropdown = null;

function togglePosDropdown() { if (posDropdown) posDropdown.toggle(); }

function renderPosSelection(containerId, values) {
    const container = document.getElementById(containerId);
    container.innerHTML = values.length === 0
        ? '<span class="pos-placeholder">POS</span>'
        : values.map(p => `<span class="pos-tag">${p}</span>`).join('');
}

function syncPosOptionState(dropdownId, values) {
    const selectedValues = new Set(values);
    document.querySelectorAll(`#${dropdownId} .pos-option`).forEach(option => {
        option.classList.toggle('selected', selectedValues.has(option.dataset.value));
    });
}

function _togglePosFor(getArr, setArr, updateFn, option, event) {
    if (event) event.stopPropagation();
    const value = option.dataset.value;
    if (!value) return;
    const arr = getArr();
    setArr(arr.includes(value) ? arr.filter(p => p !== value) : [...arr, value]);
    updateFn();
}

function updatePosSelection() {
    renderPosSelection('posSelected', selectedPos);
    syncPosOptionState('posDropdown', selectedPos);
}

function togglePosOption(option, event) {
    _togglePosFor(() => selectedPos, v => { selectedPos = v; }, updatePosSelection, option, event);
}

// POS selection for edit form
let editSelectedPos = [];

function toggleEditPosDropdown() { if (editPosDropdown) editPosDropdown.toggle(); }

function updateEditPosSelection() {
    renderPosSelection('editPosSelected', editSelectedPos);
    syncPosOptionState('editPosDropdown', editSelectedPos);
}

function toggleEditPosOption(option, event) {
    _togglePosFor(() => editSelectedPos, v => { editSelectedPos = v; }, updateEditPosSelection, option, event);
}

function syncWeightOptionState(dropdownId, value) {
    document.querySelectorAll(`#${dropdownId} .weight-option`).forEach(option => {
        option.classList.toggle('selected', parseInt(option.dataset.value, 10) === value);
    });
}

function _updateWeightSelectionImpl(selectedId, dropdownId, value) {
    document.getElementById(selectedId).textContent = String(value);
    syncWeightOptionState(dropdownId, value);
}

function updateWeightSelection() { _updateWeightSelectionImpl('weightSelected', 'weightDropdown', selectedWeight); }
function updateEditWeightSelection() { _updateWeightSelectionImpl('editWeightSelected', 'editWeightDropdown', editSelectedWeight); }

function toggleWeightDropdown() { if (weightDropdown) weightDropdown.toggle(); }
function toggleEditWeightDropdown() { if (editWeightDropdown) editWeightDropdown.toggle(); }

function _setWeightFor(setValue, updateFn, dropdown, option, event) {
    if (event) event.stopPropagation();
    const value = parseInt(option.dataset.value, 10);
    if (Number.isNaN(value)) return;
    setValue(value);
    updateFn();
    if (dropdown) dropdown.close();
}

function setWeightOption(option, event) {
    _setWeightFor(v => { selectedWeight = v; }, updateWeightSelection, weightDropdown, option, event);
}

function setEditWeightOption(option, event) {
    _setWeightFor(v => { editSelectedWeight = v; }, updateEditWeightSelection, editWeightDropdown, option, event);
}

// ========================================
// Tag Registry
// ========================================

let tagRegistry = []; // Array of { id, name }
let tagDropdownInstance = null;
let editTagDropdownInstance = null;
let batchTagFilterDropdownInstance = null;
let batchTagActionDropdownInstance = null;

const TAG_REGISTRY_KEY = 'wordMemoryTagRegistry';

function generateTagId() {
    return 't_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function saveTagRegistry() {
    localStorage.setItem(TAG_REGISTRY_KEY, JSON.stringify(tagRegistry));
}

function loadTagRegistry() {
    try {
        const saved = localStorage.getItem(TAG_REGISTRY_KEY);
        if (saved) tagRegistry = JSON.parse(saved);
    } catch (e) {
        tagRegistry = [];
    }
}

function getTagName(id) {
    const entry = tagRegistry.find(t => t.id === id);
    return entry ? entry.name : '';
}

function getTagId(name) {
    const lower = name.toLowerCase();
    const entry = tagRegistry.find(t => t.name.toLowerCase() === lower);
    return entry ? entry.id : null;
}

function createTag(name) {
    const trimmed = name.trim();
    if (!trimmed) return null;
    // Check if already exists (case-insensitive)
    const existing = getTagId(trimmed);
    if (existing) return existing;
    const id = generateTagId();
    tagRegistry.push({ id, name: trimmed });
    saveTagRegistry();
    return id;
}

function renameTag(id, newName) {
    const trimmed = newName.trim();
    if (!trimmed) return false;
    const entry = tagRegistry.find(t => t.id === id);
    if (!entry) return false;
    entry.name = trimmed;
    saveTagRegistry();
    return true;
}

function deleteTag(id) {
    tagRegistry = tagRegistry.filter(t => t.id !== id);
    // Remove from all words
    words.forEach(w => {
        if (w.tags) w.tags = w.tags.filter(t => t !== id);
    });
    // Remove from active/excluded filters
    activeTagFilters.delete(id);
    excludedTagFilters.delete(id);
    saveTagRegistry();
    saveData(false, `Delete tag`);
}

// Migrate old string-based tags to registry IDs
function migrateStringTagsToRegistry() {
    let migrated = false;
    words.forEach(w => {
        if (!w.tags || w.tags.length === 0) return;
        w.tags = w.tags.map(t => {
            // If it already looks like an ID (starts with t_), skip
            if (typeof t === 'string' && t.startsWith('t_')) return t;
            // It's an old string tag — find or create registry entry
            migrated = true;
            const existingId = getTagId(t);
            if (existingId) return existingId;
            const id = generateTagId();
            tagRegistry.push({ id, name: t });
            return id;
        });
    });
    if (migrated) {
        saveTagRegistry();
        saveData(false);
    }
}

// ========================================
// Unified Tag Selector
// ========================================

const TAG_SELECTORS = {
    add:         { selectedId: 'tagSelected', listId: 'tagOptionsList', inputId: 'tagNewInput', multi: true, placeholder: 'Tags' },
    edit:        { selectedId: 'editTagSelected', listId: 'editTagOptionsList', inputId: 'editTagNewInput', multi: true, placeholder: 'Tags' },
    batchFilter: { selectedId: 'batchTagFilterSelected', listId: 'batchTagFilterOptionsList', inputId: null, multi: false, placeholder: 'Tag' },
    batchAction: { selectedId: 'batchTagActionSelected', listId: 'batchTagActionOptionsList', inputId: 'batchTagActionNewInput', multi: true, placeholder: 'Tag' },
};

let tagSelectorState = {
    add: [],         // tag ID array
    edit: [],        // tag ID array
    batchFilter: '', // single tag ID
    batchAction: [], // tag ID array
};

function getTagDropdownInstance(key) {
    return { add: tagDropdownInstance, edit: editTagDropdownInstance, batchFilter: batchTagFilterDropdownInstance, batchAction: batchTagActionDropdownInstance }[key];
}

function renderTagSelector(key) {
    const cfg = TAG_SELECTORS[key];
    if (!cfg) return;
    const state = tagSelectorState[key];

    // Render selected display
    const selectedEl = document.getElementById(cfg.selectedId);
    if (selectedEl) {
        if (cfg.multi) {
            const ids = state || [];
            if (ids.length === 0) {
                selectedEl.innerHTML = `<span class="tag-placeholder">${cfg.placeholder}</span>`;
            } else {
                selectedEl.innerHTML = ids.map(id => `<span class="tag-sel-tag">${getTagName(id)}</span>`).join('');
            }
        } else {
            const id = state;
            if (id) {
                selectedEl.innerHTML = `<span class="tag-sel-tag">${getTagName(id)}</span>`;
            } else {
                selectedEl.innerHTML = `<span class="tag-placeholder">${cfg.placeholder}</span>`;
            }
        }
    }

    // Render options list
    const listEl = document.getElementById(cfg.listId);
    if (listEl) {
        const selectedSet = cfg.multi ? new Set(state || []) : new Set(state ? [state] : []);

        // For single-select optional filters (e.g. quiz), prepend an "All" clear option
        let clearOptionHtml = '';
        if (!cfg.multi && cfg.placeholder) {
            const clearSelected = !state ? 'selected' : '';
            clearOptionHtml = `<label class="dropdown-option tag-option ${clearSelected}" data-value="">${cfg.placeholder}</label>`;
        }

        listEl.innerHTML = clearOptionHtml + tagRegistry.map(tag => {
            const isSelected = selectedSet.has(tag.id) ? 'selected' : '';
            return `<label class="dropdown-option tag-option ${isSelected}" data-value="${tag.id}">${tag.name}</label>`;
        }).join('');

        // Bind click handlers
        listEl.querySelectorAll('.tag-option').forEach(option => {
            option.onclick = function(e) {
                e.stopPropagation();
                const tagId = this.dataset.value;
                if (cfg.multi) {
                    toggleTagInSelector(key, tagId);
                } else {
                    setTagInSelector(key, tagId);
                }
            };
        });
    }
}

function toggleTagInSelector(key, tagId) {
    const arr = tagSelectorState[key];
    const idx = arr.indexOf(tagId);
    if (idx !== -1) {
        arr.splice(idx, 1);
    } else {
        arr.push(tagId);
    }
    renderTagSelector(key);
}

function setTagInSelector(key, tagId) {
    tagSelectorState[key] = tagId;
    renderTagSelector(key);
    const dropdown = getTagDropdownInstance(key);
    if (dropdown) dropdown.close();
}

function addNewTagToSelector(key, event) {
    if (event) event.stopPropagation();
    const cfg = TAG_SELECTORS[key];
    if (!cfg || !cfg.inputId) return;
    const input = document.getElementById(cfg.inputId);
    if (!input) return;
    const raw = input.value.trim();
    if (!raw) return;
    const names = raw.split(/[,，]/).map(t => t.trim()).filter(t => t.length > 0);
    names.forEach(name => {
        const id = createTag(name);
        if (!id) return;
        if (cfg.multi) {
            if (!tagSelectorState[key].includes(id)) tagSelectorState[key].push(id);
        } else {
            tagSelectorState[key] = id;
        }
    });
    input.value = '';
    renderTagSelector(key);
    if (!cfg.multi) {
        const dropdown = getTagDropdownInstance(key);
        if (dropdown) dropdown.close();
    }
}

// Global function wrappers for onclick in HTML
function addNewTagFromInput(event) { addNewTagToSelector('add', event); }
function addNewEditTagFromInput(event) { addNewTagToSelector('edit', event); }
function addBatchTagActionFromInput(event) { addNewTagToSelector('batchAction', event); }

// Handle Enter key in all tag inputs
document.addEventListener('DOMContentLoaded', () => {
    Object.entries(TAG_SELECTORS).forEach(([key, cfg]) => {
        if (!cfg.inputId) return;
        const input = document.getElementById(cfg.inputId);
        if (input) {
            input.addEventListener('keydown', function(e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    e.stopPropagation();
                    addNewTagToSelector(key, e);
                }
            });
        }
    });
});

// ========================================
// Tag Filter Bar (filter displayed words)
// ========================================

function renderTagFilterBar() {
    const bar = document.getElementById('tagFilterBar');
    if (!bar) return;
    if (tagRegistry.length === 0) {
        bar.innerHTML = '';
        return;
    }
    bar.innerHTML = tagRegistry.map(tag => {
        const state = activeTagFilters.has(tag.id) ? 'active' : excludedTagFilters.has(tag.id) ? 'excluded' : '';
        return `<span class="tag-filter-chip ${state}" onclick="toggleTagFilter('${tag.id}')">${tag.name}</span>`;
    }).join('');
}

function toggleTagFilter(tag) {
    if (activeTagFilters.has(tag)) {
        activeTagFilters.delete(tag);
        excludedTagFilters.add(tag);
    } else if (excludedTagFilters.has(tag)) {
        excludedTagFilters.delete(tag);
    } else {
        activeTagFilters.add(tag);
    }
    renderTagFilterBar();
    renderWords();
}

function enforceGroupModeByTagAvailability() {
    if (tagRegistry.length === 0) {
        wordGroupMode = 'weight';
    }

    const btn = document.getElementById('groupModeBtn');
    if (btn) {
        btn.textContent = wordGroupMode === 'tag' ? 'Tag Grouping' : 'Weight Grouping';
    }
}

// ========================================
// Group Mode Toggle
// ========================================

function toggleGroupMode() {
    if (tagRegistry.length === 0) {
        wordGroupMode = 'weight';
        enforceGroupModeByTagAvailability();
        renderWords();
        return;
    }

    const modes = ['weight', 'tag'];
    const labels = ['By Weight', 'By Tag'];
    const idx = modes.indexOf(wordGroupMode);
    wordGroupMode = modes[(idx + 1) % modes.length];
    const btn = document.getElementById('groupModeBtn');
    if (btn) btn.textContent = labels[modes.indexOf(wordGroupMode)];
    renderWords();
}

// Close dropdowns when clicking outside
document.addEventListener('click', function(e) {
    // Check if any dropdown is open
    const hasOpenDropdown = Dropdown.instances.some(d => d.isOpen);

    if (hasOpenDropdown) {
        // Check if click is inside any dropdown container
        const clickedInsideDropdown = e.target.closest('.dropdown-container');

        if (!clickedInsideDropdown) {
            // Click is outside all dropdowns, close them and prevent other actions
            e.stopPropagation();
            e.preventDefault();
            Dropdown.closeAll();
        }
    }
}, true); // Use capture phase to handle this before other handlers

// Add word
function addWord() {
    const word = document.getElementById('wordInput').value.trim();
    const meaning = document.getElementById('meaningInput').value.trim();
    const weight = selectedWeight;
    const date = document.getElementById('dateInput').value || new Date().toISOString().split('T')[0];
    const tags = tagSelectorState.add.slice();

    if (!word) {
        alert('Please fill in word');
        return;
    }

    const newWord = {
        word: word.toLowerCase(),
        meaning: meaning,
        pos: selectedPos.slice(), // Copy array
        weight: weight,
        added: date,
        joinedAt: new Date().toISOString().slice(0, 19),
        tags: tags
    };

    words.push(newWord);
    saveData(false, `＋　${word}`);
    playActionSound('put');
    renderWords();
    showStatus(`Added "${word}"`, 'success');

    // Clear form
    document.getElementById('wordInput').value = '';
    document.getElementById('meaningInput').value = '';
    tagSelectorState.add = [];
    renderTagSelector('add');
    selectedPos = [];
    updatePosSelection();
}

// Update weight
function updateWeight(index, delta) {
    const currentWeight = words[index].weight;
    const wordName = words[index].word;

    // If current weight is -3 (invalid), delta is the new weight directly
    if (currentWeight === -3) {
        words[index].weight = delta;
        saveData(false, `＃　${wordName}`);
        renderWords();
        showStatus(`Fixed weight to ${delta}`, 'success');
        return;
    }

    const newWeight = currentWeight + delta;

    if (newWeight < -2 || newWeight > 10) {
        return;
    }

    words[index].weight = newWeight;
    saveData(false, `＃　${wordName}`);
    renderWords();
}

// Delete word
async function deleteWord(index) {
    const wordName = words[index].word;
    const shouldDelete = await showInPageConfirm({
        title: 'Delete Word',
        message: `Delete "${wordName}"?`,
        confirmText: 'Delete',
        cancelText: 'Cancel',
        confirmTone: 'danger'
    });
    if (!shouldDelete) return;

    words.splice(index, 1);
    saveData(false, `－　${wordName}`);
    playActionSound('delete');
    renderWords();
    showStatus('Word deleted', 'success');
}

// Toggle word selection
function toggleWordSelection(index) {
    if (selectedWords.has(index)) {
        selectedWords.delete(index);
    } else {
        selectedWords.add(index);
    }
    // Update checkbox visual
    const cb = document.getElementById(`select-${index}`);
    if (cb) cb.checked = selectedWords.has(index);
    updateBatchToolbar();
    updateWordSelectionUI();
}

// Toggle sort mode
function toggleSortMode() {
    const modes = ['alpha', 'alpha-desc', 'chrono', 'chrono-desc', 'join', 'join-desc'];
    const labels = ['A-Z', 'Z-A', 'Date↑', 'Date↓', 'Join↑', 'Join↓'];
    const idx = modes.indexOf(wordSortMode);
    wordSortMode = modes[(idx + 1) % modes.length];
    const btn = document.getElementById('sortModeBtn');
    if (btn) btn.textContent = labels[modes.indexOf(wordSortMode)];
    renderWords();
}

// Toggle select mode
function toggleSelectMode() {
    isSelectMode = !isSelectMode;
    if (!isSelectMode) selectedWords.clear();
    withScrollAnchor(() => { renderWords(); updateBatchToolbar(); });
}

// Update batch toolbar and buttons
function updateBatchToolbar() {
    const toolbar = document.getElementById('batchToolbar');
    const selectModeBtn = document.getElementById('selectModeBtn');
    const batchUpBtn = document.getElementById('batchUpBtn');
    const batchDownBtn = document.getElementById('batchDownBtn');
    const batchDeleteBtn = document.getElementById('batchDeleteBtn');
    const batchSetDateBtn = document.getElementById('batchSetDateBtn');

    if (toolbar) {
        toolbar.style.display = isSelectMode ? 'flex' : 'none';
    }

    if (selectModeBtn) {
        selectModeBtn.textContent = isSelectMode ? 'Cancel Selection' : 'Select Words';
    }

    const hasSelection = selectedWords.size > 0;
    if (batchUpBtn) batchUpBtn.disabled = !hasSelection;
    if (batchDownBtn) batchDownBtn.disabled = !hasSelection;
    if (batchDeleteBtn) {
        batchDeleteBtn.disabled = !hasSelection;
        batchDeleteBtn.textContent = hasSelection ? `Delete (${selectedWords.size})` : 'Delete';
    }
    if (batchSetDateBtn) batchSetDateBtn.disabled = !hasSelection;
    const batchAddTagBtn = document.getElementById('batchAddTagBtn');
    const batchRemoveTagBtn = document.getElementById('batchRemoveTagBtn');
    if (batchAddTagBtn) batchAddTagBtn.disabled = !hasSelection;
    if (batchRemoveTagBtn) batchRemoveTagBtn.disabled = !hasSelection;

    if (isSelectMode) {
        requestAnimationFrame(syncDateInputDisplayMode);
    }
}

// Select all words
function selectAll() {
    const filteredWords = words.filter(w => w.weight >= -3);

    filteredWords.forEach(w => {
        const index = words.indexOf(w);
        selectedWords.add(index);
    });

    updateBatchToolbar();
    updateWordSelectionUI();

}

// Deselect all words
function selectNone() {
    selectedWords.clear();
    updateBatchToolbar();
    updateWordSelectionUI();

}

// Invert selection
function selectInvert() {
    const filteredWords = words.filter(w => w.weight >= -3);

    const newSelection = new Set();
    filteredWords.forEach(w => {
        const index = words.indexOf(w);
        if (!selectedWords.has(index)) {
            newSelection.add(index);
        }
    });

    selectedWords = newSelection;
    updateBatchToolbar();
    updateWordSelectionUI();

}

// Add or remove words matching predicate from selectedWords
function _modifySelectionByPredicate(predicate, action) {
    let count = 0;
    words.forEach((w, i) => {
        if (!predicate(w)) return;
        if (action === 'add') {
            selectedWords.add(i);
            count++;
        } else if (selectedWords.has(i)) {
            selectedWords.delete(i);
            count++;
        }
    });
    updateBatchToolbar();
    updateWordSelectionUI();
    return count;
}

// Parse weight range inputs
function getWeightRange() {
    const minVal = document.getElementById('weightMinInput').value.trim();
    const maxVal = document.getElementById('weightMaxInput').value.trim();
    const min = minVal === '' ? -3 : parseInt(minVal, 10);
    const max = maxVal === '' ? Infinity : parseInt(maxVal, 10);
    if (isNaN(min)) return null;
    if (maxVal !== '' && isNaN(max)) return null;
    return { min, max };
}

function addWeightRange() {
    const range = getWeightRange();
    if (!range) { showStatus('Invalid weight range', 'error'); return; }
    const count = _modifySelectionByPredicate(w => w.weight >= range.min && w.weight <= range.max, 'add');
    showStatus(`+${count} word(s)`, 'success');
}

function removeWeightRange() {
    const range = getWeightRange();
    if (!range) { showStatus('Invalid weight range', 'error'); return; }
    const count = _modifySelectionByPredicate(w => w.weight >= range.min && w.weight <= range.max, 'remove');
    showStatus(`−${count} word(s)`, 'success');
}

// Parse date range inputs
function getDateRange() {
    const minVal = document.getElementById('dateMinInput').value;
    const maxVal = document.getElementById('dateMaxInput').value;
    if (!minVal && !maxVal) return null;
    return { min: minVal || null, max: maxVal || null };
}

function addDateRange() {
    const range = getDateRange();
    if (!range) { showStatus('Please set at least one date', 'error'); return; }
    const count = _modifySelectionByPredicate(w => (!range.min || w.added >= range.min) && (!range.max || w.added <= range.max), 'add');
    showStatus(`+${count} word(s)`, 'success');
}

function removeDateRange() {
    const range = getDateRange();
    if (!range) { showStatus('Please set at least one date', 'error'); return; }
    const count = _modifySelectionByPredicate(w => (!range.min || w.added >= range.min) && (!range.max || w.added <= range.max), 'remove');
    showStatus(`−${count} word(s)`, 'success');
}

// Validate regex input in real-time
function validateRegexInput() {
    const input = document.getElementById('regexFilterInput');
    const errorEl = document.getElementById('regexSyntaxError');
    const pattern = input.value;
    if (!pattern) {
        errorEl.textContent = '';
        return true;
    }
    try {
        new RegExp(pattern, 'i');
        errorEl.textContent = '';
        return true;
    } catch (e) {
        errorEl.textContent = e.message.replace(/^.*:\s*/, '');
        return false;
    }
}

function selectByRegex() {
    const pattern = document.getElementById('regexFilterInput').value.trim();
    if (!pattern) { showStatus('Please enter a regex pattern', 'error'); return; }
    if (!validateRegexInput()) return;
    const regex = new RegExp(pattern, 'i');
    const count = _modifySelectionByPredicate(w => regex.test(w.word) || regex.test(w.meaning), 'add');
    showStatus(`Matched ${count} word(s)`, 'success');
}

function deselectByRegex() {
    const pattern = document.getElementById('regexFilterInput').value.trim();
    if (!pattern) { showStatus('Please enter a regex pattern', 'error'); return; }
    if (!validateRegexInput()) return;
    const regex = new RegExp(pattern, 'i');
    const count = _modifySelectionByPredicate(w => regex.test(w.word) || regex.test(w.meaning), 'remove');
    showStatus(`Unmatched ${count} word(s)`, 'success');
}

function selectByTag() {
    const tagId = tagSelectorState.batchFilter;
    if (!tagId) { showStatus('Please select a tag', 'error'); return; }
    const count = _modifySelectionByPredicate(w => (w.tags || []).includes(tagId), 'add');
    showStatus(`+${count} word(s)`, 'success');
}

function deselectByTag() {
    const tagId = tagSelectorState.batchFilter;
    if (!tagId) { showStatus('Please select a tag', 'error'); return; }
    const count = _modifySelectionByPredicate(w => (w.tags || []).includes(tagId), 'remove');
    showStatus(`−${count} word(s)`, 'success');
}

// Batch add tag to selected words
async function batchAddTag() {
    if (selectedWords.size === 0) return;
    const tagIds = tagSelectorState.batchAction;
    if (!tagIds || tagIds.length === 0) {
        showStatus('Please select a tag', 'error');
        return;
    }

    const count = selectedWords.size;
    selectedWords.forEach(index => {
        if (!words[index].tags) words[index].tags = [];
        tagIds.forEach(tagId => {
            if (!words[index].tags.includes(tagId)) {
                words[index].tags.push(tagId);
            }
        });
    });

    const names = tagIds.map(id => getTagName(id)).join(', ');
    saveData(false, `+tag "${names}" × ${count}`);
    renderWords();
    showStatus(`Added tag to ${count} word(s)`, 'success');
}

// Batch remove tag from selected words
async function batchRemoveTag() {
    if (selectedWords.size === 0) return;
    const tagIds = tagSelectorState.batchAction;
    if (!tagIds || tagIds.length === 0) {
        showStatus('Please select a tag', 'error');
        return;
    }

    let affected = 0;
    selectedWords.forEach(index => {
        if (!words[index].tags) return;
        tagIds.forEach(tagId => {
            const idx = words[index].tags.indexOf(tagId);
            if (idx !== -1) {
                words[index].tags.splice(idx, 1);
                affected++;
            }
        });
    });

    const names = tagIds.map(id => getTagName(id)).join(', ');
    saveData(false, `−tag "${names}" × ${affected}`);
    renderWords();
    showStatus(`Removed tag "${tag}" from ${affected} word(s)`, 'success');
}

// Batch adjust weight
async function batchAdjustWeight(delta) {
    if (selectedWords.size === 0) return;

    const count = selectedWords.size;
    const action = delta > 0 ? 'increased' : 'decreased';

    const shouldAdjust = await showInPageConfirm({
        title: 'Adjust Weight',
        message: `${delta > 0 ? 'Increase' : 'Decrease'} weight for ${count} selected word(s)?`,
        confirmText: 'Confirm',
        cancelText: 'Cancel'
    });
    if (!shouldAdjust) return;

    selectedWords.forEach(index => {
        const currentWeight = words[index].weight;
        const newWeight = currentWeight + delta;

        // Don't go below -2 or above 10
        if (newWeight >= -2 && newWeight <= 10) {
            words[index].weight = newWeight;
        }
    });

    saveData(false, `Weight ${action} for ${count} word(s)`);
    renderWords();
    showStatus(`Weight adjusted for ${count} word(s)`, 'success');
}

// Batch set date
async function batchSetDate() {
    if (selectedWords.size === 0) return;

    const dateInput = document.getElementById('batchDateInput');
    const targetDate = dateInput ? dateInput.value : '';
    if (!targetDate) {
        showStatus('Please choose a date', 'error');
        return;
    }

    const count = selectedWords.size;
    const shouldSetDate = await showInPageConfirm({
        title: 'Set Date',
        message: `Set added date to ${formatAddedDateLabel(targetDate) || targetDate} for ${count} selected word(s)?`,
        confirmText: 'Set',
        cancelText: 'Cancel'
    });
    if (!shouldSetDate) return;

    let updatedCount = 0;
    selectedWords.forEach(index => {
        if (words[index]) {
            words[index].added = targetDate;
            updatedCount++;
        }
    });

    if (updatedCount === 0) {
        showStatus('No words were updated', 'error');
        return;
    }

    saveData(false, `Date set to ${targetDate} for ${updatedCount} word(s)`);
    renderWords();
    showStatus(`Date updated for ${updatedCount} word(s)`, 'success');
}

// Update word selection UI
function updateWordSelectionUI() {
    // Update all checkboxes
    words.forEach((_, index) => {
        const checkbox = document.getElementById(`select-${index}`);
        if (checkbox) {
            checkbox.checked = selectedWords.has(index);
        }
    });
}

// Batch delete confirm
async function batchDeleteConfirm() {
    if (selectedWords.size === 0) return;

    const count = selectedWords.size;
    const shouldDelete = await showInPageConfirm({
        title: 'Delete Words',
        message: `Delete ${count} selected word(s)?`,
        confirmText: 'Delete',
        cancelText: 'Cancel',
        confirmTone: 'danger'
    });
    if (!shouldDelete) return;

    const indicesToDelete = Array.from(selectedWords).sort((a, b) => b - a);
    indicesToDelete.forEach(index => {
        words.splice(index, 1);
    });
    selectedWords.clear();
    saveData(false, `－　${count}`);
    playActionSound('delete');
    renderWords();
    updateBatchToolbar();
    showStatus(`Deleted ${count} word(s)`, 'success');
}

// Detect iOS (used for pronunciation gesture workaround)
const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
let pronunciationRequestId = 0;
let activePronunciationStop = null;

function interruptCurrentPronunciation() {
    if (activePronunciationStop) {
        try {
            activePronunciationStop();
        } catch (_) {}
        activePronunciationStop = null;
    }
    if ('speechSynthesis' in window && (speechSynthesis.speaking || speechSynthesis.pending)) {
        speechSynthesis.cancel();
    }
}

// Pronunciation — uses preloaded audio cache for instant playback
function pronounceWord(word) {
    const normalizedWord = normalizeAudioWord(word);
    if (!normalizedWord) return;

    const requestId = ++pronunciationRequestId;
    interruptCurrentPronunciation();

    const cached = audioCache.get(normalizedWord);

    // If cache has a decoded buffer, play it instantly via AudioContext
    if (cached && cached.status === 'ready' && cached.buffer) {
        const AudioCtx = window.AudioContext || window.webkitAudioContext;
        if (AudioCtx) {
            const ctx = new AudioCtx();
            const source = ctx.createBufferSource();
            source.buffer = cached.buffer;
            const gainNode = ctx.createGain();
            gainNode.gain.value = appSettings.pronounceVolume;
            source.connect(gainNode);
            gainNode.connect(ctx.destination);
            const stopPlayback = () => {
                source.onended = null;
                try { source.stop(0); } catch (_) {}
                ctx.close().catch(() => {});
            };
            activePronunciationStop = stopPlayback;
            source.onended = () => {
                if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
                ctx.close().catch(() => {});
            };
            try {
                source.start(0);
            } catch (_) {
                if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
                ctx.close().catch(() => {});
                pronounceSpeechSynthesis(word, requestId);
                return;
            }
            showStatus(`Playing "${normalizedWord}"`, 'info');
            return;
        }
    }

    // If cache has a blob URL / URL but no buffer, use Audio element
    if (cached && cached.status === 'ready' && (cached.blobUrl || cached.url)) {
        const audio = new Audio(cached.blobUrl || cached.url);
        audio.volume = appSettings.pronounceVolume;
        const stopPlayback = () => {
            audio.onended = null;
            audio.onerror = null;
            audio.pause();
            audio.currentTime = 0;
        };
        activePronunciationStop = stopPlayback;
        audio.onended = () => {
            if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
        };
        audio.onerror = () => {
            if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
        };
        audio.play().then(() => {
            showStatus(`Playing "${normalizedWord}"`, 'info');
        }).catch(() => {
            if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
            pronounceSpeechSynthesis(normalizedWord, requestId);
        });
        return;
    }

    // iOS Safari: speechSynthesis.speak() must be called synchronously within the
    // user gesture call stack. Fetching audio is async and the .catch() fallback
    // loses gesture context, so speechSynthesis would silently fail.
    // On iOS, speak synchronously NOW, then fetch audio in background for next time.
    if (isIOS) {
        pronounceSpeechSynthesis(normalizedWord, requestId);
        // Cache audio in background for future clicks (will hit the cached paths above)
        preloadAudioForWord(normalizedWord).catch(() => {});
        return;
    }

    // Desktop: fetch on demand, fall back to speechSynthesis
    const AudioCtx = window.AudioContext || window.webkitAudioContext;
    const audioCtx = AudioCtx ? new AudioCtx() : null;

    fetch(`https://api.dictionaryapi.dev/api/v2/entries/en_GB/${encodeURIComponent(normalizedWord)}`)
        .then(response => {
            if (!response.ok) throw new Error('API failed');
            return response.json();
        })
        .then(data => {
            if (requestId !== pronunciationRequestId) {
                if (audioCtx) audioCtx.close().catch(() => {});
                return;
            }
            const audioUrl = data[0]?.phonetics?.find(p => p.audio)?.audio;
            if (!audioUrl) throw new Error('No audio URL');

            // Store in cache for future use
            if (!audioCache.has(normalizedWord)) {
                audioCache.set(normalizedWord, {
                    status: 'ready',
                    url: audioUrl,
                    blobUrl: null,
                    buffer: null,
                    promise: null
                });
            } else {
                const e = audioCache.get(normalizedWord);
                e.url = audioUrl;
                e.status = 'ready';
            }

            if (audioCtx) {
                return fetch(audioUrl)
                    .then(r => r.arrayBuffer())
                    .then(buf => audioCtx.decodeAudioData(buf))
                    .then(decoded => {
                        if (requestId !== pronunciationRequestId) {
                            audioCtx.close().catch(() => {});
                            return;
                        }
                        // Cache the decoded buffer
                        const entry = audioCache.get(normalizedWord);
                        if (entry) entry.buffer = decoded;
                        const source = audioCtx.createBufferSource();
                        source.buffer = decoded;
                        const gainNode = audioCtx.createGain();
                        gainNode.gain.value = appSettings.pronounceVolume;
                        source.connect(gainNode);
                        gainNode.connect(audioCtx.destination);
                        const stopPlayback = () => {
                            source.onended = null;
                            try { source.stop(0); } catch (_) {}
                            audioCtx.close().catch(() => {});
                        };
                        activePronunciationStop = stopPlayback;
                        source.onended = () => {
                            if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
                            audioCtx.close().catch(() => {});
                        };
                        try {
                            source.start(0);
                        } catch (_) {
                            if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
                            audioCtx.close().catch(() => {});
                            pronounceSpeechSynthesis(word, requestId);
                            return;
                        }
                        showStatus(`Playing "${normalizedWord}"`, 'info');
                    });
            } else {
                const audio = new Audio(audioUrl);
                audio.volume = appSettings.pronounceVolume;
                const stopPlayback = () => {
                    audio.onended = null;
                    audio.onerror = null;
                    audio.pause();
                    audio.currentTime = 0;
                };
                activePronunciationStop = stopPlayback;
                audio.onended = () => {
                    if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
                };
                audio.onerror = () => {
                    if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
                };
                return audio.play().then(() => {
                    showStatus(`Playing "${normalizedWord}"`, 'info');
                });
            }
        })
        .catch(() => {
            if (audioCtx) audioCtx.close().catch(() => {});
            pronounceSpeechSynthesis(normalizedWord, requestId);
        });
}

// Speech synthesis pronunciation — called synchronously from user gesture on iOS
function pronounceSpeechSynthesis(word, requestId = pronunciationRequestId) {
    if (requestId !== pronunciationRequestId) return;
    if (!('speechSynthesis' in window)) {
        showStatus('Pronunciation not supported', 'error');
        return;
    }
    speechSynthesis.cancel();

    const utterance = new SpeechSynthesisUtterance(word);
    utterance.lang = 'en-GB';
    utterance.rate = 0.8;
    utterance.volume = appSettings.pronounceVolume;

    // Use cached voices (populated via voiceschanged), fallback to fresh call
    const voices = cachedVoices.length > 0 ? cachedVoices : speechSynthesis.getVoices();
    const britishVoice = voices.find(v =>
        v.lang.startsWith('en-GB') || v.lang.startsWith('en-UK')
    );
    if (britishVoice) utterance.voice = britishVoice;

    // iOS Safari bug: speech pauses after ~15s. Keep-alive timer resumes it.
    let iosResumeTimer = null;
    utterance.onstart = () => {
        iosResumeTimer = setInterval(() => {
            if (speechSynthesis.paused) speechSynthesis.resume();
        }, 5000);
    };
    const cleanup = () => { if (iosResumeTimer) { clearInterval(iosResumeTimer); iosResumeTimer = null; } };
    const stopPlayback = () => {
        cleanup();
        speechSynthesis.cancel();
    };
    activePronunciationStop = stopPlayback;
    utterance.onend = () => {
        cleanup();
        if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
    };
    utterance.onerror = (e) => {
        cleanup();
        if (activePronunciationStop === stopPlayback) activePronunciationStop = null;
        // 'interrupted' is normal (user clicked again), only log real errors
        if (e.error !== 'interrupted') {
            showStatus('Pronunciation failed', 'error');
        }
    };

    speechSynthesis.speak(utterance);
    showStatus(`Pronouncing "${word}"`, 'info');
}

// Open edit modal
function openEditModal(index) {
    editingIndex = index;
    const word = words[index];

    document.getElementById('editWordInput').value = word.word;
    document.getElementById('editMeaningInput').value = word.meaning;
    document.getElementById('editDateInput').value = word.added;
    tagSelectorState.edit = (word.tags || []).slice();
    renderTagSelector('edit');

    // Set POS selection
    editSelectedPos = Array.isArray(word.pos) ? word.pos.slice() : (word.pos ? [word.pos] : []);
    updateEditPosSelection();
    editSelectedWeight = Number.isInteger(word.weight) ? word.weight : parseInt(word.weight, 10);
    if (Number.isNaN(editSelectedWeight)) {
        editSelectedWeight = 3;
    }
    updateEditWeightSelection();

    document.getElementById('editModal').classList.add('active');
    initResponsiveDateInputs();
}

// Close edit modal
function closeEditModal() {
    editingIndex = -1;
    document.getElementById('editModal').classList.remove('active');
}

function bindModalBackdropPressReleaseClose(modal, onClose) {
    if (!modal || typeof onClose !== 'function') return;

    let pressedOnBackdrop = false;
    let pressedPointerId = null;

    function resetPressedState() {
        pressedOnBackdrop = false;
        pressedPointerId = null;
    }

    modal.addEventListener('pointerdown', function(e) {
        if (e.target === modal) {
            pressedOnBackdrop = true;
            pressedPointerId = e.pointerId;
            return;
        }
        resetPressedState();
    });

    modal.addEventListener('pointerup', function(e) {
        const releasedOnBackdrop = e.target === modal;
        const samePointer = pressedPointerId === e.pointerId;
        if (pressedOnBackdrop && releasedOnBackdrop && samePointer) {
            onClose();
        }
        resetPressedState();
    });

    modal.addEventListener('pointercancel', resetPressedState);
}

// Save edit
function saveEdit() {
    if (editingIndex === -1) return;

    const word = document.getElementById('editWordInput').value.trim();
    const meaning = document.getElementById('editMeaningInput').value.trim();
    const weight = editSelectedWeight;
    const date = document.getElementById('editDateInput').value;
    const tags = tagSelectorState.edit.slice();

    if (!word) {
        alert('Please fill in word');
        return;
    }

    const oldWord = words[editingIndex].word;

    words[editingIndex] = {
        word: word.toLowerCase(),
        meaning: meaning,
        pos: editSelectedPos.slice(), // Copy array
        weight: weight,
        added: date,
        joinedAt: words[editingIndex].joinedAt,
        tags: tags
    };

    saveData(false, `✎　${oldWord}`);
    renderWords();
    closeEditModal();
}

bindModalBackdropPressReleaseClose(
    document.getElementById('editModal'),
    closeEditModal
);

// Render words
function renderWords() {
    const container = document.getElementById('wordList');

    // Update tag filter bar
    renderTagFilterBar();

    // Show all words, apply tag filter
    let filteredWords = words.filter(w => w.weight >= -3);
    if (activeTagFilters.size > 0) {
        filteredWords = filteredWords.filter(w => {
            const wTags = w.tags || [];
            return Array.from(activeTagFilters).some(t => wTags.includes(t));
        });
    }
    if (excludedTagFilters.size > 0) {
        filteredWords = filteredWords.filter(w => {
            const wTags = w.tags || [];
            return !Array.from(excludedTagFilters).some(t => wTags.includes(t));
        });
    }

    if (filteredWords.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">—</div>
                <div>No Words</div>
            </div>
        `;
        applyAudioPreloadSetting();
        return;
    }

    // Sort helper
    function sortWordList(list) {
        if (wordSortMode === 'chrono') {
            list.sort((a, b) => (a.added || '').localeCompare(b.added || ''));
        } else if (wordSortMode === 'chrono-desc') {
            list.sort((a, b) => (b.added || '').localeCompare(a.added || ''));
        } else if (wordSortMode === 'join') {
            list.sort((a, b) => {
                const ja = a.joinedAt || ((a.added || '') + 'T00:00:00');
                const jb = b.joinedAt || ((b.added || '') + 'T00:00:00');
                return ja.localeCompare(jb);
            });
        } else if (wordSortMode === 'join-desc') {
            list.sort((a, b) => {
                const ja = a.joinedAt || ((a.added || '') + 'T00:00:00');
                const jb = b.joinedAt || ((b.added || '') + 'T00:00:00');
                return jb.localeCompare(ja);
            });
        } else if (wordSortMode === 'alpha-desc') {
            list.sort((a, b) => b.word.localeCompare(a.word));
        } else {
            list.sort((a, b) => a.word.localeCompare(b.word));
        }
    }

    let groupEntries; // Array of { label, words, collapsed }

    if (wordGroupMode === 'tag') {
        // Group by tag
        const tagGroups = {};
        const untagged = [];
        filteredWords.forEach(w => {
            const tags = w.tags || [];
            if (tags.length === 0) {
                untagged.push(w);
            } else {
                tags.forEach(t => {
                    if (!tagGroups[t]) tagGroups[t] = [];
                    tagGroups[t].push(w);
                });
            }
        });
        // Sort tag groups by name
        const sortedTagIds = Object.keys(tagGroups).sort((a, b) => getTagName(a).localeCompare(getTagName(b)));
        groupEntries = sortedTagIds.map(tagId => {
            sortWordList(tagGroups[tagId]);
            return { label: getTagName(tagId) || tagId, words: tagGroups[tagId], collapsed: false };
        });
        if (untagged.length > 0) {
            sortWordList(untagged);
            groupEntries.push({ label: 'Untagged', words: untagged, collapsed: true });
        }
    } else {
        // Group by weight (default)
        const groups = {};
        filteredWords.forEach(w => {
            if (!groups[w.weight]) groups[w.weight] = [];
            groups[w.weight].push(w);
        });
        const sortedWeights = Object.keys(groups).map(Number).sort((a, b) => b - a);
        groupEntries = sortedWeights.map(weight => {
            sortWordList(groups[weight]);
            const isNegative = weight < 0;
            return { label: getWeightLabel(weight), words: groups[weight], collapsed: isNegative };
        });
    }

    let html = '';
    groupEntries.forEach(({ label: groupLabel, words: groupWords, collapsed }) => {
        const expandedClass = collapsed ? '' : 'expanded';
        const collapseIcon = collapsed ? '▼' : '▲';

        html += `
            <div class="word-group">
                <div class="collapsible-header" onclick="toggleCollapse(this)">
                    <div class="group-header" style="margin-bottom: 0; border: none;">${groupLabel} (${groupWords.length})</div>
                    <div class="collapse-icon">${collapseIcon}</div>
                </div>
                <div class="collapsible-content ${expandedClass}">
        `;

        groupWords.forEach(w => {
            const originalIndex = words.indexOf(w);
            const isInvalid = w.weight === -3;
            const masteredClass = (w.weight < 0 && !isInvalid) ? 'mastered' : '';
            const invalidClass = isInvalid ? 'invalid' : '';
            const weightText = isInvalid ? '!' : String(w.weight);
            const weightShapeClass = weightText.length > 1 ? 'is-wide' : '';
            const weightDisplay = weightText.length > 1
                ? `<span class="word-weight-text is-squeezed">${weightText}</span>`
                : `<span class="word-weight-text">${weightText}</span>`;

            const posArray = Array.isArray(w.pos) ? w.pos : (w.pos ? [w.pos] : []);
            const posTags = posArray.length > 0
                ? posArray.map(p => `<span class="word-pos">${p}</span>`).join('')
                : '';

            const tagsArray = (Array.isArray(w.tags) ? w.tags : []).map(id => getTagName(id)).filter(n => n);
            const tagBadges = tagsArray.length > 0
                ? '<div class="word-tags">' + tagsArray.map(name => `<span class="word-tag">${name}</span>`).join('') + '</div>'
                : '';

            html += `
                <div class="word-item ${masteredClass} ${invalidClass}${isSelectMode ? ' select-mode' : ''}" data-word="${w.word}" ${isSelectMode ? `onclick="toggleWordSelection(${originalIndex})"` : ''}>
                    <div class="word-header">
                        <div style="display: flex; align-items: center; gap: 8px;">
                            ${isSelectMode ? `<input type="checkbox" id="select-${originalIndex}" class="word-checkbox" ${selectedWords.has(originalIndex) ? 'checked' : ''} tabindex="-1">` : ''}
                            <div>
                                <span class="word-title">${w.word}</span>
                                ${posTags}
                                ${!isSelectMode ? `<button class="btn-pronounce" onclick="pronounceWord('${w.word}')" title="Pronounce (British)" aria-label="Pronounce (British)">${SPEAKER_ICON_SVG}</button>` : ''}
                            </div>
                        </div>
                        <div class="word-weight ${weightShapeClass}">${weightDisplay}</div>
                    </div>
                    <div class="word-meaning"${hideMeaning ? ' style="visibility:hidden;height:0;margin:0;overflow:hidden;"' : ''}>${w.meaning}</div>
                    ${tagBadges}
                    <div class="word-meta">Date: ${formatAddedDateLabel(w.added)}</div>
                    ${!isSelectMode ? `<div class="word-actions">
                        ${w.weight >= 0 ? `<button class="btn-remember" onclick="updateWeight(${originalIndex}, -1)">Down</button>` : ''}
                        ${w.weight >= -1 && w.weight < 10 ? `<button class="btn-forget" onclick="updateWeight(${originalIndex}, 1)">Up</button>` : ''}
                        ${isInvalid ? `<button class="btn-secondary" onclick="updateWeight(${originalIndex}, 3)">Fix</button>` : ''}
                        <button class="btn-edit" onclick="openEditModal(${originalIndex})">Edit</button>
                        <button class="btn-delete" onclick="deleteWord(${originalIndex})">Del</button>
                    </div>` : ''}
                </div>
            `;
        });

        html += `
                </div>
            </div>
        `;
    });

    container.innerHTML = html;
    updateBatchToolbar();
    applyAudioPreloadSetting();
}

// Toggle collapse
function toggleCollapse(header) {
    const content = header.nextElementSibling;
    const icon = header.querySelector('.collapse-icon');

    if (content.classList.contains('expanded')) {
        content.classList.remove('expanded');
        icon.textContent = '▼';
    } else {
        content.classList.add('expanded');
        icon.textContent = '▲';
    }
}

// Data persistence
function saveData(showMessage = false, description = null) {
    localStorage.setItem('wordMemoryData', JSON.stringify(words));

    // Create version if description is provided
    if (description && versionControl) {
        versionControl.createVersion(words, description);
    }

    if (showMessage) {
        showStatus('Auto-saved to localStorage', 'success');
    }
}

// Manual save
function manualSave() {
    localStorage.setItem('wordMemoryData', JSON.stringify(words));
    showStatus(`Manually saved ${words.length} words`, 'success');
}

function loadData() {
    const saved = localStorage.getItem('wordMemoryData');
    if (saved) {
        words = JSON.parse(saved);
    }
}

// Export data
function _downloadJsonFile(obj, filename) {
    const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

// Export data only (without version history)
function exportDataOnly() {
    const projectId = getProjectId();
    _downloadJsonFile({ projectId, words, tagRegistry }, `wordlist-${toSafeFilenamePart(projectId)}.json`);
    showStatus('Exported data only', 'success');
}

function buildFullExportData() {
    if (!versionControl) return null;

    const projectId = getProjectId();
    return {
        projectId,
        words: words,
        tagRegistry: tagRegistry,
        versionHistory: {
            format: 'tree-v2',
            versions: Object.fromEntries(versionControl.versions),
            rootId: versionControl.rootId,
            currentId: versionControl.currentId,
            exportDate: new Date().toISOString()
        }
    };
}

// Export data with complete version history
function exportWithVersionHistory() {
    const exportData = buildFullExportData();
    if (!exportData) { showStatus('Version control not initialised', 'error'); return; }
    const projectId = getProjectId();
    _downloadJsonFile(exportData, `wordlist-${toSafeFilenamePart(projectId)}-full.json`);
    showStatus('Exported data with version history', 'success');
}

function viewCurrentFullRawFile() {
    const exportData = buildFullExportData();
    if (!exportData) {
        showStatus('Version control not initialised', 'error');
        return;
    }

    const dataStr = JSON.stringify(exportData, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    window.open(url, '_blank');
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    showStatus('Opened full raw file in new tab', 'success');
}

// Legacy export function (for backward compatibility)
function exportData() {
    exportDataOnly();
}

// Validate version history structure
function validateVersionHistory(versionHistory) {
    if (!versionHistory.format || (versionHistory.format !== 'tree-v1' && versionHistory.format !== 'tree-v2')) {
        return { valid: false, error: 'Unsupported format version' };
    }

    if (!versionHistory.versions || !versionHistory.rootId) {
        return { valid: false, error: 'Missing versions or rootId' };
    }

    const versions = new Map(Object.entries(versionHistory.versions));

    // Validate rootId exists
    if (!versions.has(versionHistory.rootId)) {
        return { valid: false, error: 'Root version not found' };
    }

    // Validate currentId exists
    if (versionHistory.currentId && !versions.has(versionHistory.currentId)) {
        return { valid: false, error: 'Current version not found' };
    }

    // Validate tree structure integrity
    for (const [id, version] of versions) {
        // Validate parent reference
        if (version.parentId && !versions.has(version.parentId)) {
            return { valid: false, error: `Invalid parent reference: ${version.parentId}` };
        }

        // Validate children references
        for (const childId of version.children || []) {
            if (!versions.has(childId)) {
                return { valid: false, error: `Invalid child reference: ${childId}` };
            }
        }
    }

    return { valid: true };
}

// Process and validate imported words data
function processImportedWords(imported) {
    let validCount = 0;
    let invalidCount = 0;

    const processed = imported.map(item => {
        // Convert old POS format (string) to new format (array)
        let posArray = [];
        if (item.pos) {
            if (Array.isArray(item.pos)) {
                posArray = item.pos;
            } else if (typeof item.pos === 'string' && item.pos.trim() !== '') {
                posArray = [item.pos.trim()];
            }
        }

        // Check if word exists
        if (!item.word || typeof item.word !== 'string' || item.word.trim() === '') {
            invalidCount++;
            return {
                word: item.word || '[no word]',
                meaning: item.meaning || '',
                pos: posArray,
                weight: -3,
                added: item.added || new Date().toISOString().split('T')[0],
                joinedAt: item.joinedAt || undefined
            };
        }

        // Check if weight is valid
        const weight = parseInt(item.weight);
        if (isNaN(weight) || weight < -2) {
            invalidCount++;
            return {
                word: item.word.toLowerCase(),
                meaning: item.meaning || '',
                pos: posArray,
                weight: -3,
                added: item.added || new Date().toISOString().split('T')[0],
                joinedAt: item.joinedAt || undefined
            };
        }

        // Valid word
        validCount++;
        return {
            word: item.word.toLowerCase(),
            meaning: item.meaning || '',
            pos: posArray,
            weight: weight,
            added: item.added || new Date().toISOString().split('T')[0],
            joinedAt: item.joinedAt || undefined
        };
    });

    return { processed, validCount, invalidCount };
}

function _applyImportedTagRegistry(importedTagRegistry) {
    if (importedTagRegistry && Array.isArray(importedTagRegistry)) {
        tagRegistry = importedTagRegistry;
    } else {
        tagRegistry = [];
        migrateStringTagsToRegistry();
    }
    saveTagRegistry();
}

// Import words only (clear version history)
function importWordsOnly(wordsData, importedTagRegistry) {
    const { processed, validCount, invalidCount } = processImportedWords(wordsData);
    words = processed;
    _applyImportedTagRegistry(importedTagRegistry);
    if (versionControl) versionControl.clearHistory();
    saveData(false, `Imported ${processed.length} word(s)`);
    renderWords();
    updateHistoryButtonLabel();
    if (invalidCount > 0) {
        showStatus(`Imported ${validCount} valid, ${invalidCount} invalid words (version history cleared)`, 'success');
    } else {
        showStatus(`Imported ${processed.length} words (version history cleared)`, 'success');
    }
}

// Import as overwrite: use imported data, create a new branch in registry-tree
function importAsOverwrite(wordsData, description, importedTagRegistry) {
    const { processed } = processImportedWords(wordsData);
    words = processed;
    _applyImportedTagRegistry(importedTagRegistry);
    localStorage.setItem('wordMemoryData', JSON.stringify(words));
    if (versionControl) {
        versionControl.createVersion(words, description || `Overwrite import (${processed.length} words)`);
    }
    renderWords();
    updateHistoryButtonLabel();
    showStatus(`Overwrite imported ${processed.length} words as new branch`, 'success');
}

// Import data with version history (as fork — grafts imported tree under local root)
function importWithVersionHistory(importedData) {
    const vh = importedData.versionHistory;
    const importedVersionCount = Object.keys(vh.versions).length;
    const newCurrentId = versionControl.graftTree(vh.versions, vh.rootId, vh.currentId);
    const currentVersion = versionControl.versions.get(newCurrentId);
    const resolvedData = currentVersion ? versionControl.resolveData(newCurrentId) : null;
    if (resolvedData) {
        words = JSON.parse(JSON.stringify(resolvedData));
    } else {
        const { processed } = processImportedWords(importedData.words);
        words = processed;
    }
    _applyImportedTagRegistry(importedData.tagRegistry);
    localStorage.setItem('wordMemoryData', JSON.stringify(words));
    renderWords();
    updateHistoryButtonLabel();
    showStatus(`Forked: imported ${importedVersionCount} version(s) as new branch`, 'success');
}

// Import data with version history (replace — clears local history entirely)
function importReplaceWithVersionHistory(importedData) {
    const { processed } = processImportedWords(importedData.words);
    const vh = importedData.versionHistory;
    versionControl.versions = new Map(Object.entries(vh.versions));
    versionControl.rootId = vh.rootId;
    versionControl.currentId = vh.currentId;
    versionControl._clearCache();
    versionControl.saveHistory();
    const currentVersion = versionControl.versions.get(versionControl.currentId);
    const resolvedData = currentVersion ? versionControl.resolveData(versionControl.currentId) : null;
    words = resolvedData ? JSON.parse(JSON.stringify(resolvedData)) : processed;
    _applyImportedTagRegistry(importedData.tagRegistry);
    if (importedData.projectId) {
        setProjectId(importedData.projectId);
    }
    localStorage.setItem('wordMemoryData', JSON.stringify(words));
    renderWords();
    updateHistoryButtonLabel();
    showStatus(`Replaced: imported ${versionControl.versions.size} version(s)`, 'success');
}

function isJsonImportFile(file) {
    if (!file) return false;

    const type = String(file.type || '').toLowerCase();
    const name = String(file.name || '').toLowerCase();

    return type === 'application/json' || type === 'text/json' || name.endsWith('.json');
}

async function importDataFromFile(file) {
    if (!file) return;

    if (!isJsonImportFile(file)) {
        showStatus('Only JSON files can be imported', 'error');
        return;
    }

    const shouldStartImport = await showInPageConfirm({
        title: 'Import JSON File',
        message: `Import "${file.name}"?`,
        confirmText: 'Import',
        cancelText: 'Cancel'
    });
    if (!shouldStartImport) {
        return;
    }

    try {
        const imported = JSON.parse(await file.text());
        // Check if it contains version history
        if (imported && imported.versionHistory) {
            const validation = validateVersionHistory(imported.versionHistory);

            if (!validation.valid) {
                const shouldContinueImport = await showInPageConfirm({
                    title: 'Version History Corrupted',
                    message: `Version history is corrupted: ${validation.error}\n\nContinue importing data only (version history will be cleared)?`,
                    confirmText: 'Continue',
                    cancelText: 'Cancel'
                });
                if (shouldContinueImport) {
                    importWordsOnly(imported.words || imported, imported.tagRegistry);
                }
                return;
            }

            const versionCount = Object.keys(imported.versionHistory.versions).length;
            const choice = await showInPageChoice({
                title: 'Import With History',
                message: `${imported.words.length} words, ${versionCount} version(s)`,
                choices: [
                    { text: 'Fork', value: 'fork' },
                    { text: 'Replace', value: 'replace', tone: 'danger' }
                ]
            });
            if (choice === 'fork') {
                importWithVersionHistory(imported);
            } else if (choice === 'replace') {
                importReplaceWithVersionHistory(imported);
            }
        } else if (Array.isArray(imported)) {
            // Old format - words array only
            const shouldImportWordsOnly = await showInPageConfirm({
                title: 'Import Words',
                message: `Import ${imported.length} words? This will overwrite current data and clear version history.`,
                confirmText: 'Import',
                cancelText: 'Cancel'
            });
            if (shouldImportWordsOnly) {
                importWordsOnly(imported);
            }
        } else if (imported && imported.words && Array.isArray(imported.words)) {
            const shouldOverwrite = await showInPageConfirm({
                title: 'Import as Branch',
                message: `Import ${imported.words.length} words as a new branch in version history?`,
                confirmText: 'Import',
                cancelText: 'Cancel'
            });
            if (shouldOverwrite) {
                importAsOverwrite(imported.words, `Import fork (${imported.words.length} words)`, imported.tagRegistry);
            }
        } else {
            showStatus('Invalid file format', 'error');
        }
    } catch (error) {
        showStatus('Failed to parse file', 'error');
    }
}

// Import data (enhanced with version history and projectId support)
async function importData(event) {
    const input = event && event.target ? event.target : null;
    const file = input && input.files ? input.files[0] : null;
    if (!file) return;

    await importDataFromFile(file);

    if (input) {
        input.value = '';
    }
}

let dragImportDepth = 0;

function isFileDragEvent(event) {
    if (!event || !event.dataTransfer) return false;
    return Array.from(event.dataTransfer.types || []).includes('Files');
}

function setDragImportOverlayVisible(isVisible) {
    if (!document.body) return;
    document.body.classList.toggle('drag-import-active', isVisible);
}

function clearDragImportOverlay() {
    dragImportDepth = 0;
    setDragImportOverlayVisible(false);
}

function initDragAndDropImport() {
    if (!document.body) return;
    if (document.body.dataset.dragImportBound === '1') return;
    document.body.dataset.dragImportBound = '1';

    window.addEventListener('dragenter', function(event) {
        if (!isFileDragEvent(event)) return;
        event.preventDefault();
        dragImportDepth += 1;
        setDragImportOverlayVisible(true);
    });

    window.addEventListener('dragover', function(event) {
        if (!isFileDragEvent(event)) return;
        event.preventDefault();
        if (event.dataTransfer) {
            event.dataTransfer.dropEffect = 'copy';
        }
        setDragImportOverlayVisible(true);
    });

    window.addEventListener('dragleave', function(event) {
        if (!isFileDragEvent(event)) return;
        event.preventDefault();
        dragImportDepth = Math.max(0, dragImportDepth - 1);
        if (dragImportDepth === 0) {
            setDragImportOverlayVisible(false);
        }
    });

    window.addEventListener('dragend', clearDragImportOverlay);
    window.addEventListener('blur', clearDragImportOverlay);

    window.addEventListener('drop', async function(event) {
        if (!isFileDragEvent(event)) return;
        event.preventDefault();
        clearDragImportOverlay();

        const droppedFiles = Array.from((event.dataTransfer && event.dataTransfer.files) || []);
        if (!droppedFiles.length) return;

        const jsonFile = droppedFiles.find(isJsonImportFile);
        if (!jsonFile) {
            showStatus('Only JSON files can be imported', 'error');
            return;
        }

        if (droppedFiles.length > 1) {
            showStatus(`Multiple files detected, importing "${jsonFile.name}"`, 'info');
        }

        await importDataFromFile(jsonFile);
    });
}

// Enter key to add word
document.getElementById('wordInput').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') addWord();
});
document.getElementById('meaningInput').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') addWord();
});

// Apply a resolved version's data to words + localStorage + re-render
function _applyVersionWords(versionId) {
    const data = versionControl.resolveData(versionId);
    if (!data) return false;
    words = JSON.parse(JSON.stringify(data));
    localStorage.setItem('wordMemoryData', JSON.stringify(words));
    renderWords();
    return true;
}

function _initActionSounds() {
    if (!appSettings.actionSoundEnabled) {
        deleteActionSoundAudio = null;
        putActionSoundAudio = null;
    } else {
        if (!deleteActionSoundAudio) {
            deleteActionSoundAudio = new Audio(getPreloadedAssetURL('assets/delete.mp3'));
            deleteActionSoundAudio.preload = 'auto';
        }
        if (!putActionSoundAudio) {
            putActionSoundAudio = new Audio(getPreloadedAssetURL('assets/put.mp3'));
            putActionSoundAudio.preload = 'auto';
        }
    }
}

// Undo/Redo functions
function performUndo() {
    if (!versionControl) { showStatus('Version control not initialised', 'error'); return; }
    if (!versionControl.canUndo()) { showStatus('Nothing to undo', 'info'); return; }
    const version = versionControl.undo();
    if (version) {
        _applyVersionWords(version.id);
        showStatus(`Undo: ${version.description}`, 'success');
    }
}

function performRedo() {
    if (!versionControl) { showStatus('Version control not initialised', 'error'); return; }
    if (!versionControl.canRedo()) { showStatus('Nothing to redo', 'info'); return; }
    const version = versionControl.redo();
    if (version) {
        _applyVersionWords(version.id);
        showStatus(`Redo: ${version.description}`, 'success');
    }
}

// Keyboard shortcuts for Undo/Redo
document.addEventListener('keydown', function(e) {
    // Check if any input field is focused
    const activeElement = document.activeElement;
    const isInputFocused = activeElement && (
        activeElement.tagName === 'INPUT' ||
        activeElement.tagName === 'TEXTAREA' ||
        activeElement.isContentEditable
    );

    // Don't intercept shortcuts when typing in input fields
    if (isInputFocused) {
        return;
    }

    if (shouldHandleRegistryPreviewSelectAll(e, isInputFocused)) {
        e.preventDefault();
        selectRegistryPreviewContent();
        return;
    }

    // Cmd+Z (Mac) or Ctrl+Z (Windows/Linux) - Undo
    if ((e.metaKey || e.ctrlKey) && e.key === 'z' && !e.shiftKey) {
        e.preventDefault();
        performUndo();
    }

    // Cmd+Y (Mac) or Ctrl+Y (Windows/Linux) - Redo
    // Cmd+Shift+Z (Mac) or Ctrl+Shift+Z (Windows/Linux) - Redo
    if ((e.metaKey || e.ctrlKey) && (e.key === 'y' || (e.shiftKey && e.key === 'z'))) {
        e.preventDefault();
        performRedo();
    }
});

// Settings functions
let _settingsDirty = false;
let _settingsSnapshot = null;

function openSettingsModal() {
    if (!versionControl) return;

    const settings = versionControl.getSettings();
    const actionSoundEnabledInput = document.getElementById('actionSoundEnabledInput');
    const audioPreloadEnabledInput = document.getElementById('audioPreloadEnabledInput');
    const quizAutoPlayInput = document.getElementById('quizAutoPlayInput');
    const pronounceVolumeInput = document.getElementById('pronounceVolumeInput');
    document.getElementById('maxVersionsInput').value = settings.maxVersions;
    document.getElementById('projectIdInput').value = getProjectId();
    if (actionSoundEnabledInput) {
        actionSoundEnabledInput.checked = appSettings.actionSoundEnabled;
    }
    if (audioPreloadEnabledInput) {
        audioPreloadEnabledInput.checked = appSettings.audioPreloadEnabled;
    }
    if (quizAutoPlayInput) {
        quizAutoPlayInput.checked = appSettings.quizAutoPlay;
    }
    if (pronounceVolumeInput) {
        pronounceVolumeInput.value = String(getPronounceVolumePercent());
    }
    updatePronounceVolumePreview();
    renderTagManager();

    // Reset dirty tracking
    _settingsDirty = false;
    _settingsSnapshot = {
        maxVersions: String(settings.maxVersions),
        projectId: getProjectId(),
        actionSoundEnabled: appSettings.actionSoundEnabled,
        audioPreloadEnabled: appSettings.audioPreloadEnabled,
        pronounceVolume: String(getPronounceVolumePercent())
    };

    document.querySelectorAll('.settings-toggle-item:not([data-click-bound])').forEach(item => {
        item.dataset.clickBound = 'true';
        item.addEventListener('click', (e) => {
            if (e.target.closest('input[type="checkbox"]') || e.target.closest('label')) return;
            const checkbox = item.querySelector('input[type="checkbox"]');
            if (checkbox) checkbox.click();
        });
    });

    updateHistoryButtonLabel();
    document.getElementById('settingsModal').classList.add('active');
}

function closeSettingsModal() {
    document.getElementById('settingsModal').classList.remove('active');
    _settingsDirty = false;
    _settingsSnapshot = null;
}

function isSettingsDirty() {
    if (_settingsDirty) return true;
    if (!_settingsSnapshot) return false;
    const maxVersionsEl = document.getElementById('maxVersionsInput');
    const projectIdEl = document.getElementById('projectIdInput');
    const actionSoundEl = document.getElementById('actionSoundEnabledInput');
    const audioPreloadEl = document.getElementById('audioPreloadEnabledInput');
    const quizAutoPlayEl = document.getElementById('quizAutoPlayInput');
    const pronounceVolumeEl = document.getElementById('pronounceVolumeInput');
    if (maxVersionsEl && maxVersionsEl.value !== _settingsSnapshot.maxVersions) return true;
    if (projectIdEl && projectIdEl.value !== _settingsSnapshot.projectId) return true;
    if (actionSoundEl && actionSoundEl.checked !== _settingsSnapshot.actionSoundEnabled) return true;
    if (audioPreloadEl && audioPreloadEl.checked !== _settingsSnapshot.audioPreloadEnabled) return true;
    if (quizAutoPlayEl && quizAutoPlayEl.checked !== _settingsSnapshot.quizAutoPlay) return true;
    if (pronounceVolumeEl && pronounceVolumeEl.value !== _settingsSnapshot.pronounceVolume) return true;
    return false;
}

async function cancelSettings() {
    if (isSettingsDirty()) {
        const shouldDiscard = await showInPageConfirm({
            title: 'Discard Changes',
            message: 'Settings have been modified. Discard changes?',
            confirmText: 'Discard',
            cancelText: 'Keep Editing',
            confirmTone: 'danger'
        });
        if (!shouldDiscard) return;
    }
    closeSettingsModal();
}

// --- Tag Manager (Settings) ---

function renderTagManager() {
    const container = document.getElementById('tagManagerList');
    if (!container) return;

    if (tagRegistry.length === 0) {
        container.innerHTML = '<div class="tag-manager-empty">No tags yet.</div>';
        return;
    }

    // Count usage per tag
    const usageCount = {};
    tagRegistry.forEach(t => usageCount[t.id] = 0);
    words.forEach(w => {
        if (w.tags) w.tags.forEach(id => {
            if (usageCount[id] !== undefined) usageCount[id]++;
        });
    });

    container.innerHTML = tagRegistry.map(t => `
        <div class="tag-manager-item" data-tag-id="${t.id}">
            <input type="text" value="${t.name.replace(/"/g, '&quot;')}" onchange="tagManagerRename('${t.id}', this.value)">
            <span class="tag-manager-count">${usageCount[t.id] || 0} words</span>
            <button class="btn-danger" onclick="tagManagerDelete('${t.id}')">Delete</button>
        </div>
    `).join('');
}

function tagManagerAdd() {
    const input = document.getElementById('tagManagerNewInput');
    if (!input) return;
    const name = input.value.trim();
    if (!name) return;

    // Check duplicate
    if (getTagId(name)) {
        alert('Tag "' + name + '" already exists.');
        return;
    }

    createTag(name);
    input.value = '';
    renderTagManager();
    renderTagFilterBar();
    _settingsDirty = true;
}

function tagManagerRename(id, newName) {
    newName = newName.trim();
    if (!newName) {
        alert('Tag name cannot be empty.');
        renderTagManager();
        return;
    }

    // Check duplicate (different id, same name)
    const existing = getTagId(newName);
    if (existing && existing !== id) {
        alert('Tag "' + newName + '" already exists.');
        renderTagManager();
        return;
    }

    renameTag(id, newName);
    renderTagManager();
    renderTagFilterBar();
    _settingsDirty = true;
}

async function tagManagerDelete(id) {
    const tag = tagRegistry.find(t => t.id === id);
    if (!tag) return;

    const count = words.filter(w => w.tags && w.tags.includes(id)).length;
    const msg = count > 0
        ? `Delete tag "${tag.name}"? It will be removed from ${count} word(s).`
        : `Delete tag "${tag.name}"?`;

    const shouldDelete = await showInPageConfirm({
        title: 'Delete Tag',
        message: msg,
        confirmText: 'Delete',
        cancelText: 'Cancel',
        confirmTone: 'danger'
    });
    if (!shouldDelete) return;

    deleteTag(id);
    saveData();
    playActionSound('delete');
    renderTagManager();
    renderTagFilterBar();
    _settingsDirty = true;
}

function saveSettings() {
    if (!versionControl) return;

    const maxVersions = parseInt(document.getElementById('maxVersionsInput').value, 10);
    const projectIdInput = document.getElementById('projectIdInput');
    const actionSoundEnabledInput = document.getElementById('actionSoundEnabledInput');
    const audioPreloadEnabledInput = document.getElementById('audioPreloadEnabledInput');
    const quizAutoPlayInput = document.getElementById('quizAutoPlayInput');
    const pronounceVolumeInput = document.getElementById('pronounceVolumeInput');
    const projectId = projectIdInput ? projectIdInput.value.trim() : '';
    const nextActionSoundEnabled = actionSoundEnabledInput ? actionSoundEnabledInput.checked : DEFAULT_APP_SETTINGS.actionSoundEnabled;
    const nextAudioPreloadEnabled = audioPreloadEnabledInput ? audioPreloadEnabledInput.checked : DEFAULT_APP_SETTINGS.audioPreloadEnabled;
    const nextQuizAutoPlay = quizAutoPlayInput ? quizAutoPlayInput.checked : DEFAULT_APP_SETTINGS.quizAutoPlay;
    const nextPronounceVolumePercent = pronounceVolumeInput ? parseInt(pronounceVolumeInput.value, 10) : 100;
    const nextPronounceVolume = Number.isFinite(nextPronounceVolumePercent) ? Math.max(0, Math.min(1, nextPronounceVolumePercent / 100)) : DEFAULT_APP_SETTINGS.pronounceVolume;

    if (isNaN(maxVersions) || maxVersions < 10 || maxVersions > 200) {
        alert('Max versions must be between 10 and 200');
        return;
    }

    if (!projectId) {
        alert('Profile ID cannot be empty');
        return;
    }

    if (projectId.length > 128) {
        alert('Profile ID must be 128 characters or less');
        return;
    }

    versionControl.updateSettings({ maxVersions: maxVersions });
    setProjectId(projectId);
    if (projectIdInput) {
        projectIdInput.value = projectId;
    }

    appSettings = sanitizeAppSettings({
        actionSoundEnabled: nextActionSoundEnabled,
        audioPreloadEnabled: nextAudioPreloadEnabled,
        pronounceVolume: nextPronounceVolume,
        quizAutoPlay: nextQuizAutoPlay
    });
    saveAppSettings();
    applyAudioPreloadSetting();
    enforceGroupModeByTagAvailability();
    renderWords();

    _initActionSounds();

    showStatus('Settings saved', 'success');
    closeSettingsModal();
}

async function clearAllData() {
    const choice = await showInPageChoice({
        title: 'Clear All Data',
        message: 'This will permanently delete all words, version history, settings, and cached data. This cannot be undone.',
        choices: [
            { text: 'Export Fully', value: 'export' },
            { text: 'Clear All', value: 'clear', tone: 'danger' }
        ]
    });
    if (choice === 'export') {
        exportWithVersionHistory();
        return;
    }
    if (choice !== 'clear') return;

    // Clear localStorage
    localStorage.clear();

    // Clear sessionStorage
    sessionStorage.clear();

    // Clear cookies
    document.cookie.split(';').forEach(cookie => {
        const name = cookie.split('=')[0].trim();
        if (name) {
            document.cookie = name + '=;expires=Thu, 01 Jan 1970 00:00:00 GMT;path=/';
        }
    });

    // Clear caches (Cache API)
    if ('caches' in window) {
        try {
            const keys = await caches.keys();
            await Promise.all(keys.map(key => caches.delete(key)));
        } catch (e) {}
    }

    // Reset in-memory state
    words = [];
    tagRegistry = [];
    if (versionControl) {
        versionControl.versions = new Map();
        versionControl.rootId = null;
        versionControl.currentId = null;
    }

    closeSettingsModal();
    renderWords();
    showStatus('All data cleared', 'success');
}

function generateProjectIdForSettings() {
    const projectIdInput = document.getElementById('projectIdInput');
    if (!projectIdInput) return;
    projectIdInput.value = createProjectId();
    projectIdInput.focus();
    projectIdInput.select();
}

// Version history functions — Registry Editor Style
function formatVersionDate(date) {
    if (!(date instanceof Date)) date = new Date(date);
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    const hh = String(date.getHours()).padStart(2, '0');
    const mi = String(date.getMinutes()).padStart(2, '0');
    return `${mm}-${dd}-${hh}-${mi}`;
}

let _historySelectedId = null;
let _historyExpandedSet = new Set();
let _historyPreviewMode = 'version'; // 'version' or 'diff'
let _historyCheckedSet = new Set(); // For batch operations
let _confirmModalResolver = null;

function getVersionNode(versionId) {
    if (!versionControl || !versionId) return null;
    return versionControl.versions.get(versionId) || null;
}

function getSortedChildIds(versionId) {
    const version = getVersionNode(versionId);
    if (!version) return [];

    const uniqueChildIds = Array.from(new Set(version.children || []));
    return uniqueChildIds
        .filter(childId => versionControl.versions.has(childId))
        .sort((a, b) => {
            const vA = versionControl.versions.get(a);
            const vB = versionControl.versions.get(b);
            const tA = Date.parse(vA && vA.timestamp ? vA.timestamp : 0) || 0;
            const tB = Date.parse(vB && vB.timestamp ? vB.timestamp : 0) || 0;
            if (tA !== tB) return tA - tB;
            return a.localeCompare(b);
        });
}

function normalizeHistoryState() {
    if (!versionControl) return;
    const validIds = versionControl.versions;

    _historyCheckedSet = new Set(
        Array.from(_historyCheckedSet).filter(id => validIds.has(id))
    );

    _historyExpandedSet = new Set(
        Array.from(_historyExpandedSet).filter(id => validIds.has(id))
    );

    if (versionControl.rootId && validIds.has(versionControl.rootId)) {
        _historyExpandedSet.add(versionControl.rootId);
    }

    if (_historySelectedId && !validIds.has(_historySelectedId)) {
        _historySelectedId = versionControl.currentId && validIds.has(versionControl.currentId)
            ? versionControl.currentId
            : (versionControl.rootId && validIds.has(versionControl.rootId) ? versionControl.rootId : null);
    }
}

function expandPathToVersion(versionId) {
    let cursor = versionId;
    const visited = new Set();

    while (cursor && !visited.has(cursor)) {
        visited.add(cursor);
        _historyExpandedSet.add(cursor);
        const version = getVersionNode(cursor);
        cursor = version ? version.parentId : null;
    }
}

function buildVersionPath(versionId) {
    const chain = [];
    let cursor = versionId;
    const visited = new Set();

    while (cursor && !visited.has(cursor)) {
        visited.add(cursor);
        const version = getVersionNode(cursor);
        if (!version) break;
        chain.unshift({ id: cursor, version });
        cursor = version.parentId;
    }

    return chain;
}

function getVersionDisplayLabel(version) {
    if (!version) return '';
    const desc = (version.description || '').trim();
    return desc || formatVersionDate(version.timestamp);
}

function getVersionDepth(versionId) {
    let depth = 0;
    let cursor = versionId;
    const visited = new Set();

    while (cursor && !visited.has(cursor)) {
        visited.add(cursor);
        const version = getVersionNode(cursor);
        if (!version || !version.parentId) break;
        depth++;
        cursor = version.parentId;
    }

    return depth;
}

function isDescendantOf(versionId, ancestorId) {
    if (!versionId || !ancestorId || versionId === ancestorId) return false;
    let cursor = versionId;
    const visited = new Set();

    while (cursor && !visited.has(cursor)) {
        visited.add(cursor);
        const version = getVersionNode(cursor);
        if (!version) return false;
        if (version.parentId === ancestorId) return true;
        cursor = version.parentId;
    }

    return false;
}

function resolveInPageConfirm(result) {
    const modal = document.getElementById('confirmModal');
    if (modal) {
        modal.classList.remove('active');
    }

    const resolver = _confirmModalResolver;
    _confirmModalResolver = null;
    if (resolver) {
        resolver(result);
    }
}

function initInPageConfirmModal() {
    const modal = document.getElementById('confirmModal');
    const confirmBtn = document.getElementById('confirmModalConfirmBtn');
    const cancelBtn = document.getElementById('confirmModalCancelBtn');
    if (!modal || !confirmBtn || !cancelBtn) return;
    if (modal.dataset.bound === '1') return;

    modal.dataset.bound = '1';
    confirmBtn.addEventListener('click', function() {
        resolveInPageConfirm(true);
    });
    cancelBtn.addEventListener('click', function() {
        resolveInPageConfirm(false);
    });
    bindModalBackdropPressReleaseClose(modal, function() {
        resolveInPageConfirm(false);
    });
    document.addEventListener('keydown', function(e) {
        if (e.key !== 'Escape') return;
        if (!modal.classList.contains('active')) return;
        e.preventDefault();
        resolveInPageConfirm(false);
    });
}

function showInPageConfirm({
    title = 'Confirm',
    message = 'Are you sure?',
    confirmText = 'Confirm',
    cancelText = 'Cancel',
    confirmTone = 'default'
} = {}) {
    const modal = document.getElementById('confirmModal');
    const titleEl = document.getElementById('confirmModalTitle');
    const messageEl = document.getElementById('confirmModalMessage');
    const confirmBtn = document.getElementById('confirmModalConfirmBtn');
    const cancelBtn = document.getElementById('confirmModalCancelBtn');

    if (!modal || !titleEl || !messageEl || !confirmBtn || !cancelBtn) {
        return Promise.resolve(false);
    }

    if (_confirmModalResolver) {
        _confirmModalResolver(false);
        _confirmModalResolver = null;
    }

    titleEl.textContent = title;
    messageEl.textContent = message;
    confirmBtn.textContent = confirmText;
    cancelBtn.textContent = cancelText;
    confirmBtn.classList.toggle('is-danger', confirmTone === 'danger');

    return new Promise((resolve) => {
        _confirmModalResolver = resolve;
        modal.classList.add('active');
        requestAnimationFrame(() => confirmBtn.focus({ preventScroll: true }));
    });
}

// Show a modal with multiple choice buttons. Returns the chosen value string, or null if cancelled.
function showInPageChoice({
    title = 'Choose',
    message = '',
    choices = [],  // [{ text: 'Fork', value: 'fork', tone: 'default' }, ...]
    cancelText = 'Cancel'
} = {}) {
    const modal = document.getElementById('confirmModal');
    const titleEl = document.getElementById('confirmModalTitle');
    const messageEl = document.getElementById('confirmModalMessage');
    const actionsEl = modal ? modal.querySelector('.modal-actions') : null;

    if (!modal || !titleEl || !messageEl || !actionsEl) {
        return Promise.resolve(null);
    }

    if (_confirmModalResolver) {
        _confirmModalResolver(null);
        _confirmModalResolver = null;
    }

    titleEl.textContent = title;
    messageEl.textContent = message;

    // Replace action buttons
    actionsEl.innerHTML = '';
    choices.forEach(choice => {
        const btn = document.createElement('button');
        btn.className = choice.tone === 'danger' ? 'btn-primary is-danger' : 'btn-primary';
        btn.textContent = choice.text;
        btn.addEventListener('click', () => resolveInPageConfirm(choice.value));
        actionsEl.appendChild(btn);
    });
    const cancelBtn = document.createElement('button');
    cancelBtn.className = 'btn-secondary';
    cancelBtn.textContent = cancelText;
    cancelBtn.addEventListener('click', () => resolveInPageConfirm(null));
    actionsEl.appendChild(cancelBtn);

    return new Promise((resolve) => {
        _confirmModalResolver = resolve;
        modal.classList.add('active');
        const firstBtn = actionsEl.querySelector('button');
        if (firstBtn) requestAnimationFrame(() => firstBtn.focus({ preventScroll: true }));
    }).finally(() => {
        // Restore original confirm/cancel buttons for showInPageConfirm
        actionsEl.innerHTML = '';
        const confirmBtn = document.createElement('button');
        confirmBtn.className = 'btn-primary';
        confirmBtn.id = 'confirmModalConfirmBtn';
        confirmBtn.textContent = 'Confirm';
        confirmBtn.addEventListener('click', () => resolveInPageConfirm(true));
        const cancelBtnRestore = document.createElement('button');
        cancelBtnRestore.className = 'btn-secondary';
        cancelBtnRestore.id = 'confirmModalCancelBtn';
        cancelBtnRestore.textContent = 'Cancel';
        cancelBtnRestore.addEventListener('click', () => resolveInPageConfirm(false));
        actionsEl.appendChild(confirmBtn);
        actionsEl.appendChild(cancelBtnRestore);
    });
}

function openHistoryModal() {
    if (!versionControl) return;

    _historyExpandedSet.clear();
    _historyCheckedSet.clear();
    _historySelectedId = versionControl.currentId || versionControl.rootId || null;

    if (versionControl.rootId) {
        _historyExpandedSet.add(versionControl.rootId);
    }
    if (_historySelectedId) {
        expandPathToVersion(_historySelectedId);
    }

    _historyPreviewMode = 'version';
    normalizeHistoryState();
    closeSettingsModal();
    document.getElementById('historyModal').classList.add('active');
    renderRegistryTree();
    renderRegistryPreview();
    updateRegistryStatus();
    updateRegistryAddressBar();
    updatePreviewTabs();
    updateRegistryToolbarButtons();
}

async function closeHistoryModal() {
    if (!versionControl) {
        document.getElementById('historyModal').classList.remove('active');
        closeMobilePreview();
        return;
    }

    // If a version is selected and it's not the current version, ask to apply
    if (_historySelectedId && _historySelectedId !== versionControl.currentId) {
        const version = versionControl.versions.get(_historySelectedId);
        if (version) {
            const desc = version.description || formatVersionDate(version.timestamp);
            const shouldApply = await showInPageConfirm({
                title: 'Apply Version',
                message: `Apply selected version?\n"${desc}"`,
                confirmText: 'Apply',
                cancelText: 'Keep Current'
            });
            if (shouldApply) {
                goToVersionById(_historySelectedId);
            }
        }
    }
    document.getElementById('historyModal').classList.remove('active');
    closeMobilePreview();
}

function updateRegistryAddressBar() {
    const pathEl = document.getElementById('registryPath');
    if (!pathEl || !versionControl) return;

    const chain = _historySelectedId ? buildVersionPath(_historySelectedId) : [];

    // Render clickable breadcrumbs
    let html = '<span class="registry-path-seg registry-path-root" onclick="navToRoot()">History</span>';
    chain.forEach((seg, i) => {
        html += '<span class="registry-path-sep">\\</span>';
        const isLast = i === chain.length - 1;
        const label = getVersionDisplayLabel(seg.version);
        html += `<span class="registry-path-seg${isLast ? ' registry-path-active' : ''}" onclick="selectVersionNode('${seg.id}')">${escapeHtml(label)}</span>`;
    });
    pathEl.innerHTML = html;
    requestAnimationFrame(() => {
        pathEl.scrollLeft = pathEl.scrollWidth;
    });
}

function navToRoot() {
    // Select root and collapse all
    if (!versionControl || !versionControl.rootId) return;
    _historySelectedId = versionControl.rootId;
    _historyExpandedSet.clear();
    _historyExpandedSet.add(versionControl.rootId);
    _historyCheckedSet.clear();
    renderRegistryTree();
    renderRegistryPreview();
    updateRegistryAddressBar();
    updateRegistryToolbarButtons();
}

function updateHistoryButtonLabel(totalVersions) {
    const historyBtn = document.getElementById('openHistoryBtn');
    if (!historyBtn) return;

    const total = Number.isFinite(totalVersions)
        ? totalVersions
        : (versionControl ? versionControl.versions.size : 0);
    historyBtn.textContent = `View History (ver.${total})`;
}

function updateRegistryStatus() {
    const left = document.getElementById('registryStatusLeft');
    const right = document.getElementById('registryStatusRight');
    const total = versionControl.versions.size;
    const canU = versionControl.canUndo();
    const canR = versionControl.canRedo();
    left.textContent = `${total} version(s)`;
    right.textContent = `Undo: ${canU ? 'Yes' : 'No'} | Redo: ${canR ? 'Yes' : 'No'}`;

    const info = document.getElementById('registryVersionInfo');
    info.textContent = `${total} ver.`;
    updateHistoryButtonLabel(total);
}

function updatePreviewTabs() {
    const tabs = document.getElementById('registryPreviewTabs');
    // Show tabs only on PC (handled by CSS mostly, but set active state)
    tabs.querySelectorAll('.registry-tab').forEach(t => {
        t.classList.toggle('active', t.dataset.mode === _historyPreviewMode);
    });
}

function switchPreviewMode(mode) {
    _historyPreviewMode = mode;
    updatePreviewTabs();
    renderRegistryPreview();
}

// ---- Tree Rendering ----

function renderRegistryTree() {
    if (!versionControl) return;
    const container = document.getElementById('historyTree');
    if (!container) return;

    normalizeHistoryState();

    if (!versionControl.rootId) {
        container.innerHTML = '<div class="registry-preview-empty">No Version History</div>';
        return;
    }

    container.innerHTML = buildRegistryTreeHTML();
}

function buildRegistryTreeHTML() {
    if (!versionControl || !versionControl.rootId) return '';

    const rows = [];
    const visited = new Set();

    const walk = (versionId, depth, visualDepth) => {
        if (!versionId || visited.has(versionId)) return;
        const version = getVersionNode(versionId);
        if (!version) return;
        visited.add(versionId);

        const sortedChildren = getSortedChildIds(versionId);
        const hasChildren = sortedChildren.length > 0;
        const isExpanded = _historyExpandedSet.has(versionId);

        rows.push({
            versionId,
            version,
            depth,
            visualDepth,
            hasChildren,
            isExpanded,
            childCount: sortedChildren.length
        });

        if (!hasChildren || !isExpanded) return;

        const childVisualDepth = visualDepth + (sortedChildren.length > 1 ? 1 : 0);
        sortedChildren.forEach((childId) => {
            walk(childId, depth + 1, childVisualDepth);
        });
    };

    walk(versionControl.rootId, 0, 0);

    return rows.map(renderTreeRowHTML).join('');
}

function renderTreeRowHTML(row) {
    const { versionId, version, depth, visualDepth, hasChildren, isExpanded, childCount } = row;
    const isCurrent = versionId === versionControl.currentId;
    const isSelected = versionId === _historySelectedId;
    const isChecked = _historyCheckedSet.has(versionId);

    const toggleClass = hasChildren
        ? (isExpanded ? 'tree-toggle has-children expanded' : 'tree-toggle has-children')
        : 'tree-toggle no-children';

    let rowClass = 'tree-node-row';
    if (isCurrent) rowClass += ' current-version';
    if (isSelected) rowClass += ' selected';

    const label = getVersionDisplayLabel(version);
    const wordCount = Number.isFinite(version.wordCount)
        ? version.wordCount
        : 0;
    const branchCount = childCount;
    const compressedDepth = Math.max(0, depth - visualDepth);

    let html = `<div class="${rowClass}" style="--tree-fork-depth:${visualDepth}" data-id="${versionId}" onclick="selectVersionNode('${versionId}')" ondblclick="goToVersionById('${versionId}')">`;
    html += `<span class="tree-indent-block" aria-hidden="true"></span>`;
    html += `<input type="checkbox" class="tree-checkbox" ${isChecked ? 'checked' : ''} onclick="event.stopPropagation(); toggleTreeCheck('${versionId}')" tabindex="-1">`;
    html += `<span class="${toggleClass}" onclick="event.stopPropagation(); toggleTreeNode('${versionId}')"></span>`;
    html += `<span class="tree-label">`;
    html += `<span class="tree-label-main">`;
    html += `<span class="tree-label-desc">${escapeHtml(label)}</span>`;
    html += `</span>`;
    html += `<span class="tree-label-meta">`;
    if (branchCount > 1) {
        html += `<span class="tree-label-branch">${branchCount} branches</span>`;
    }
    if (compressedDepth >= 3) {
        html += `<span class="tree-label-compressed">+${compressedDepth}</span>`;
    }
    html += `<span class="tree-label-count">(${wordCount})</span>`;
    html += `</span>`;
    html += `</span>`;
    html += `</div>`;

    return html;
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str === null || str === undefined ? '' : String(str);
    return div.innerHTML;
}

function toggleTreeNode(versionId) {
    const childIds = getSortedChildIds(versionId);
    if (childIds.length === 0) return;

    if (_historyExpandedSet.has(versionId)) {
        _historyExpandedSet.delete(versionId);

        if (_historySelectedId && isDescendantOf(_historySelectedId, versionId)) {
            _historySelectedId = versionId;
            renderRegistryPreview();
            updateRegistryAddressBar();
        }
    } else {
        _historyExpandedSet.add(versionId);
    }
    renderRegistryTree();
}

function toggleTreeCheck(versionId) {
    if (_historyCheckedSet.has(versionId)) {
        _historyCheckedSet.delete(versionId);
    } else {
        _historyCheckedSet.add(versionId);
    }
    updateRegistryToolbarButtons();
    renderRegistryTree();
}

function updateRegistryToolbarButtons() {
    const deleteBtn = document.getElementById('registryDeleteBtn');
    const renameBtn = document.getElementById('registryRenameBtn');
    const hasChecked = _historyCheckedSet.size > 0;
    const hasSingleChecked = _historyCheckedSet.size === 1;

    if (deleteBtn) {
        deleteBtn.disabled = !hasChecked;
        deleteBtn.textContent = hasChecked ? `Del (${_historyCheckedSet.size})` : 'Del';
    }
    if (renameBtn) {
        renameBtn.disabled = !hasSingleChecked;
    }
}

async function deleteSelectedVersions() {
    if (_historyCheckedSet.size === 0) return;

    let selectedIds = Array.from(_historyCheckedSet)
        .filter(id => versionControl.versions.has(id))
        .sort((a, b) => getVersionDepth(b) - getVersionDepth(a));

    const hadRootSelected = selectedIds.includes(versionControl.rootId);
    selectedIds = selectedIds.filter(id => id !== versionControl.rootId);
    const selectedCount = selectedIds.length;

    if (selectedCount === 0) {
        _historyCheckedSet.clear();
        updateRegistryToolbarButtons();
        if (hadRootSelected) {
            alert('Cannot delete the root version.');
        }
        return;
    }

    const currentId = versionControl.currentId;
    const deletingCurrent = selectedIds.includes(currentId);

    const shouldDelete = await showInPageConfirm({
        title: 'Delete Versions',
        message: `Delete ${selectedCount} version(s)?\n\nTheir children will be kept and reparented.`,
        confirmText: 'Delete',
        cancelText: 'Cancel',
        confirmTone: 'danger'
    });
    if (!shouldDelete) return;

    for (const versionId of selectedIds) {
        const version = versionControl.versions.get(versionId);
        if (!version) continue;
        versionControl.deleteVersionAndReparentChildren(versionId);
    }

    if (!versionControl.currentId || !versionControl.versions.has(versionControl.currentId)) {
        versionControl.currentId = (versionControl.rootId && versionControl.versions.has(versionControl.rootId))
            ? versionControl.rootId
            : null;
    }

    versionControl.saveHistory();

    if (deletingCurrent) {
        const currentVersion = versionControl.versions.get(versionControl.currentId);
        if (currentVersion) {
            _applyVersionWords(currentVersion.id);
        }
    }

    _historyCheckedSet.clear();

    if (_historySelectedId && !versionControl.versions.has(_historySelectedId)) {
        _historySelectedId = versionControl.currentId;
    }

    renderRegistryTree();
    renderRegistryPreview();
    updateRegistryStatus();
    updateRegistryAddressBar();
    updateRegistryToolbarButtons();
    playActionSound('delete');

    if (hadRootSelected) {
        showStatus(`Deleted ${selectedCount} version(s). Root version was skipped.`, 'success');
        return;
    }
    showStatus(`Deleted ${selectedCount} version(s)`, 'success');
}

function renameSelectedVersion() {
    if (_historyCheckedSet.size !== 1) return;

    const versionId = Array.from(_historyCheckedSet)[0];
    const version = versionControl.versions.get(versionId);
    if (!version) return;

    const newDesc = prompt('Rename version:', version.description || '');
    if (newDesc === null) return;

    version.description = newDesc.trim() || version.description;
    versionControl.saveHistory();

    _historyCheckedSet.clear();
    renderRegistryTree();
    renderRegistryPreview();
    updateRegistryAddressBar();
    updateRegistryToolbarButtons();
}

function selectVersionNode(versionId) {
    if (!versionControl || !versionControl.versions.has(versionId)) return;

    _historySelectedId = versionId;
    expandPathToVersion(versionId);

    renderRegistryTree();
    renderRegistryPreview();
    updateRegistryAddressBar();

    // On portrait mobile, open mobile preview
    if (isPortraitMobile()) {
        openMobilePreview(versionId);
    }
}

function isPortraitMobile() {
    return window.matchMedia(PORTRAIT_MOBILE_QUERY).matches;
}

// ---- Mobile Preview (portrait) ----

function openMobilePreview(versionId) {
    const overlay = document.getElementById('registryMobilePreview');
    const content = document.getElementById('mobilePreviewContent');
    const title = document.getElementById('mobilePreviewTitle');
    const gotoBtn = document.getElementById('mobileGotoBtn');

    const version = versionControl.versions.get(versionId);
    if (!version) return;

    title.textContent = version.description || formatVersionDate(version.timestamp);
    gotoBtn.setAttribute('onclick', `goToVersionById('${versionId}'); closeMobilePreview();`);

    content.innerHTML = renderJsonHighlight(versionControl.resolveData(versionId));
    overlay.classList.add('active');
}

function closeMobilePreview() {
    document.getElementById('registryMobilePreview').classList.remove('active');
}

// ---- JSON Preview Rendering ----

function renderRegistryPreview() {
    const container = document.getElementById('registryPreview');

    if (!_historySelectedId) {
        container.innerHTML = '<div class="registry-preview-empty">Select a version to preview</div>';
        return;
    }

    const version = versionControl.versions.get(_historySelectedId);
    if (!version) {
        container.innerHTML = '<div class="registry-preview-empty">Version not found</div>';
        return;
    }

    if (_historyPreviewMode === 'diff') {
        const currentVersion = versionControl.getCurrentVersion();
        if (!currentVersion || currentVersion.id === _historySelectedId) {
            container.innerHTML = '<div class="registry-preview-empty">Select a different version to compare with current</div>';
            return;
        }
        container.innerHTML = renderDiffView(
            versionControl.resolveData(version.id),
            versionControl.resolveData(currentVersion.id),
            version, currentVersion
        );
    } else {
        container.innerHTML = renderJsonHighlight(versionControl.resolveData(version.id));
    }
}

function renderJsonHighlight(data) {
    const jsonStr = JSON.stringify(data, null, 2);
    const lines = jsonStr.split('\n');
    let html = '';
    lines.forEach((line, i) => {
        const lineNum = i + 1;
        const highlighted = highlightJsonLine(line);
        html += `<span class="json-line"><span class="json-line-number">${lineNum}</span>${highlighted}</span>`;
    });
    return html;
}

function highlightJsonLine(line) {
    // Replace special HTML chars first
    line = line.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    // Tokenize and highlight
    let result = '';
    let i = 0;
    while (i < line.length) {
        // Whitespace
        if (line[i] === ' ' || line[i] === '\t') {
            result += line[i]; i++; continue;
        }
        // Quoted string
        if (line[i] === '"') {
            let j = i + 1;
            while (j < line.length && line[j] !== '"') {
                if (line[j] === '\\') j++; // skip escaped char
                j++;
            }
            j++; // include closing quote
            const str = line.substring(i, j);
            // Check if this is a key (followed by :)
            let k = j;
            while (k < line.length && line[k] === ' ') k++;
            if (line[k] === ':') {
                result += `<span class="json-key">${str}</span>`;
            } else {
                result += `<span class="json-string">${str}</span>`;
            }
            i = j; continue;
        }
        // Number
        if (line[i] === '-' || (line[i] >= '0' && line[i] <= '9')) {
            let j = i;
            if (line[j] === '-') j++;
            while (j < line.length && ((line[j] >= '0' && line[j] <= '9') || line[j] === '.')) j++;
            if (j > i && (j === i + 1 ? line[i] !== '-' : true)) {
                result += `<span class="json-number">${line.substring(i, j)}</span>`;
                i = j; continue;
            }
        }
        // Boolean / null
        const remaining = line.substring(i);
        const kwMatch = remaining.match(/^(true|false|null)\b/);
        if (kwMatch) {
            result += `<span class="json-boolean">${kwMatch[1]}</span>`;
            i += kwMatch[1].length; continue;
        }
        // Brackets
        if ('{[]}'.includes(line[i]) || line[i] === '}') {
            result += `<span class="json-bracket">${line[i]}</span>`;
            i++; continue;
        }
        // Other chars (colon, comma, etc.)
        result += line[i]; i++;
    }
    return result;
}

// ---- Diff View ----

function renderDiffView(dataA, dataB, versionA, versionB) {
    const jsonA = JSON.stringify(dataA, null, 2).split('\n');
    const jsonB = JSON.stringify(dataB, null, 2).split('\n');

    const diff = computeLineDiff(jsonA, jsonB);

    const dateA = formatVersionDate(versionA.timestamp);
    const dateB = formatVersionDate(versionB.timestamp);

    let html = `<div class="diff-header">Selected: ${dateA} vs Current: ${dateB}</div>`;

    let lineNum = 0;
    diff.forEach(entry => {
        if (entry.type === 'equal') {
            lineNum++;
            html += `<span class="json-line"><span class="json-line-number">${lineNum}</span>${highlightJsonLine(entry.line)}</span>`;
        } else if (entry.type === 'removed') {
            html += `<span class="json-line diff-removed"><span class="json-line-number">-</span>${highlightJsonLine(entry.line)}</span>`;
        } else if (entry.type === 'added') {
            lineNum++;
            html += `<span class="json-line diff-added"><span class="json-line-number">+</span>${highlightJsonLine(entry.line)}</span>`;
        }
    });

    return html;
}

// Simple LCS-based line diff
function computeLineDiff(linesA, linesB) {
    const m = linesA.length;
    const n = linesB.length;

    // Build LCS table
    const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            if (linesA[i - 1] === linesB[j - 1]) {
                dp[i][j] = dp[i - 1][j - 1] + 1;
            } else {
                dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
            }
        }
    }

    // Backtrack
    const result = [];
    let i = m, j = n;
    while (i > 0 || j > 0) {
        if (i > 0 && j > 0 && linesA[i - 1] === linesB[j - 1]) {
            result.unshift({ type: 'equal', line: linesA[i - 1] });
            i--; j--;
        } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
            result.unshift({ type: 'added', line: linesB[j - 1] });
            j--;
        } else {
            result.unshift({ type: 'removed', line: linesA[i - 1] });
            i--;
        }
    }

    return result;
}

function goToVersionById(versionId) {
    if (!versionControl) return;

    const version = versionControl.goToVersion(versionId);
    if (version) {
        _applyVersionWords(versionId);
        renderRegistryTree();
        renderRegistryPreview();
        updateRegistryStatus();
        updateRegistryAddressBar();
        showStatus(`Jumped to version: ${version.description}`, 'success');
    }
}

bindModalBackdropPressReleaseClose(
    document.getElementById('settingsModal'),
    cancelSettings
);

bindModalBackdropPressReleaseClose(
    document.getElementById('historyModal'),
    closeHistoryModal
);

// ========================================
// Quiz Mode
// ========================================

let quizState = null;
let _dingAudio = null;

function _playDing() {
    if (!_dingAudio) {
        _dingAudio = new Audio(getPreloadedAssetURL('assets/ding.mp3'));
        _dingAudio.preload = 'auto';
    }
    try {
        _dingAudio.currentTime = 0;
        const p = _dingAudio.play();
        if (p && typeof p.catch === 'function') p.catch(() => {});
    } catch (_) {}
}

// --- Levenshtein distance ---
function levenshtein(a, b) {
    const m = a.length, n = b.length;
    if (m === 0) return n;
    if (n === 0) return m;
    const dp = [];
    for (let i = 0; i <= m; i++) {
        dp[i] = new Array(n + 1);
        dp[i][0] = i;
    }
    for (let j = 0; j <= n; j++) dp[0][j] = j;
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            dp[i][j] = a[i - 1] === b[j - 1]
                ? dp[i - 1][j - 1]
                : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
        }
    }
    return dp[m][n];
}

// --- Find distractors (similar-looking words) ---
function findDistractors(targetWord, allWords, count = 3) {
    const others = allWords.filter(w => w !== targetWord && (w.meaning || '').trim());
    if (others.length <= count) return others;

    // Sort by character similarity (Levenshtein), then add randomness
    const scored = others.map(w => ({
        w,
        dist: levenshtein(w.word.toLowerCase(), targetWord.word.toLowerCase())
    }));
    scored.sort((a, b) => a.dist - b.dist);

    // Take closest 2× count candidates, shuffle, pick count
    const candidates = scored.slice(0, Math.min(count * 4, scored.length));
    candidates.sort(() => Math.random() - 0.5);
    return candidates.slice(0, count).map(x => x.w);
}

// --- Quiz setup ---
function openQuizSetup() {
    _updateQuizSetupInfo();
    document.getElementById('quizSetupModal').classList.add('active');

    // Preload pronunciation for all quiz pool words in the background
    if (appSettings.audioPreloadEnabled) {
        const pool = _getQuizPool();
        pool.forEach(w => queueWordAudioPreload(w.word));
    }
}

function closeQuizSetup() {
    document.getElementById('quizSetupModal').classList.remove('active');
}

// Pool = selectedWords (if any are selected) OR all words with meanings
function _getQuizPool() {
    if (isSelectMode && selectedWords.size > 0) {
        return Array.from(selectedWords)
            .map(i => words[i])
            .filter(w => w && (w.meaning || '').trim() !== '');
    }
    return words.filter(w => (w.meaning || '').trim() !== '');
}

function _updateQuizSetupInfo() {
    const pool = _getQuizPool();
    const infoEl = document.getElementById('quizSetupInfo');
    if (!infoEl) return;
    if (isSelectMode && selectedWords.size > 0) {
        infoEl.textContent = `${pool.length} selected word(s)`;
    } else {
        infoEl.textContent = `${pool.length} word(s) available`;
    }
}

function startQuiz() {
    const modeEl = document.querySelector('input[name="quizMode"]:checked');
    const mode = modeEl ? modeEl.value : 'spelling';
    const countRaw = document.getElementById('quizCountInput').value.trim();
    const count = parseInt(countRaw) || 0;

    let pool = _getQuizPool();

    if (pool.length === 0) {
        showStatus('No words match the criteria', 'error');
        return;
    }
    if (mode === 'mc' && words.filter(w => (w.meaning || '').trim()).length < 4) {
        showStatus('Need at least 4 words with meanings for multiple choice', 'error');
        return;
    }

    // Shuffle
    const shuffled = [...pool].sort(() => Math.random() - 0.5);
    const quizWords = count > 0 ? shuffled.slice(0, count) : shuffled;

    quizState = {
        mode,
        words: quizWords,
        currentIndex: 0,
        results: [],      // { wordRef, correct, userAnswer }
        answered: false,
        currentChoices: null,
        correctChoiceIndex: -1
    };

    closeQuizSetup();
    _showQuizQuestion();
}

function _showQuizQuestion() {
    if (!quizState) return;

    if (quizState.currentIndex >= quizState.words.length) {
        _endQuiz();
        return;
    }

    const w = quizState.words[quizState.currentIndex];
    const total = quizState.words.length;
    const current = quizState.currentIndex + 1;
    quizState.answered = false;

    // Preload audio for current word (and next)
    if (appSettings.audioPreloadEnabled) {
        queueWordAudioPreload(w.word);
        const next = quizState.words[quizState.currentIndex + 1];
        if (next) queueWordAudioPreload(next.word);
    }

    document.getElementById('quizProgress').textContent = `${current} / ${total}`;
    document.getElementById('quizResultArea').style.display = 'none';
    document.getElementById('quizRevealRow').style.display = 'none';
    document.getElementById('quizNextBtn').style.display = 'none';
    const skipBtn = document.getElementById('quizSkipBtn');
    if (skipBtn) skipBtn.style.display = '';

    if (quizState.mode === 'spelling') {
        document.getElementById('quizSpellingSection').style.display = '';
        document.getElementById('quizMCSection').style.display = 'none';

        const posArray = Array.isArray(w.pos) ? w.pos : (w.pos ? [w.pos] : []);
        document.getElementById('quizMeaningDisplay').textContent = w.meaning;
        document.getElementById('quizPosDisplay').textContent =
            posArray.length > 0 ? posArray.map(p => p + '.').join(' ') : '';
        document.getElementById('quizSpellingInput').value = '';
        setTimeout(() => document.getElementById('quizSpellingInput').focus(), 80);
    } else {
        document.getElementById('quizSpellingSection').style.display = 'none';
        document.getElementById('quizMCSection').style.display = '';

        const posArray = Array.isArray(w.pos) ? w.pos : (w.pos ? [w.pos] : []);
        document.getElementById('quizWordDisplay').textContent = w.word;
        document.getElementById('quizWordPosDisplay').textContent =
            posArray.length > 0 ? posArray.map(p => p + '.').join(' ') : '';

        // Build choices: 1 correct + 3 distractors, shuffled
        const allWithMeaning = words.filter(wd => (wd.meaning || '').trim());
        const distractors = findDistractors(w, allWithMeaning, 3);
        // Ensure we have 4 choices (pad with random if needed)
        while (distractors.length < 3) {
            const extra = allWithMeaning.find(wd => wd !== w && !distractors.includes(wd));
            if (extra) distractors.push(extra);
            else break;
        }
        const choices = [w, ...distractors].sort(() => Math.random() - 0.5);
        quizState.currentChoices = choices;
        quizState.correctChoiceIndex = choices.indexOf(w);

        const choicesEl = document.getElementById('quizChoices');
        choicesEl.innerHTML = choices.map((c, i) =>
            `<button class="quiz-choice-btn" onclick="handleMCChoice(${i})">${c.meaning}</button>`
        ).join('');
    }

    document.getElementById('quizQuestionModal').classList.add('active');

    // Auto-play pronunciation if enabled (only for MC mode where the word is shown,
    // or spelling mode — user hears the word they need to type)
    if (appSettings.quizAutoPlay) {
        setTimeout(() => pronounceQuizWord(), 300);
    }
}

function pronounceQuizWord() {
    if (!quizState) return;
    const w = quizState.words[quizState.currentIndex];
    if (w) pronounceWord(w.word);
}

function quizSpellingKeydown(event) {
    if (event.key === 'Enter') {
        if (!quizState || quizState.answered) {
            nextQuizQuestion();
        } else {
            checkSpelling();
        }
    }
}

function _isTypingTarget(target) {
    if (!target) return false;
    const tagName = target.tagName;
    return tagName === 'INPUT' || tagName === 'TEXTAREA' || target.isContentEditable;
}

function _getQuizChoiceIndexFromKey(event) {
    const key = event.key;
    if (key === '1' || event.code === 'Digit1' || event.code === 'Numpad1') return 0;
    if (key === '2' || event.code === 'Digit2' || event.code === 'Numpad2') return 1;
    if (key === '3' || event.code === 'Digit3' || event.code === 'Numpad3') return 2;
    if (key === '4' || event.code === 'Digit4' || event.code === 'Numpad4') return 3;
    return -1;
}

function goToPreviousQuizQuestion() {
    if (!quizState) return;
    if (quizState.currentIndex <= 0) {
        showStatus('Already at the first question', 'info');
        return;
    }

    const previousIndex = quizState.currentIndex - 1;
    // Rewind results so previous question can be answered again without duplicate scoring.
    quizState.results = quizState.results.slice(0, previousIndex);
    quizState.currentIndex = previousIndex;
    quizState.answered = false;
    _showQuizQuestion();
}

// Global quiz shortcuts: S pronounce, L skip, J previous question, 1-4 choose in MC mode
document.addEventListener('keydown', function(e) {
    if (!quizState) return;
    if (!document.getElementById('quizQuestionModal').classList.contains('active')) return;

    const isTyping = _isTypingTarget(e.target);

    if (!isTyping && (e.key === 'l' || e.key === 'L')) {
        e.preventDefault();
        skipQuizQuestion();
        return;
    }

    if (!isTyping && (e.key === 'j' || e.key === 'J')) {
        e.preventDefault();
        goToPreviousQuizQuestion();
        return;
    }

    if (e.key === 's' || e.key === 'S') {
        if (isTyping) return;
        e.preventDefault();
        pronounceQuizWord();
        return;
    }

    if (!isTyping && quizState.mode === 'mc' && !quizState.answered) {
        const choiceIndex = _getQuizChoiceIndexFromKey(e);
        if (choiceIndex !== -1 && quizState.currentChoices && choiceIndex < quizState.currentChoices.length) {
            e.preventDefault();
            handleMCChoice(choiceIndex);
        }
    }
});

function checkSpelling() {
    if (!quizState || quizState.answered) return;

    const w = quizState.words[quizState.currentIndex];
    const input = document.getElementById('quizSpellingInput').value.trim().toLowerCase();
    const correct = input === w.word.toLowerCase();

    quizState.answered = true;
    quizState.results.push({ wordRef: w, correct, userAnswer: input });
    _showAnswerFeedback(correct, w.word);
}

function handleMCChoice(index) {
    if (!quizState || quizState.answered) return;

    const w = quizState.words[quizState.currentIndex];
    const correct = index === quizState.correctChoiceIndex;

    quizState.answered = true;
    const chosenWord = quizState.currentChoices[index];
    quizState.results.push({ wordRef: w, correct, userAnswer: chosenWord ? chosenWord.word : '' });

    // Visually highlight choices
    const buttons = document.querySelectorAll('.quiz-choice-btn');
    buttons.forEach((btn, i) => {
        if (i === quizState.correctChoiceIndex) {
            btn.classList.add('quiz-choice-correct');
        } else if (i === index && !correct) {
            btn.classList.add('quiz-choice-wrong');
        }
        btn.disabled = true;
    });

    _showAnswerFeedback(correct, w.word);
}

function _showAnswerFeedback(correct, correctWord) {
    const resultArea = document.getElementById('quizResultArea');
    resultArea.style.display = '';
    resultArea.className = 'quiz-result-area ' + (correct ? 'quiz-correct' : 'quiz-wrong');
    resultArea.textContent = correct ? 'Correct!' : 'Incorrect';

    if (!correct) {
        const revealRow = document.getElementById('quizRevealRow');
        const revealWord = document.getElementById('quizRevealWord');
        revealRow.style.display = '';
        revealWord.textContent = correctWord;
        // Replay pronunciation once on wrong answer so user can hear it again immediately.
        setTimeout(() => {
            if (!quizState) return;
            pronounceQuizWord();
        }, 120);
    }

    if (correct) _playDing();

    const skipBtn = document.getElementById('quizSkipBtn');
    if (skipBtn) skipBtn.style.display = 'none';

    document.getElementById('quizNextBtn').style.display = '';
    document.getElementById('quizNextBtn').focus();
}

function nextQuizQuestion() {
    if (!quizState) return;
    quizState.currentIndex++;
    _showQuizQuestion();
}

function skipQuizQuestion() {
    if (!quizState || quizState.answered) return;
    const w = quizState.words[quizState.currentIndex];
    quizState.answered = true;
    quizState.results.push({ wordRef: w, correct: false, userAnswer: '', skipped: true });
    _showAnswerFeedback(false, w.word);
}

function abandonQuiz() {
    document.getElementById('quizQuestionModal').classList.remove('active');
    quizState = null;
}

function _renderQuizEndWordList(listEl, results, cssClass, idPrefix) {
    listEl.innerHTML = results.map((r, i) => {
        const id = `${idPrefix}-${i}`;
        return `<label class="quiz-end-word-row">` +
            `<input type="checkbox" class="quiz-end-checkbox" id="${id}" checked>` +
            `<span class="quiz-end-word ${cssClass}">${r.wordRef.word}</span>` +
            `</label>`;
    }).join('');
}

function _endQuiz() {
    document.getElementById('quizQuestionModal').classList.remove('active');
    if (!quizState) return;

    const total = quizState.results.length;
    const correctCount = quizState.results.filter(r => r.correct).length;
    const wrongCount = total - correctCount;

    document.getElementById('quizEndScore').textContent = `${correctCount} / ${total}`;

    // Wrong words
    const wrongResults = quizState.results.filter(r => !r.correct);
    const wrongSection = document.getElementById('quizEndWrongSection');
    const wrongListEl = document.getElementById('quizEndWrongList');
    const wrongBtn = document.getElementById('quizEndAdjustWrongBtn');
    if (wrongResults.length > 0) {
        _renderQuizEndWordList(wrongListEl, wrongResults, 'quiz-end-word-wrong', 'qw');
        wrongBtn.textContent = `+1 weight for wrong (${wrongCount})`;
        wrongBtn.disabled = false;
        wrongSection.style.display = '';
    } else {
        wrongSection.style.display = 'none';
    }

    // Correct words
    const correctResults = quizState.results.filter(r => r.correct);
    const correctSection = document.getElementById('quizEndCorrectSection');
    const correctListEl = document.getElementById('quizEndCorrectList');
    const correctBtn = document.getElementById('quizEndAdjustCorrectBtn');
    if (correctResults.length > 0) {
        _renderQuizEndWordList(correctListEl, correctResults, 'quiz-end-word-correct', 'qc');
        correctBtn.textContent = `-1 weight for correct (${correctCount})`;
        correctBtn.disabled = false;
        correctSection.style.display = '';
    } else {
        correctSection.style.display = 'none';
    }

    document.getElementById('quizEndModal').classList.add('active');
}

function adjustQuizWeights(delta, type) {
    if (!quizState) return;

    const idPrefix = type === 'wrong' ? 'qw' : 'qc';
    const sourceResults = quizState.results.filter(r => type === 'wrong' ? !r.correct : r.correct);
    let adjusted = 0;

    sourceResults.forEach((result, i) => {
        const checkbox = document.getElementById(`${idPrefix}-${i}`);
        if (checkbox && !checkbox.checked) return; // skip unchecked
        const wordRef = result.wordRef;
        const idx = words.indexOf(wordRef);
        if (idx !== -1) {
            const newWeight = words[idx].weight + delta;
            if (newWeight >= -2 && newWeight <= 10) {
                words[idx].weight = newWeight;
                adjusted++;
            }
        }
    });

    if (adjusted > 0) {
        const sign = delta > 0 ? '+1' : '-1';
        saveData(false, `Quiz ${sign} × ${adjusted}`);
        renderWords();
        showStatus(`Weight ${delta > 0 ? 'increased' : 'decreased'} for ${adjusted} word(s)`, 'success');
    }

    const btnId = type === 'wrong' ? 'quizEndAdjustWrongBtn' : 'quizEndAdjustCorrectBtn';
    const btn = document.getElementById(btnId);
    if (btn) btn.disabled = true;
}

function closeQuizEnd() {
    document.getElementById('quizEndModal').classList.remove('active');
    quizState = null;
}

bindModalBackdropPressReleaseClose(
    document.getElementById('quizSetupModal'),
    closeQuizSetup
);
bindModalBackdropPressReleaseClose(
    document.getElementById('quizEndModal'),
    closeQuizEnd
);

initInPageConfirmModal();

// Registry divider drag-to-resize
(function() {
    const divider = document.querySelector('.registry-divider');
    const treePane = document.getElementById('registryTreePane');
    const body = document.querySelector('.registry-body');
    if (!divider || !treePane || !body) return;

    let isDragging = false;
    let startX = 0;
    let startWidth = 0;

    divider.addEventListener('pointerdown', function(e) {
        isDragging = true;
        startX = e.clientX;
        startWidth = treePane.getBoundingClientRect().width;
        divider.classList.add('is-dragging');
        divider.setPointerCapture(e.pointerId);
        e.preventDefault();
    });

    divider.addEventListener('pointermove', function(e) {
        if (!isDragging) return;
        const dx = e.clientX - startX;
        const bodyWidth = body.getBoundingClientRect().width;
        let newWidth = startWidth + dx;
        // Clamp between 120px and 60% of body
        newWidth = Math.max(120, Math.min(newWidth, bodyWidth * 0.6));
        treePane.style.width = newWidth + 'px';
        treePane.style.maxWidth = 'none';
    });

    divider.addEventListener('pointerup', function(e) {
        if (!isDragging) return;
        isDragging = false;
        divider.classList.remove('is-dragging');
        divider.releasePointerCapture(e.pointerId);
    });

    divider.addEventListener('lostpointercapture', function() {
        isDragging = false;
        divider.classList.remove('is-dragging');
    });

    // Double-click to reset
    divider.addEventListener('dblclick', function() {
        treePane.style.width = '';
        treePane.style.maxWidth = '';
    });
})();
