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

    // Create a new version (creates child version under current version)
    createVersion(data, description = 'Manual save') {
        const newVersion = {
            id: this.generateId(),
            parentId: this.currentId,
            children: [],
            timestamp: new Date().toISOString(),
            data: JSON.parse(JSON.stringify(data)), // Deep copy
            description: description,
            wordCount: data.length,
            lastAccessed: new Date().toISOString()
        };

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

        // Limit number of versions (keep recently accessed)
        if (this.versions.size > this.maxVersions) {
            this.pruneOldVersions();
        }

        this.saveHistory();
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
    deleteSoundEnabled: true,
    audioPreloadEnabled: true
};

function createProjectId() {
    return crypto.randomUUID
        ? crypto.randomUUID()
        : (Date.now().toString(36) + Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2));
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

function sanitizeAppSettings(rawSettings) {
    const parsed = rawSettings && typeof rawSettings === 'object' ? rawSettings : {};
    const deleteSoundEnabled = typeof parsed.deleteSoundEnabled === 'boolean'
        ? parsed.deleteSoundEnabled
        : DEFAULT_APP_SETTINGS.deleteSoundEnabled;
    const audioPreloadEnabled = typeof parsed.audioPreloadEnabled === 'boolean'
        ? parsed.audioPreloadEnabled
        : DEFAULT_APP_SETTINGS.audioPreloadEnabled;

    return {
        deleteSoundEnabled,
        audioPreloadEnabled
    };
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
let wordSortMode = 'alpha'; // 'alpha' or 'chrono'
let cachedVoices = [];
let appSettings = sanitizeAppSettings(null);

// Audio preload cache: word -> { url, buffer (decoded AudioBuffer), status }
const audioCache = new Map();
// status: 'fetching' | 'ready' | 'failed'
let audioPreloadObserver = null;
let deleteSoundAudio = null;

function preloadAudioForWord(word) {
    if (audioCache.has(word)) return;
    audioCache.set(word, { status: 'fetching', url: null, buffer: null });

    fetch(`https://api.dictionaryapi.dev/api/v2/entries/en_GB/${encodeURIComponent(word)}`)
        .then(r => {
            if (!r.ok) throw new Error('API failed');
            return r.json();
        })
        .then(data => {
            const audioUrl = data[0]?.phonetics?.find(p => p.audio)?.audio;
            if (!audioUrl) throw new Error('No audio URL');
            const entry = audioCache.get(word);
            if (!entry) return;
            entry.url = audioUrl;
            // Prefetch and decode the audio buffer
            return fetch(audioUrl)
                .then(r => r.arrayBuffer())
                .then(buf => {
                    const AudioCtx = window.AudioContext || window.webkitAudioContext;
                    if (!AudioCtx) {
                        entry.status = 'ready';
                        return;
                    }
                    const ctx = new AudioCtx();
                    return ctx.decodeAudioData(buf).then(decoded => {
                        entry.buffer = decoded;
                        entry.status = 'ready';
                        ctx.close();
                    }).catch(() => {
                        entry.status = 'ready'; // URL is still valid, just buffer failed
                        ctx.close();
                    });
                });
        })
        .catch(() => {
            const entry = audioCache.get(word);
            if (entry) entry.status = 'failed';
        });
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
                if (word) preloadAudioForWord(word);
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
        return;
    }

    if (audioPreloadObserver) {
        audioPreloadObserver.disconnect();
        audioPreloadObserver = null;
    }
}

function playDeleteSound() {
    if (!appSettings.deleteSoundEnabled) return;

    if (!deleteSoundAudio) {
        deleteSoundAudio = new Audio('assets/delete.mp3');
        deleteSoundAudio.preload = 'auto';
    }

    try {
        deleteSoundAudio.currentTime = 0;
        const playPromise = deleteSoundAudio.play();
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

// Initialize
window.onload = function() {
    loadData();
    appSettings = loadAppSettings();

    // Initialize version control
    versionControl = new VersionControl(50);

    // Create initial version if no versions exist
    if (versionControl.versions.size === 0 && words.length > 0) {
        versionControl.createVersion(words, 'Initial version');
    } else if (versionControl.versions.size === 0 && words.length === 0) {
        versionControl.createVersion([], 'Initial empty state');
    }

    if (appSettings.deleteSoundEnabled) {
        deleteSoundAudio = new Audio('assets/delete.mp3');
        deleteSoundAudio.preload = 'auto';
    }

    renderWords();
    updateWeightSelection();
    updateEditWeightSelection();

    // Set default date to today
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('dateInput').value = today;

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

    Dropdown.register(posDropdown);
    Dropdown.register(weightDropdown);
    Dropdown.register(editPosDropdown);
    Dropdown.register(editWeightDropdown);

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

function toggleMode() {
    modeToggle.classList.toggle('active');
    hideMeaning = !hideMeaning;
    renderWords();
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

function togglePosDropdown() {
    if (posDropdown) {
        posDropdown.toggle();
    }
}

function renderPosSelection(containerId, values) {
    const container = document.getElementById(containerId);
    if (values.length === 0) {
        container.innerHTML = '<span class="pos-placeholder">POS</span>';
    } else {
        container.innerHTML = values.map(p => `<span class="pos-tag">${p}</span>`).join('');
    }
}

function syncPosOptionState(dropdownId, values) {
    const selectedValues = new Set(values);
    document.querySelectorAll(`#${dropdownId} .pos-option`).forEach(option => {
        option.classList.toggle('selected', selectedValues.has(option.dataset.value));
    });
}

function updatePosSelection() {
    renderPosSelection('posSelected', selectedPos);
    syncPosOptionState('posDropdown', selectedPos);
}

function togglePosOption(option, event) {
    if (event) {
        event.stopPropagation();
    }

    const value = option.dataset.value;
    if (!value) return;

    if (selectedPos.includes(value)) {
        selectedPos = selectedPos.filter(pos => pos !== value);
    } else {
        selectedPos = [...selectedPos, value];
    }

    updatePosSelection();
}

// POS selection for edit form
let editSelectedPos = [];

function toggleEditPosDropdown() {
    if (editPosDropdown) {
        editPosDropdown.toggle();
    }
}

function updateEditPosSelection() {
    renderPosSelection('editPosSelected', editSelectedPos);
    syncPosOptionState('editPosDropdown', editSelectedPos);
}

function toggleEditPosOption(option, event) {
    if (event) {
        event.stopPropagation();
    }

    const value = option.dataset.value;
    if (!value) return;

    if (editSelectedPos.includes(value)) {
        editSelectedPos = editSelectedPos.filter(pos => pos !== value);
    } else {
        editSelectedPos = [...editSelectedPos, value];
    }

    updateEditPosSelection();
}

function syncWeightOptionState(dropdownId, value) {
    document.querySelectorAll(`#${dropdownId} .weight-option`).forEach(option => {
        const optionValue = parseInt(option.dataset.value, 10);
        option.classList.toggle('selected', optionValue === value);
    });
}

function updateWeightSelection() {
    document.getElementById('weightSelected').textContent = String(selectedWeight);
    syncWeightOptionState('weightDropdown', selectedWeight);
}

function updateEditWeightSelection() {
    document.getElementById('editWeightSelected').textContent = String(editSelectedWeight);
    syncWeightOptionState('editWeightDropdown', editSelectedWeight);
}

function toggleWeightDropdown() {
    if (weightDropdown) {
        weightDropdown.toggle();
    }
}

function toggleEditWeightDropdown() {
    if (editWeightDropdown) {
        editWeightDropdown.toggle();
    }
}

function setWeightOption(option, event) {
    if (event) {
        event.stopPropagation();
    }

    const value = parseInt(option.dataset.value, 10);
    if (Number.isNaN(value)) return;
    selectedWeight = value;
    updateWeightSelection();
    if (weightDropdown) {
        weightDropdown.close();
    }
}

function setEditWeightOption(option, event) {
    if (event) {
        event.stopPropagation();
    }

    const value = parseInt(option.dataset.value, 10);
    if (Number.isNaN(value)) return;
    editSelectedWeight = value;
    updateEditWeightSelection();
    if (editWeightDropdown) {
        editWeightDropdown.close();
    }
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

    if (!word) {
        alert('Please fill in word');
        return;
    }

    const newWord = {
        word: word.toLowerCase(),
        meaning: meaning,
        pos: selectedPos.slice(), // Copy array
        weight: weight,
        added: date
    };

    words.push(newWord);
    saveData(false, `＋　${word}`);
    renderWords();
    showStatus(`Added "${word}"`, 'success');

    // Clear form
    document.getElementById('wordInput').value = '';
    document.getElementById('meaningInput').value = '';
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
    playDeleteSound();
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
    wordSortMode = wordSortMode === 'alpha' ? 'chrono' : 'alpha';
    const btn = document.getElementById('sortModeBtn');
    if (btn) btn.textContent = wordSortMode === 'alpha' ? 'A-Z' : 'Date';
    renderWords();
}

// Toggle select mode
function toggleSelectMode() {
    isSelectMode = !isSelectMode;
    if (!isSelectMode) {
        selectedWords.clear();
    }
    renderWords();
    updateBatchToolbar();
}

// Update batch toolbar and buttons
function updateBatchToolbar() {
    const toolbar = document.getElementById('batchToolbar');
    const selectModeBtn = document.getElementById('selectModeBtn');
    const batchUpBtn = document.getElementById('batchUpBtn');
    const batchDownBtn = document.getElementById('batchDownBtn');
    const batchDeleteBtn = document.getElementById('batchDeleteBtn');

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

// Add weight range to selection
function addWeightRange() {
    const range = getWeightRange();
    if (!range) {
        showStatus('Invalid weight range', 'error');
        return;
    }

    let count = 0;
    words.forEach((w, index) => {
        if (w.weight >= range.min && w.weight <= range.max) {
            selectedWords.add(index);
            count++;
        }
    });

    updateBatchToolbar();
    updateWordSelectionUI();
    showStatus(`+${count} word(s)`, 'success');
}

// Remove weight range from selection
function removeWeightRange() {
    const range = getWeightRange();
    if (!range) {
        showStatus('Invalid weight range', 'error');
        return;
    }

    let count = 0;
    words.forEach((w, index) => {
        if (w.weight >= range.min && w.weight <= range.max) {
            if (selectedWords.has(index)) {
                selectedWords.delete(index);
                count++;
            }
        }
    });

    updateBatchToolbar();
    updateWordSelectionUI();
    showStatus(`−${count} word(s)`, 'success');
}

// Parse date range inputs
function getDateRange() {
    const minVal = document.getElementById('dateMinInput').value;
    const maxVal = document.getElementById('dateMaxInput').value;

    // At least one date must be provided
    if (!minVal && !maxVal) return null;

    return {
        min: minVal || null,
        max: maxVal || null
    };
}

// Add date range to selection
function addDateRange() {
    const range = getDateRange();
    if (!range) {
        showStatus('Please set at least one date', 'error');
        return;
    }

    let count = 0;
    words.forEach((w, index) => {
        const date = w.added;
        if ((!range.min || date >= range.min) && (!range.max || date <= range.max)) {
            selectedWords.add(index);
            count++;
        }
    });

    updateBatchToolbar();
    updateWordSelectionUI();
    showStatus(`+${count} word(s)`, 'success');
}

// Remove date range from selection
function removeDateRange() {
    const range = getDateRange();
    if (!range) {
        showStatus('Please set at least one date', 'error');
        return;
    }

    let count = 0;
    words.forEach((w, index) => {
        const date = w.added;
        if ((!range.min || date >= range.min) && (!range.max || date <= range.max)) {
            if (selectedWords.has(index)) {
                selectedWords.delete(index);
                count++;
            }
        }
    });

    updateBatchToolbar();
    updateWordSelectionUI();
    showStatus(`−${count} word(s)`, 'success');
}

// Select by regex (add matching words to selection)
function selectByRegex() {
    const pattern = document.getElementById('regexFilterInput').value.trim();
    if (!pattern) {
        showStatus('Please enter a regex pattern', 'error');
        return;
    }

    try {
        const regex = new RegExp(pattern, 'i');
        let matchCount = 0;

        words.forEach((w, index) => {
            if (regex.test(w.word) || regex.test(w.meaning)) {
                selectedWords.add(index);
                matchCount++;
            }
        });

        updateBatchToolbar();
        updateWordSelectionUI();
    
        showStatus(`Matched ${matchCount} word(s)`, 'success');
    } catch (e) {
        showStatus('Invalid regex pattern', 'error');
    }
}

// Deselect by regex (remove matching words from selection)
function deselectByRegex() {
    const pattern = document.getElementById('regexFilterInput').value.trim();
    if (!pattern) {
        showStatus('Please enter a regex pattern', 'error');
        return;
    }

    try {
        const regex = new RegExp(pattern, 'i');
        let matchCount = 0;

        words.forEach((w, index) => {
            if (regex.test(w.word) || regex.test(w.meaning)) {
                selectedWords.delete(index);
                matchCount++;
            }
        });

        updateBatchToolbar();
        updateWordSelectionUI();
    
        showStatus(`Unmatched ${matchCount} word(s)`, 'success');
    } catch (e) {
        showStatus('Invalid regex pattern', 'error');
    }
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
    playDeleteSound();
    renderWords();
    updateBatchToolbar();
    showStatus(`Deleted ${count} word(s)`, 'success');
}

// Pronunciation — uses preloaded audio cache for instant playback
function pronounceWord(word) {
    const cached = audioCache.get(word);

    // If cache has a decoded buffer, play it instantly via AudioContext
    if (cached && cached.status === 'ready' && cached.buffer) {
        const AudioCtx = window.AudioContext || window.webkitAudioContext;
        if (AudioCtx) {
            const ctx = new AudioCtx();
            const source = ctx.createBufferSource();
            source.buffer = cached.buffer;
            source.connect(ctx.destination);
            source.start(0);
            source.onended = () => ctx.close();
            showStatus(`Playing "${word}"`, 'info');
            return;
        }
    }

    // If cache has a URL but no buffer, use Audio element
    if (cached && cached.status === 'ready' && cached.url) {
        const audio = new Audio(cached.url);
        audio.play().then(() => {
            showStatus(`Playing "${word}"`, 'info');
        }).catch(() => pronounceFallback(word));
        return;
    }

    // If cache failed or not yet loaded, fetch on demand
    const AudioCtx = window.AudioContext || window.webkitAudioContext;
    const audioCtx = AudioCtx ? new AudioCtx() : null;

    fetch(`https://api.dictionaryapi.dev/api/v2/entries/en_GB/${encodeURIComponent(word)}`)
        .then(response => {
            if (!response.ok) throw new Error('API failed');
            return response.json();
        })
        .then(data => {
            const audioUrl = data[0]?.phonetics?.find(p => p.audio)?.audio;
            if (!audioUrl) throw new Error('No audio URL');

            // Store in cache for future use
            if (!audioCache.has(word)) audioCache.set(word, { status: 'ready', url: audioUrl, buffer: null });
            else { const e = audioCache.get(word); e.url = audioUrl; e.status = 'ready'; }

            if (audioCtx) {
                return fetch(audioUrl)
                    .then(r => r.arrayBuffer())
                    .then(buf => audioCtx.decodeAudioData(buf))
                    .then(decoded => {
                        // Cache the decoded buffer
                        const entry = audioCache.get(word);
                        if (entry) entry.buffer = decoded;
                        const source = audioCtx.createBufferSource();
                        source.buffer = decoded;
                        source.connect(audioCtx.destination);
                        source.start(0);
                        source.onended = () => audioCtx.close();
                        showStatus(`Playing "${word}"`, 'info');
                    });
            } else {
                const audio = new Audio(audioUrl);
                return audio.play().then(() => {
                    showStatus(`Playing "${word}"`, 'info');
                });
            }
        })
        .catch(() => {
            if (audioCtx) audioCtx.close().catch(() => {});
            pronounceFallback(word);
        });
}

function pronounceFallback(word) {
    if (!('speechSynthesis' in window)) {
        showStatus('Pronunciation not supported', 'error');
        return;
    }

    // iOS fix: cancel then delay before speak — immediate speak after cancel silently fails
    speechSynthesis.cancel();

    const doSpeak = () => {
        const utterance = new SpeechSynthesisUtterance(word);
        utterance.lang = 'en-GB';
        utterance.rate = 0.8;

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
        utterance.onend = cleanup;
        utterance.onerror = (e) => {
            cleanup();
            // 'interrupted' is normal (user clicked again), only log real errors
            if (e.error !== 'interrupted') {
                showStatus('Pronunciation failed', 'error');
            }
        };

        speechSynthesis.speak(utterance);
        showStatus(`Pronouncing "${word}"`, 'info');
    };

    // iOS needs ~100ms gap after cancel() before speak() will work
    setTimeout(doSpeak, 100);
}

// Open edit modal
function openEditModal(index) {
    editingIndex = index;
    const word = words[index];

    document.getElementById('editWordInput').value = word.word;
    document.getElementById('editMeaningInput').value = word.meaning;
    document.getElementById('editDateInput').value = word.added;

    // Set POS selection
    editSelectedPos = Array.isArray(word.pos) ? word.pos.slice() : (word.pos ? [word.pos] : []);
    updateEditPosSelection();
    editSelectedWeight = Number.isInteger(word.weight) ? word.weight : parseInt(word.weight, 10);
    if (Number.isNaN(editSelectedWeight)) {
        editSelectedWeight = 3;
    }
    updateEditWeightSelection();

    document.getElementById('editModal').classList.add('active');
}

// Close edit modal
function closeEditModal() {
    editingIndex = -1;
    document.getElementById('editModal').classList.remove('active');
}

// Save edit
function saveEdit() {
    if (editingIndex === -1) return;

    const word = document.getElementById('editWordInput').value.trim();
    const meaning = document.getElementById('editMeaningInput').value.trim();
    const weight = editSelectedWeight;
    const date = document.getElementById('editDateInput').value;

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
        added: date
    };

    saveData(false, `✎　${oldWord}`);
    renderWords();
    closeEditModal();
}

// Close modal on background click
document.getElementById('editModal').addEventListener('click', function(e) {
    if (e.target === this) {
        closeEditModal();
    }
});

// Render words
function renderWords() {
    const container = document.getElementById('wordList');

    // Show all words
    let filteredWords = words.filter(w => w.weight >= -3);

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

    // Group by weight
    const groups = {};
    filteredWords.forEach(w => {
        if (!groups[w.weight]) {
            groups[w.weight] = [];
        }
        groups[w.weight].push(w);
    });

    // Sort within groups
    Object.keys(groups).forEach(weight => {
        if (wordSortMode === 'chrono') {
            groups[weight].sort((a, b) => (a.added || '').localeCompare(b.added || ''));
        } else {
            groups[weight].sort((a, b) => a.word.localeCompare(b.word));
        }
    });

    // Sort groups by weight
    const sortedWeights = Object.keys(groups).map(Number).sort((a, b) => b - a);

    let html = '';
    sortedWeights.forEach(weight => {
        const groupWords = groups[weight];
        const isNegative = weight < 0;
        const groupLabel = getWeightLabel(weight);
        const expandedClass = isNegative ? '' : 'expanded';
        const collapseIcon = isNegative ? '▼' : '▲';

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
            const masteredClass = (weight < 0 && !isInvalid) ? 'mastered' : '';
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
                    <div class="word-meta">Added: ${formatAddedDateLabel(w.added)}</div>
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
// Export data only (without version history)
function exportDataOnly() {
    const exportObj = {
        projectId: getProjectId(),
        words: words
    };
    const dataStr = JSON.stringify(exportObj, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `words-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showStatus('Exported data only', 'success');
}

// Export data with complete version history
function exportWithVersionHistory() {
    if (!versionControl) {
        showStatus('Version control not initialized', 'error');
        return;
    }

    const exportData = {
        projectId: getProjectId(),
        words: words,
        versionHistory: {
            format: 'tree-v1',
            versions: Object.fromEntries(versionControl.versions),
            rootId: versionControl.rootId,
            currentId: versionControl.currentId,
            exportDate: new Date().toISOString()
        }
    };

    const dataStr = JSON.stringify(exportData, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `words-with-history-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showStatus('Exported data with version history', 'success');
}

// Legacy export function (for backward compatibility)
function exportData() {
    exportDataOnly();
}

// Validate version history structure
function validateVersionHistory(versionHistory) {
    if (!versionHistory.format || versionHistory.format !== 'tree-v1') {
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
                added: item.added || new Date().toISOString().split('T')[0]
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
                added: item.added || new Date().toISOString().split('T')[0]
            };
        }

        // Valid word
        validCount++;
        return {
            word: item.word.toLowerCase(),
            meaning: item.meaning || '',
            pos: posArray,
            weight: weight,
            added: item.added || new Date().toISOString().split('T')[0]
        };
    });

    return { processed, validCount, invalidCount };
}

// Import words only (clear version history)
function importWordsOnly(wordsData) {
    const { processed, validCount, invalidCount } = processImportedWords(wordsData);
    words = processed;

    // Clear version history and create new initial version
    if (versionControl) {
        versionControl.clearHistory();
    }

    saveData(false, `Imported ${processed.length} word(s)`);
    renderWords();

    if (invalidCount > 0) {
        showStatus(`Imported ${validCount} valid, ${invalidCount} invalid words (version history cleared)`, 'success');
    } else {
        showStatus(`Imported ${processed.length} words (version history cleared)`, 'success');
    }
}

// Import as overwrite: use imported data, create a new branch in registry-tree
function importAsOverwrite(wordsData, description) {
    const { processed } = processImportedWords(wordsData);
    words = processed;
    localStorage.setItem('wordMemoryData', JSON.stringify(words));

    // Create a new version branching from current
    if (versionControl) {
        versionControl.createVersion(words, description || `Overwrite import (${processed.length} words)`);
    }

    renderWords();
    showStatus(`Overwrite imported ${processed.length} words as new branch`, 'success');
}

// Import data with version history
function importWithVersionHistory(importedData) {
    const { processed, validCount, invalidCount } = processImportedWords(importedData.words);
    words = processed;

    // Import version history
    versionControl.versions = new Map(Object.entries(importedData.versionHistory.versions));
    versionControl.rootId = importedData.versionHistory.rootId;
    versionControl.currentId = importedData.versionHistory.currentId;
    versionControl.saveHistory();

    // Switch to current version
    const currentVersion = versionControl.versions.get(versionControl.currentId);
    if (currentVersion) {
        words = JSON.parse(JSON.stringify(currentVersion.data));
    }

    localStorage.setItem('wordMemoryData', JSON.stringify(words));
    renderWords();

    const versionCount = versionControl.versions.size;
    if (invalidCount > 0) {
        showStatus(`Imported ${validCount} valid, ${invalidCount} invalid words with ${versionCount} version(s)`, 'success');
    } else {
        showStatus(`Imported ${processed.length} words with ${versionCount} version(s)`, 'success');
    }
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
        const localProjectId = getProjectId();
        const importedProjectId = imported && typeof imported === 'object'
            ? imported.projectId || null
            : null;
        const isSameProject = importedProjectId && importedProjectId === localProjectId;

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
                    importWordsOnly(imported.words || imported);
                }
                return;
            }

            if (isSameProject) {
                // Same project: ask overwrite (creates new branch) or full replace
                const shouldOverwriteSameProject = await showInPageConfirm({
                    title: 'Same Project Detected',
                    message: 'Overwrite current data? (imported data will be added as a new branch in version history)',
                    confirmText: 'Overwrite',
                    cancelText: 'Cancel'
                });
                if (shouldOverwriteSameProject) {
                    importAsOverwrite(imported.words, `Import overwrite (${imported.words.length} words)`);
                }
            } else {
                const shouldImportWithHistory = await showInPageConfirm({
                    title: 'Import With History',
                    message: `Import ${imported.words.length} words with ${Object.keys(imported.versionHistory.versions).length} version(s)? This will overwrite current data and version history.`,
                    confirmText: 'Import',
                    cancelText: 'Cancel'
                });
                if (shouldImportWithHistory) {
                    importWithVersionHistory(imported);
                    // Adopt the imported project's ID
                    if (importedProjectId) {
                        setProjectId(importedProjectId);
                    }
                }
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
            if (isSameProject) {
                const shouldOverwriteWordsOnly = await showInPageConfirm({
                    title: 'Same Project Detected',
                    message: 'Overwrite current data? (imported data will be added as a new branch in version history)',
                    confirmText: 'Overwrite',
                    cancelText: 'Cancel'
                });
                if (shouldOverwriteWordsOnly) {
                    importAsOverwrite(imported.words, `Import overwrite (${imported.words.length} words)`);
                }
            } else {
                const shouldImportWords = await showInPageConfirm({
                    title: 'Import Words',
                    message: `Import ${imported.words.length} words? This will overwrite current data and clear version history.`,
                    confirmText: 'Import',
                    cancelText: 'Cancel'
                });
                if (shouldImportWords) {
                    importWordsOnly(imported.words);
                    if (importedProjectId) {
                        setProjectId(importedProjectId);
                    }
                }
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

// Undo/Redo functions
function performUndo() {
    if (!versionControl) {
        showStatus('Version control not initialized', 'error');
        return;
    }

    if (!versionControl.canUndo()) {
        showStatus('Nothing to undo', 'info');
        return;
    }

    const version = versionControl.undo();
    if (version) {
        words = JSON.parse(JSON.stringify(version.data)); // Deep copy
        localStorage.setItem('wordMemoryData', JSON.stringify(words));
        renderWords();
        showStatus(`Undo: ${version.description}`, 'success');
    }
}

function performRedo() {
    if (!versionControl) {
        showStatus('Version control not initialized', 'error');
        return;
    }

    if (!versionControl.canRedo()) {
        showStatus('Nothing to redo', 'info');
        return;
    }

    const version = versionControl.redo();
    if (version) {
        words = JSON.parse(JSON.stringify(version.data)); // Deep copy
        localStorage.setItem('wordMemoryData', JSON.stringify(words));
        renderWords();
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
function openSettingsModal() {
    if (!versionControl) return;

    const settings = versionControl.getSettings();
    const deleteSoundEnabledInput = document.getElementById('deleteSoundEnabledInput');
    const audioPreloadEnabledInput = document.getElementById('audioPreloadEnabledInput');
    document.getElementById('maxVersionsInput').value = settings.maxVersions;
    document.getElementById('projectIdInput').value = getProjectId();
    if (deleteSoundEnabledInput) {
        deleteSoundEnabledInput.checked = appSettings.deleteSoundEnabled;
    }
    if (audioPreloadEnabledInput) {
        audioPreloadEnabledInput.checked = appSettings.audioPreloadEnabled;
    }
    document.getElementById('settingsModal').classList.add('active');
}

function closeSettingsModal() {
    document.getElementById('settingsModal').classList.remove('active');
}

function saveSettings() {
    if (!versionControl) return;

    const maxVersions = parseInt(document.getElementById('maxVersionsInput').value, 10);
    const projectIdInput = document.getElementById('projectIdInput');
    const deleteSoundEnabledInput = document.getElementById('deleteSoundEnabledInput');
    const audioPreloadEnabledInput = document.getElementById('audioPreloadEnabledInput');
    const projectId = projectIdInput ? projectIdInput.value.trim() : '';
    const nextDeleteSoundEnabled = deleteSoundEnabledInput ? deleteSoundEnabledInput.checked : DEFAULT_APP_SETTINGS.deleteSoundEnabled;
    const nextAudioPreloadEnabled = audioPreloadEnabledInput ? audioPreloadEnabledInput.checked : DEFAULT_APP_SETTINGS.audioPreloadEnabled;

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
        deleteSoundEnabled: nextDeleteSoundEnabled,
        audioPreloadEnabled: nextAudioPreloadEnabled
    });
    saveAppSettings();
    applyAudioPreloadSetting();

    if (!appSettings.deleteSoundEnabled) {
        deleteSoundAudio = null;
    } else if (!deleteSoundAudio) {
        deleteSoundAudio = new Audio('assets/delete.mp3');
        deleteSoundAudio.preload = 'auto';
    }

    showStatus('Settings saved', 'success');
    closeSettingsModal();
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
    modal.addEventListener('click', function(e) {
        if (e.target === modal) {
            resolveInPageConfirm(false);
        }
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
        : (Array.isArray(version.data) ? version.data.length : 0);
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
            words = JSON.parse(JSON.stringify(currentVersion.data));
            localStorage.setItem('wordMemoryData', JSON.stringify(words));
            renderWords();
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
    playDeleteSound();

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

    content.innerHTML = renderJsonHighlight(version.data);
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
        container.innerHTML = renderDiffView(version.data, currentVersion.data, version, currentVersion);
    } else {
        container.innerHTML = renderJsonHighlight(version.data);
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
        words = JSON.parse(JSON.stringify(version.data));
        localStorage.setItem('wordMemoryData', JSON.stringify(words));
        renderWords();
        renderRegistryTree();
        renderRegistryPreview();
        updateRegistryStatus();
        updateRegistryAddressBar();
        showStatus(`Jumped to version: ${version.description}`, 'success');
    }
}

// Close modals on background click
document.getElementById('settingsModal').addEventListener('click', function(e) {
    if (e.target === this) {
        closeSettingsModal();
    }
});

document.getElementById('historyModal').addEventListener('click', function(e) {
    if (e.target === this) {
        closeHistoryModal();
    }
});

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
