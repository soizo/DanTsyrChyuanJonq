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
                    console.log('Migrating from linear to tree structure...');
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
            console.error('Failed to load version history:', error);
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
        console.log(`Migrated ${linearVersions.length} versions to tree structure`);
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
            console.error('Failed to save version history:', error);
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
                    console.error('Still failed after removing old versions:', e);
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

// Data storage
let words = [];
let isFullMode = false;
let editingIndex = -1;
let selectedWords = new Set();
let isSelectMode = false;
let versionControl = null;

const SPEAKER_ICON_SVG = '<svg class="icon-speaker" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path fill="currentColor" d="M3 10v4c0 1.1.9 2 2 2h2.35l3.38 2.7c.93.74 2.27.08 2.27-1.1V6.4c0-1.18-1.34-1.84-2.27-1.1L7.35 8H5c-1.1 0-2 .9-2 2Zm14.5 2c0-1.77-1.02-3.29-2.5-4.03v8.06c1.48-.74 2.5-2.26 2.5-4.03ZM15 3.23v2.06c2.89.86 5 3.54 5 6.71s-2.11 5.85-5 6.71v2.06c4.01-.91 7-4.49 7-8.77s-2.99-7.86-7-8.77Z"/></svg>';

// Dropdown Base Class
class Dropdown {
    constructor(containerId, selectedId, dropdownId) {
        this.container = document.getElementById(containerId);
        this.selected = document.getElementById(selectedId);
        this.dropdown = document.getElementById(dropdownId);
        this.isOpen = false;

        // Check if all elements exist
        if (!this.container || !this.selected || !this.dropdown) {
            console.warn(`Dropdown initialization failed: ${containerId}, ${selectedId}, ${dropdownId}`);
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

// Initialize
window.onload = function() {
    loadData();

    // Initialize version control
    versionControl = new VersionControl(50);

    // Create initial version if no versions exist
    if (versionControl.versions.size === 0 && words.length > 0) {
        versionControl.createVersion(words, 'Initial version');
    } else if (versionControl.versions.size === 0 && words.length === 0) {
        versionControl.createVersion([], 'Initial empty state');
    }

    renderWords();
    updateWeightSelection();
    updateEditWeightSelection();

    // Set default date to today
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('dateInput').value = today;

    // Load voices for speech synthesis
    if ('speechSynthesis' in window) {
        speechSynthesis.getVoices();
        speechSynthesis.onvoiceschanged = () => {
            speechSynthesis.getVoices();
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
};

// Mode toggle
const modeToggle = document.getElementById('modeToggle');

function toggleMode() {
    modeToggle.classList.toggle('active');
    isFullMode = !isFullMode;
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
        container.innerHTML = '<span class="pos-placeholder">Select...</span>';
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
    saveData(false, `Added word "${word}"`);
    renderWords();
    showStatus(`Added "${word}"`, 'success');

    // Clear form
    document.getElementById('wordInput').value = '';
    document.getElementById('meaningInput').value = '';
    selectedPos = [];
    updatePosSelection();
    selectedWeight = 3;
    updateWeightSelection();
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('dateInput').value = today;
}

// Update weight
function updateWeight(index, delta) {
    const currentWeight = words[index].weight;
    const wordName = words[index].word;

    // If current weight is -3 (invalid), delta is the new weight directly
    if (currentWeight === -3) {
        words[index].weight = delta;
        saveData(false, `Fixed weight for "${wordName}"`);
        renderWords();
        showStatus(`Fixed weight to ${delta}`, 'success');
        return;
    }

    const newWeight = currentWeight + delta;

    if (newWeight < -2 || newWeight > 10) {
        return;
    }

    words[index].weight = newWeight;
    saveData(false, `Updated weight for "${wordName}"`);
    renderWords();
}

// Delete word
function deleteWord(index) {
    const wordName = words[index].word;
    if (confirm(`Delete "${wordName}"?`)) {
        words.splice(index, 1);
        saveData(false, `Deleted word "${wordName}"`);
        renderWords();
        showStatus('Word deleted', 'success');
    }
}

// Toggle word selection
function toggleWordSelection(index) {
    if (selectedWords.has(index)) {
        selectedWords.delete(index);
    } else {
        selectedWords.add(index);
    }
    updateBatchToolbar();
    updateWordSelectionUI();

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
}

// Select all words
function selectAll() {
    const filteredWords = words.filter(w => {
        if (isFullMode) {
            return w.weight >= -3;
        } else {
            return w.weight >= 0;
        }
    });

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
    const filteredWords = words.filter(w => {
        if (isFullMode) {
            return w.weight >= -3;
        } else {
            return w.weight >= 0;
        }
    });

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
function batchAdjustWeight(delta) {
    if (selectedWords.size === 0) return;

    const count = selectedWords.size;
    const action = delta > 0 ? 'increased' : 'decreased';

    if (confirm(`${delta > 0 ? 'increase' : 'decrease'} weight for ${count} selected word(s)?`)) {
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
function batchDeleteConfirm() {
    if (selectedWords.size === 0) return;

    const count = selectedWords.size;
    if (confirm(`Delete ${count} selected word(s)?`)) {
        const indicesToDelete = Array.from(selectedWords).sort((a, b) => b - a);
        indicesToDelete.forEach(index => {
            words.splice(index, 1);
        });
        selectedWords.clear();
        saveData(false, `Deleted ${count} word(s)`);
        renderWords();
        updateBatchToolbar();
        showStatus(`Deleted ${count} word(s)`, 'success');
    }
}

// Pronunciation
async function pronounceWord(word) {
    try {
        // Try Free Dictionary API first for better quality
        const response = await fetch(`https://api.dictionaryapi.dev/api/v2/entries/en_GB/${word}`);
        if (response.ok) {
            const data = await response.json();
            const audioUrl = data[0]?.phonetics?.find(p => p.audio)?.audio;

            if (audioUrl) {
                const audio = new Audio(audioUrl);
                audio.play();
                showStatus(`Playing "${word}"`, 'info');
                return;
            }
        }
    } catch (error) {
        console.log('Dictionary API failed, using Web Speech API');
    }

    // Fallback to Web Speech API with British English
    if ('speechSynthesis' in window) {
        const utterance = new SpeechSynthesisUtterance(word);
        utterance.lang = 'en-GB'; // British English
        utterance.rate = 0.8; // Slightly slower for clarity

        // Try to find a British English voice
        const voices = speechSynthesis.getVoices();
        const britishVoice = voices.find(voice =>
            voice.lang.startsWith('en-GB') || voice.lang.startsWith('en-UK')
        );

        if (britishVoice) {
            utterance.voice = britishVoice;
        }

        speechSynthesis.speak(utterance);
        showStatus(`Pronouncing "${word}"`, 'info');
    } else {
        showStatus('Pronunciation not supported', 'error');
    }
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

    saveData(false, `Edited word "${oldWord}"`);
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

    // Filter words
    let filteredWords = words.filter(w => {
        if (isFullMode) {
            return w.weight >= -3; // Show -3, -2, -1, 0, 1, 2, 3, 4, 5+
        } else {
            return w.weight >= 0; // Show only 0, 1, 2, 3, 4, 5+
        }
    });

    if (filteredWords.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">—</div>
                <div>${isFullMode ? 'No Words' : 'No Words to Review'}</div>
            </div>
        `;
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
        groups[weight].sort((a, b) => a.word.localeCompare(b.word));
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
                <div class="word-item ${masteredClass} ${invalidClass}">
                    <div class="word-header">
                        <div style="display: flex; align-items: center; gap: 8px;">
                            ${isSelectMode ? `<input type="checkbox" id="select-${originalIndex}" class="word-checkbox" onchange="toggleWordSelection(${originalIndex})" ${selectedWords.has(originalIndex) ? 'checked' : ''}>` : ''}
                            <div>
                                <span class="word-title">${w.word}</span>
                                ${posTags}
                                <button class="btn-pronounce" onclick="pronounceWord('${w.word}')" title="Pronounce (British)" aria-label="Pronounce (British)">${SPEAKER_ICON_SVG}</button>
                            </div>
                        </div>
                        <div class="word-weight ${weightShapeClass}">${weightDisplay}</div>
                    </div>
                    <div class="word-meaning">${w.meaning}</div>
                    <div class="word-meta">Added: ${w.added}</div>
                    <div class="word-actions">
                        ${w.weight >= 0 ? `<button class="btn-remember" onclick="updateWeight(${originalIndex}, -1)">Down</button>` : ''}
                        ${w.weight >= -1 && w.weight < 10 ? `<button class="btn-forget" onclick="updateWeight(${originalIndex}, 1)">Up</button>` : ''}
                        ${isInvalid ? `<button class="btn-secondary" onclick="updateWeight(${originalIndex}, 3)">Fix</button>` : ''}
                        <button class="btn-edit" onclick="openEditModal(${originalIndex})">Edit</button>
                        <button class="btn-delete" onclick="deleteWord(${originalIndex})">Del</button>
                    </div>
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
    const dataStr = JSON.stringify(words, null, 2);
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

// Import data (enhanced with version history support)
function importData(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const imported = JSON.parse(e.target.result);

            // Check if it contains version history
            if (imported.versionHistory) {
                const validation = validateVersionHistory(imported.versionHistory);

                if (!validation.valid) {
                    // Version history is corrupted, ask user
                    if (confirm(`Version history is corrupted: ${validation.error}\n\nContinue importing data only (version history will be cleared)?`)) {
                        importWordsOnly(imported.words || imported);
                    }
                    event.target.value = '';
                    return;
                }

                // Version history is valid, ask user to confirm full import
                if (confirm(`Import ${imported.words.length} words with ${Object.keys(imported.versionHistory.versions).length} version(s)? This will overwrite current data and version history.`)) {
                    importWithVersionHistory(imported);
                }
            } else if (Array.isArray(imported)) {
                // Old format - words array only
                if (confirm(`Import ${imported.length} words? This will overwrite current data and clear version history.`)) {
                    importWordsOnly(imported);
                }
            } else if (imported.words && Array.isArray(imported.words)) {
                // Object with words array but no version history
                if (confirm(`Import ${imported.words.length} words? This will overwrite current data and clear version history.`)) {
                    importWordsOnly(imported.words);
                }
            } else {
                showStatus('Invalid file format', 'error');
            }
        } catch (error) {
            showStatus('Failed to parse file', 'error');
        }
        event.target.value = '';
    };
    reader.readAsText(file);
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
    document.getElementById('maxVersionsInput').value = settings.maxVersions;
    document.getElementById('settingsModal').classList.add('active');
}

function closeSettingsModal() {
    document.getElementById('settingsModal').classList.remove('active');
}

function saveSettings() {
    if (!versionControl) return;

    const maxVersions = parseInt(document.getElementById('maxVersionsInput').value, 10);

    if (isNaN(maxVersions) || maxVersions < 10 || maxVersions > 200) {
        alert('Max versions must be between 10 and 200');
        return;
    }

    versionControl.updateSettings({ maxVersions: maxVersions });
    showStatus('Settings saved', 'success');
    closeSettingsModal();
}

// Version history functions
function openHistoryModal() {
    if (!versionControl) return;

    renderVersionHistory();
    closeSettingsModal();
    document.getElementById('historyModal').classList.add('active');
}

function closeHistoryModal() {
    document.getElementById('historyModal').classList.remove('active');
}

function renderVersionHistory() {
    if (!versionControl) return;

    const tree = versionControl.getVersionTree();
    const container = document.getElementById('historyList');

    // Update version info with tree structure
    const currentVersion = versionControl.getCurrentVersion();
    const totalVersions = versionControl.versions.size;
    document.getElementById('currentVersionIndex').textContent = currentVersion ? '✓' : '-';
    document.getElementById('totalVersions').textContent = totalVersions;
    document.getElementById('canUndo').textContent = versionControl.canUndo() ? 'Yes' : 'No';
    document.getElementById('canRedo').textContent = versionControl.canRedo() ? 'Yes' : 'No';

    if (tree.length === 0) {
        container.innerHTML = '<div class="empty-state"><div class="empty-state-icon">—</div><div>No Version History</div></div>';
        return;
    }

    let html = '';
    tree.forEach(version => {
        const isCurrent = version.isCurrent;
        const hasMultipleBranches = version.children.length > 1;

        // Create tree prefix with indentation
        const indent = '  '.repeat(version.depth); // 2 spaces per depth level
        let branchIcon = '';

        if (version.depth > 0) {
            if (hasMultipleBranches) {
                branchIcon = '┬'; // Has multiple branches
            } else if (version.hasChildren) {
                branchIcon = '├'; // Has single child
            } else {
                branchIcon = '└'; // Leaf node
            }
        }

        const prefix = version.depth > 0 ? indent + branchIcon + ' ' : '';

        const date = new Date(version.timestamp);
        const dateStr = date.toLocaleString('en-GB', {
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit'
        });

        html += `
            <div class="version-item ${isCurrent ? 'current' : ''} depth-${version.depth}" onclick="goToVersionById('${version.id}')">
                <div class="version-header">
                    <span class="version-tree-prefix">${prefix}</span>
                    <span class="version-date">${dateStr}</span>
                    ${hasMultipleBranches ? '<span class="branch-indicator">⑂</span>' : ''}
                </div>
                <div class="version-description">${version.description}</div>
                <div class="version-meta">${version.wordCount} word(s)${isCurrent ? ' • Current' : ''}</div>
            </div>
        `;
    });

    container.innerHTML = html;
}

function goToVersionById(versionId) {
    if (!versionControl) return;

    const version = versionControl.goToVersion(versionId);
    if (version) {
        words = JSON.parse(JSON.stringify(version.data)); // Deep copy
        localStorage.setItem('wordMemoryData', JSON.stringify(words));
        renderWords();
        renderVersionHistory(); // Update the history display
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
