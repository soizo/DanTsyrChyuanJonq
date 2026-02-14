// Version Control System
class VersionControl {
    constructor(maxVersions = 50) {
        this.versions = [];
        this.currentIndex = -1;
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
            const savedIndex = localStorage.getItem('wordMemoryVersionIndex');
            const savedSettings = localStorage.getItem('wordMemorySettings');

            if (savedVersions) {
                this.versions = JSON.parse(savedVersions);
            }

            if (savedIndex !== null) {
                this.currentIndex = parseInt(savedIndex, 10);
            }

            if (savedSettings) {
                const settings = JSON.parse(savedSettings);
                this.maxVersions = settings.maxVersions || 50;
            }
        } catch (error) {
            console.error('Failed to load version history:', error);
            this.versions = [];
            this.currentIndex = -1;
        }
    }

    // Save version history to localStorage
    saveHistory() {
        try {
            localStorage.setItem('wordMemoryVersions', JSON.stringify(this.versions));
            localStorage.setItem('wordMemoryVersionIndex', this.currentIndex.toString());
        } catch (error) {
            console.error('Failed to save version history:', error);
            // If localStorage is full, try to remove old versions
            if (error.name === 'QuotaExceededError') {
                this.removeOldVersions(Math.floor(this.maxVersions / 2));
                try {
                    localStorage.setItem('wordMemoryVersions', JSON.stringify(this.versions));
                    localStorage.setItem('wordMemoryVersionIndex', this.currentIndex.toString());
                } catch (e) {
                    console.error('Still failed after removing old versions:', e);
                }
            }
        }
    }

    // Create a new version
    createVersion(data, description = 'Manual save') {
        // If we're not at the latest version, remove all versions after current
        if (this.currentIndex < this.versions.length - 1) {
            this.versions = this.versions.slice(0, this.currentIndex + 1);
        }

        // Create new version
        const version = {
            id: this.generateId(),
            timestamp: new Date().toISOString(),
            data: JSON.parse(JSON.stringify(data)), // Deep copy
            description: description,
            wordCount: data.length
        };

        this.versions.push(version);
        this.currentIndex = this.versions.length - 1;

        // Limit number of versions
        if (this.versions.length > this.maxVersions) {
            this.removeOldVersions(1);
        }

        this.saveHistory();
        return version;
    }

    // Remove old versions
    removeOldVersions(count) {
        if (this.versions.length <= 1) return;

        const toRemove = Math.min(count, this.versions.length - 1);
        this.versions.splice(0, toRemove);
        this.currentIndex = Math.max(0, this.currentIndex - toRemove);
    }

    // Undo - go to previous version
    undo() {
        if (!this.canUndo()) {
            return null;
        }

        this.currentIndex--;
        this.saveHistory();
        return this.getCurrentVersion();
    }

    // Redo - go to next version
    redo() {
        if (!this.canRedo()) {
            return null;
        }

        this.currentIndex++;
        this.saveHistory();
        return this.getCurrentVersion();
    }

    // Check if can undo
    canUndo() {
        return this.currentIndex > 0;
    }

    // Check if can redo
    canRedo() {
        return this.currentIndex < this.versions.length - 1;
    }

    // Get current version
    getCurrentVersion() {
        if (this.currentIndex >= 0 && this.currentIndex < this.versions.length) {
            return this.versions[this.currentIndex];
        }
        return null;
    }

    // Go to specific version
    goToVersion(index) {
        if (index >= 0 && index < this.versions.length) {
            this.currentIndex = index;
            this.saveHistory();
            return this.versions[index];
        }
        return null;
    }

    // Get version history
    getHistory() {
        return {
            versions: this.versions,
            currentIndex: this.currentIndex,
            canUndo: this.canUndo(),
            canRedo: this.canRedo()
        };
    }

    // Clear all history
    clearHistory() {
        this.versions = [];
        this.currentIndex = -1;
        this.saveHistory();
    }

    // Update settings
    updateSettings(settings) {
        if (settings.maxVersions) {
            this.maxVersions = settings.maxVersions;
            localStorage.setItem('wordMemorySettings', JSON.stringify({ maxVersions: this.maxVersions }));

            // Trim versions if needed
            if (this.versions.length > this.maxVersions) {
                this.removeOldVersions(this.versions.length - this.maxVersions);
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
    if (versionControl.versions.length === 0 && words.length > 0) {
        versionControl.createVersion(words, 'Initial version');
    } else if (versionControl.versions.length === 0 && words.length === 0) {
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

// Select by weight range
function selectByWeight(minWeight, maxWeight) {
    selectedWords.clear();

    words.forEach((w, index) => {
        if (w.weight >= minWeight && w.weight <= maxWeight) {
            selectedWords.add(index);
        }
    });

    updateBatchToolbar();
    updateWordSelectionUI();
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
function exportData() {
    const dataStr = JSON.stringify(words, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `words-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
}

// Import data
function importData(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const imported = JSON.parse(e.target.result);
            if (Array.isArray(imported)) {
                if (confirm(`Import ${imported.length} words? This will overwrite current data.`)) {
                    // Validate and clean imported data
                    let validCount = 0;
                    let invalidCount = 0;

                    words = imported.map(item => {
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

                    // Clear version history and create new initial version
                    if (versionControl) {
                        versionControl.clearHistory();
                    }

                    saveData(false, `Imported ${imported.length} word(s)`);
                    renderWords();

                    if (invalidCount > 0) {
                        showStatus(`Imported ${validCount} valid, ${invalidCount} invalid words`, 'success');
                    } else {
                        showStatus(`Imported ${imported.length} words`, 'success');
                    }
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

    const history = versionControl.getHistory();
    const container = document.getElementById('historyList');

    // Update version info
    document.getElementById('currentVersionIndex').textContent = history.currentIndex + 1;
    document.getElementById('totalVersions').textContent = history.versions.length;
    document.getElementById('canUndo').textContent = history.canUndo ? 'Yes' : 'No';
    document.getElementById('canRedo').textContent = history.canRedo ? 'Yes' : 'No';

    if (history.versions.length === 0) {
        container.innerHTML = '<div class="empty-state"><div class="empty-state-icon">—</div><div>No Version History</div></div>';
        return;
    }

    let html = '';
    history.versions.forEach((version, index) => {
        const isCurrent = index === history.currentIndex;
        const date = new Date(version.timestamp);
        const dateStr = date.toLocaleString('en-GB', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        });

        html += `
            <div class="version-item ${isCurrent ? 'current' : ''}" onclick="goToVersionByIndex(${index})">
                <div class="version-header">
                    <span class="version-number">#${index + 1}</span>
                    <span class="version-date">${dateStr}</span>
                </div>
                <div class="version-description">${version.description}</div>
                <div class="version-meta">${version.wordCount} word(s)${isCurrent ? ' • Current' : ''}</div>
            </div>
        `;
    });

    container.innerHTML = html;
}

function goToVersionByIndex(index) {
    if (!versionControl) return;

    const version = versionControl.goToVersion(index);
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
