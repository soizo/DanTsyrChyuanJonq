// Data storage
let words = [];
let isFullMode = false;
let editingIndex = -1;
let selectedWords = new Set();
let isSelectMode = false;

const SPEAKER_ICON_SVG = '<svg class="icon-speaker" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path fill="currentColor" d="M3 10v4c0 1.1.9 2 2 2h2.35l3.38 2.7c.93.74 2.27.08 2.27-1.1V6.4c0-1.18-1.34-1.84-2.27-1.1L7.35 8H5c-1.1 0-2 .9-2 2Zm14.5 2c0-1.77-1.02-3.29-2.5-4.03v8.06c1.48-.74 2.5-2.26 2.5-4.03ZM15 3.23v2.06c2.89.86 5 3.54 5 6.71s-2.11 5.85-5 6.71v2.06c4.01-.91 7-4.49 7-8.77s-2.99-7.86-7-8.77Z"/></svg>';

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

function togglePosDropdown() {
    const dropdown = document.getElementById('posDropdown');
    dropdown.classList.toggle('active');
    dropdown.closest('.pos-selector').classList.toggle('open', dropdown.classList.contains('active'));
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

function togglePosOption(option) {
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
    const dropdown = document.getElementById('editPosDropdown');
    dropdown.classList.toggle('active');
    dropdown.closest('.pos-selector').classList.toggle('open', dropdown.classList.contains('active'));
}

function updateEditPosSelection() {
    renderPosSelection('editPosSelected', editSelectedPos);
    syncPosOptionState('editPosDropdown', editSelectedPos);
}

function toggleEditPosOption(option) {
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
    const dropdown = document.getElementById('weightDropdown');
    dropdown.classList.toggle('active');
    dropdown.closest('.weight-selector').classList.toggle('open', dropdown.classList.contains('active'));
}

function toggleEditWeightDropdown() {
    const dropdown = document.getElementById('editWeightDropdown');
    dropdown.classList.toggle('active');
    dropdown.closest('.weight-selector').classList.toggle('open', dropdown.classList.contains('active'));
}

function setWeightOption(option) {
    const value = parseInt(option.dataset.value, 10);
    if (Number.isNaN(value)) return;
    selectedWeight = value;
    updateWeightSelection();
    const dropdown = document.getElementById('weightDropdown');
    dropdown.classList.remove('active');
    dropdown.closest('.weight-selector').classList.remove('open');
}

function setEditWeightOption(option) {
    const value = parseInt(option.dataset.value, 10);
    if (Number.isNaN(value)) return;
    editSelectedWeight = value;
    updateEditWeightSelection();
    const dropdown = document.getElementById('editWeightDropdown');
    dropdown.classList.remove('active');
    dropdown.closest('.weight-selector').classList.remove('open');
}

// Close dropdowns when clicking outside
document.addEventListener('click', function(e) {
    if (!e.target.closest('.pos-selector')) {
        document.querySelectorAll('.pos-dropdown').forEach(d => {
            d.classList.remove('active');
            d.closest('.pos-selector').classList.remove('open');
        });
    }
    if (!e.target.closest('.weight-selector')) {
        document.querySelectorAll('.weight-dropdown').forEach(d => {
            d.classList.remove('active');
            d.closest('.weight-selector').classList.remove('open');
        });
    }
});

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
    saveData();
    renderWords();
    showStatus(`✓ Added "${word}"`, 'success');

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

    // If current weight is -3 (invalid), delta is the new weight directly
    if (currentWeight === -3) {
        words[index].weight = delta;
        saveData();
        renderWords();
        showStatus(`✓ Fixed weight to ${delta}`, 'success');
        return;
    }

    const newWeight = currentWeight + delta;

    if (newWeight < -2 || newWeight > 10) {
        return;
    }

    words[index].weight = newWeight;
    saveData();
    renderWords();
}

// Delete word
function deleteWord(index) {
    if (confirm(`Delete "${words[index].word}"?`)) {
        words.splice(index, 1);
        saveData();
        renderWords();
        showStatus('✓ Word deleted', 'success');
    }
}

// Toggle word selection
function toggleWordSelection(index) {
    if (selectedWords.has(index)) {
        selectedWords.delete(index);
    } else {
        selectedWords.add(index);
    }
    updateBatchDeleteButton();
    updateWordSelectionUI();
}

// Toggle select mode
function toggleSelectMode() {
    if (isSelectMode && selectedWords.size > 0) {
        // If in select mode and have selections, perform batch delete
        batchDeleteConfirm();
    } else {
        // Toggle select mode
        isSelectMode = !isSelectMode;
        if (!isSelectMode) {
            selectedWords.clear();
        }
        renderWords();
    }
}

// Update batch delete button
function updateBatchDeleteButton() {
    const btn = document.getElementById('batchDeleteBtn');
    if (btn) {
        if (isSelectMode) {
            if (selectedWords.size > 0) {
                btn.textContent = `Delete ${selectedWords.size} Selected`;
                btn.classList.add('btn-delete-mode');
            } else {
                btn.textContent = 'Cancel';
                btn.classList.remove('btn-delete-mode');
            }
        } else {
            btn.textContent = 'Select Words';
            btn.classList.remove('btn-delete-mode');
        }
        btn.disabled = false;
    }
}

// Update word selection UI
function updateWordSelectionUI() {
    selectedWords.forEach(index => {
        const checkbox = document.getElementById(`select-${index}`);
        if (checkbox) checkbox.checked = true;
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
        isSelectMode = false;
        saveData();
        renderWords();
        showStatus(`✓ Deleted ${count} word(s)`, 'success');
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
                showStatus(`🔊 Playing "${word}"`, 'info');
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
        showStatus(`🔊 Pronouncing "${word}"`, 'info');
    } else {
        showStatus('✗ Pronunciation not supported', 'error');
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

    words[editingIndex] = {
        word: word.toLowerCase(),
        meaning: meaning,
        pos: editSelectedPos.slice(), // Copy array
        weight: weight,
        added: date
    };

    saveData();
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
    updateBatchDeleteButton();
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
function saveData(showMessage = false) {
    localStorage.setItem('wordMemoryData', JSON.stringify(words));
    if (showMessage) {
        showStatus('✓ Auto-saved to localStorage', 'success');
    }
}

// Manual save
function manualSave() {
    localStorage.setItem('wordMemoryData', JSON.stringify(words));
    showStatus(`✓ Manually saved ${words.length} words`, 'success');
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

                    saveData();
                    renderWords();

                    if (invalidCount > 0) {
                        showStatus(`✓ Imported ${validCount} valid, ${invalidCount} invalid words`, 'success');
                    } else {
                        showStatus(`✓ Imported ${imported.length} words`, 'success');
                    }
                }
            } else {
                showStatus('✗ Invalid file format', 'error');
            }
        } catch (error) {
            showStatus('✗ Failed to parse file', 'error');
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
