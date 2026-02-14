// Data storage
let words = [];
let isFullMode = false;
let editingIndex = -1;

// Initialize
window.onload = function() {
    loadData();
    renderWords();

    // Set default date to today
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('dateInput').value = today;
};

// Mode toggle
document.getElementById('modeToggle').addEventListener('click', function() {
    this.classList.toggle('active');
    isFullMode = !isFullMode;
    renderWords();
});

// Add word
function addWord() {
    const word = document.getElementById('wordInput').value.trim();
    const meaning = document.getElementById('meaningInput').value.trim();
    const pos = document.getElementById('posInput').value.trim();
    const weight = parseInt(document.getElementById('weightInput').value);
    const date = document.getElementById('dateInput').value || new Date().toISOString().split('T')[0];

    if (!word || !meaning) {
        alert('Please fill in word and meaning');
        return;
    }

    const newWord = {
        word: word.toLowerCase(),
        meaning: meaning,
        pos: pos,
        weight: weight,
        added: date
    };

    words.push(newWord);
    saveData();
    renderWords();

    // Clear form
    document.getElementById('wordInput').value = '';
    document.getElementById('meaningInput').value = '';
    document.getElementById('posInput').value = '';
    document.getElementById('weightInput').value = '3';
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('dateInput').value = today;
}

// Update weight
function updateWeight(index, delta) {
    const newWeight = words[index].weight + delta;

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
    }
}

// Open edit modal
function openEditModal(index) {
    editingIndex = index;
    const word = words[index];

    document.getElementById('editWordInput').value = word.word;
    document.getElementById('editMeaningInput').value = word.meaning;
    document.getElementById('editPosInput').value = word.pos || '';
    document.getElementById('editWeightInput').value = word.weight;
    document.getElementById('editDateInput').value = word.added;

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
    const pos = document.getElementById('editPosInput').value.trim();
    const weight = parseInt(document.getElementById('editWeightInput').value);
    const date = document.getElementById('editDateInput').value;

    if (!word || !meaning) {
        alert('Please fill in word and meaning');
        return;
    }

    words[editingIndex] = {
        word: word.toLowerCase(),
        meaning: meaning,
        pos: pos,
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
            return w.weight >= -1;
        } else {
            return w.weight >= 0;
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
        const isMastered = weight < 0;

        const groupLabel = weight >= 5 ? `Weight ${weight} — Hardest` :
                          weight === 4 ? `Weight ${weight} — Hard` :
                          weight === 3 ? `Weight ${weight} — Memorize` :
                          weight === 2 ? `Weight ${weight} — Normal` :
                          weight === 1 ? `Weight ${weight} — Recognize` :
                          weight === 0 ? `Weight ${weight} — Basic` :
                          weight === -1 ? `Weight ${weight} — Mastered` :
                          `Weight ${weight} — Easy`;

        if (isMastered) {
            html += `
                <div class="word-group">
                    <div class="collapsible-header" onclick="toggleCollapse(this)">
                        <div class="group-header" style="margin-bottom: 0; border: none;">${groupLabel} (${groupWords.length})</div>
                        <div class="collapse-icon">▼</div>
                    </div>
                    <div class="collapsible-content">
            `;
        } else {
            html += `
                <div class="word-group">
                    <div class="group-header">${groupLabel} (${groupWords.length})</div>
            `;
        }

        groupWords.forEach(w => {
            const originalIndex = words.indexOf(w);
            const masteredClass = isMastered ? 'mastered' : '';

            html += `
                <div class="word-item ${masteredClass}">
                    <div class="word-header">
                        <div>
                            <span class="word-title">${w.word}</span>
                            ${w.pos ? `<span class="word-pos">${w.pos}</span>` : ''}
                        </div>
                        <div class="word-weight">${w.weight}</div>
                    </div>
                    <div class="word-meaning">${w.meaning}</div>
                    <div class="word-meta">Added: ${w.added}</div>
                    <div class="word-actions">
                        ${w.weight >= 0 ? `<button class="btn-remember" onclick="updateWeight(${originalIndex}, -1)">Down</button>` : ''}
                        ${w.weight >= -1 && w.weight < 10 ? `<button class="btn-forget" onclick="updateWeight(${originalIndex}, 1)">Up</button>` : ''}
                        <button class="btn-edit" onclick="openEditModal(${originalIndex})">Edit</button>
                        <button class="btn-delete" onclick="deleteWord(${originalIndex})">Del</button>
                    </div>
                </div>
            `;
        });

        if (isMastered) {
            html += `
                    </div>
                </div>
            `;
        } else {
            html += `</div>`;
        }
    });

    container.innerHTML = html;
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
function saveData() {
    localStorage.setItem('wordMemoryData', JSON.stringify(words));
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
                    words = imported;
                    saveData();
                    renderWords();
                }
            } else {
                alert('Invalid file format');
            }
        } catch (error) {
            alert('Failed to parse file');
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
document.getElementById('posInput').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') addWord();
});
