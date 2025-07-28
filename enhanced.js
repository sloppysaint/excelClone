// Enhanced Interactive Features for Spreadsheet

class EnhancedSpreadsheet {
    constructor() {
        this.currentCell = null;
        this.selectedCells = [];
        this.clipboard = null;
        this.history = [];
        this.historyIndex = -1;
        this.isSelecting = false;
        this.selectionStart = null;
        
        this.init();
    }

    init() {
        this.setupEventListeners();
        this.generateGrid();
        this.initializeFeatures();
    }

    setupEventListeners() {
        // Menu button events
        document.getElementById('newBtn').addEventListener('click', () => this.newSpreadsheet());
        document.getElementById('saveBtn').addEventListener('click', () => this.saveSpreadsheet());
        document.getElementById('undoBtn').addEventListener('click', () => this.undo());
        document.getElementById('redoBtn').addEventListener('click', () => this.redo());
        document.getElementById('deleteBtn').addEventListener('click', () => this.deleteSelection());

        // Format button events
        document.getElementById('boldBtn').addEventListener('click', () => this.toggleFormat('bold'));
        document.getElementById('underlineBtn').addEventListener('click', () => this.toggleFormat('underline'));
        document.getElementById('italicBtn').addEventListener('click', () => this.toggleFormat('italic'));

        // Alignment events
        document.querySelectorAll('.align-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const alignment = e.currentTarget.dataset.align;
                this.setAlignment(alignment);
            });
        });

        // Color events
        document.getElementById('textColorBtn').addEventListener('click', () => {
            document.getElementById('textColorPicker').click();
        });
        document.getElementById('bgColorBtn').addEventListener('click', () => {
            document.getElementById('bgColorPicker').click();
        });
        
        document.getElementById('textColorPicker').addEventListener('change', (e) => {
            this.setTextColor(e.target.value);
        });
        document.getElementById('bgColorPicker').addEventListener('change', (e) => {
            this.setBackgroundColor(e.target.value);
        });

        // Font events
        document.querySelector('.font_family_input').addEventListener('change', (e) => {
            this.setFontFamily(e.target.value);
        });
        document.querySelector('.font_size_input').addEventListener('change', (e) => {
            this.setFontSize(e.target.value + 'px');
        });

        // Formula bar events
        document.querySelector('.formula_input').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                this.applyFormula(e.target.value);
            }
        });

        // Sheet events
        document.querySelector('.add-sheet_container').addEventListener('click', () => {
            this.addNewSheet();
        });

        // Context menu events
        document.addEventListener('contextmenu', (e) => {
            if (e.target.classList.contains('cell')) {
                e.preventDefault();
                this.showContextMenu(e);
            }
        });

        document.addEventListener('click', () => {
            this.hideContextMenu();
        });

        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            this.handleKeyboardShortcuts(e);
        });

        // Grid interaction events
        document.addEventListener('click', (e) => {
            if (e.target.classList.contains('cell')) {
                this.selectCell(e.target);
            }
        });
    }

    generateGrid() {
        const topRow = document.querySelector('.top_row');
        const leftCol = document.querySelector('.left_col');
        const grid = document.querySelector('.grid');

        // Clear existing content
        topRow.innerHTML = '';
        leftCol.innerHTML = '';
        grid.innerHTML = '';

        // Generate column headers (A, B, C, ...)
        for (let i = 0; i < 26; i++) {
            const header = document.createElement('div');
            header.className = 'cell';
            header.textContent = String.fromCharCode(65 + i);
            header.addEventListener('click', () => this.selectColumn(i));
            topRow.appendChild(header);
        }

        // Generate row numbers and cells
        for (let row = 0; row < 100; row++) {
            // Row header
            const rowHeader = document.createElement('div');
            rowHeader.className = 'cell';
            rowHeader.textContent = row + 1;
            rowHeader.addEventListener('click', () => this.selectRow(row));
            leftCol.appendChild(rowHeader);

            // Row of cells
            const rowDiv = document.createElement('div');
            rowDiv.className = 'row';

            for (let col = 0; col < 26; col++) {
                const cell = document.createElement('div');
                cell.className = 'cell';
                cell.contentEditable = true;
                cell.dataset.row = row;
                cell.dataset.col = col;
                cell.dataset.address = String.fromCharCode(65 + col) + (row + 1);

                // Cell event listeners
                cell.addEventListener('focus', () => this.onCellFocus(cell));
                cell.addEventListener('blur', () => this.onCellBlur(cell));
                cell.addEventListener('input', () => this.onCellInput(cell));
                cell.addEventListener('keydown', (e) => this.onCellKeydown(e, cell));
                cell.addEventListener('mousedown', (e) => this.onCellMouseDown(e, cell));
                cell.addEventListener('mouseup', (e) => this.onCellMouseUp(e, cell));
                cell.addEventListener('mouseover', (e) => this.onCellMouseOver(e, cell));

                rowDiv.appendChild(cell);
            }
            grid.appendChild(rowDiv);
        }
    }

    initializeFeatures() {
        // Initialize tooltips
        this.initTooltips();
        
        // Set default values
        document.querySelector('.address_input').value = 'A1';
        
        // Auto-save functionality
        setInterval(() => {
            this.autoSave();
        }, 30000); // Auto-save every 30 seconds
    }

    // Cell interaction methods
    onCellFocus(cell) {
        this.currentCell = cell;
        document.querySelector('.address_input').value = cell.dataset.address;
        document.querySelector('.formula_input').value = cell.textContent;
        
        // Update formatting buttons
        this.updateFormatButtons(cell);
    }

    onCellBlur(cell) {
        this.saveState();
    }

    onCellInput(cell) {
        // Real-time formula calculation
        if (cell.textContent.startsWith('=')) {
            this.calculateFormula(cell);
        }
    }

    onCellKeydown(e, cell) {
        switch(e.key) {
            case 'Enter':
                e.preventDefault();
                this.moveToNextCell(cell, 'down');
                break;
            case 'Tab':
                e.preventDefault();
                this.moveToNextCell(cell, e.shiftKey ? 'left' : 'right');
                break;
            case 'ArrowUp':
                e.preventDefault();
                this.moveToNextCell(cell, 'up');
                break;
            case 'ArrowDown':
                e.preventDefault();
                this.moveToNextCell(cell, 'down');
                break;
            case 'ArrowLeft':
                if (e.ctrlKey) {
                    e.preventDefault();
                    this.moveToNextCell(cell, 'left');
                }
                break;
            case 'ArrowRight':
                if (e.ctrlKey) {
                    e.preventDefault();
                    this.moveToNextCell(cell, 'right');
                }
                break;
            case 'Delete':
                cell.textContent = '';
                this.saveState();
                break;
        }
    }

    onCellMouseDown(e, cell) {
        this.isSelecting = true;
        this.selectionStart = cell;
        this.selectCell(cell);
    }

    onCellMouseUp(e, cell) {
        this.isSelecting = false;
    }

    onCellMouseOver(e, cell) {
        if (this.isSelecting && this.selectionStart) {
            this.selectRange(this.selectionStart, cell);
        }
    }

    // Navigation methods
    moveToNextCell(currentCell, direction) {
        const row = parseInt(currentCell.dataset.row);
        const col = parseInt(currentCell.dataset.col);
        let newRow = row;
        let newCol = col;

        switch(direction) {
            case 'up':
                newRow = Math.max(0, row - 1);
                break;
            case 'down':
                newRow = Math.min(99, row + 1);
                break;
            case 'left':
                newCol = Math.max(0, col - 1);
                break;
            case 'right':
                newCol = Math.min(25, col + 1);
                break;
        }

        const nextCell = document.querySelector(`[data-row="${newRow}"][data-col="${newCol}"]`);
        if (nextCell) {
            nextCell.focus();
        }
    }

    // Selection methods
    selectCell(cell) {
        // Clear previous selection
        document.querySelectorAll('.cell.selected').forEach(c => {
            c.classList.remove('selected');
        });

        cell.classList.add('selected');
        this.selectedCells = [cell];
        this.currentCell = cell;
    }

    selectRange(startCell, endCell) {
        // Clear previous selection
        document.querySelectorAll('.cell.selected').forEach(c => {
            c.classList.remove('selected');
        });

        const startRow = parseInt(startCell.dataset.row);
        const startCol = parseInt(startCell.dataset.col);
        const endRow = parseInt(endCell.dataset.row);
        const endCol = parseInt(endCell.dataset.col);

        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);

        this.selectedCells = [];

        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                const cell = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
                if (cell) {
                    cell.classList.add('selected');
                    this.selectedCells.push(cell);
                }
            }
        }
    }

    selectRow(rowIndex) {
        // Clear previous selection
        document.querySelectorAll('.cell.selected').forEach(c => {
            c.classList.remove('selected');
        });

        this.selectedCells = [];
        for (let col = 0; col < 26; col++) {
            const cell = document.querySelector(`[data-row="${rowIndex}"][data-col="${col}"]`);
            if (cell) {
                cell.classList.add('selected');
                this.selectedCells.push(cell);
            }
        }
    }

    selectColumn(colIndex) {
        // Clear previous selection
        document.querySelectorAll('.cell.selected').forEach(c => {
            c.classList.remove('selected');
        });

        this.selectedCells = [];
        for (let row = 0; row < 100; row++) {
            const cell = document.querySelector(`[data-row="${row}"][data-col="${colIndex}"]`);
            if (cell) {
                cell.classList.add('selected');
                this.selectedCells.push(cell);
            }
        }
    }

    // Formatting methods
    toggleFormat(format) {
        if (!this.selectedCells.length) return;

        this.selectedCells.forEach(cell => {
            const currentStyle = window.getComputedStyle(cell);
            switch(format) {
                case 'bold':
                    cell.style.fontWeight = currentStyle.fontWeight === 'bold' ? 'normal' : 'bold';
                    break;
                case 'italic':
                    cell.style.fontStyle = currentStyle.fontStyle === 'italic' ? 'normal' : 'italic';
                    break;
                case 'underline':
                    cell.style.textDecoration = currentStyle.textDecoration.includes('underline') ? 'none' : 'underline';
                    break;
            }
        });

        this.updateFormatButtons();
        this.saveState();
    }

    setAlignment(alignment) {
        if (!this.selectedCells.length) return;

        // Update alignment buttons
        document.querySelectorAll('.align-btn').forEach(btn => {
            btn.classList.remove('selected');
        });
        document.querySelector(`[data-align="${alignment}"]`).classList.add('selected');

        // Apply alignment to selected cells
        this.selectedCells.forEach(cell => {
            cell.style.textAlign = alignment;
        });

        this.saveState();
    }

    setTextColor(color) {
        if (!this.selectedCells.length) return;

        this.selectedCells.forEach(cell => {
            cell.style.color = color;
        });

        document.getElementById('textColorIndicator').style.background = color;
        this.saveState();
    }

    setBackgroundColor(color) {
        if (!this.selectedCells.length) return;

        this.selectedCells.forEach(cell => {
            cell.style.backgroundColor = color;
        });

        document.getElementById('bgColorIndicator').style.background = color;
        this.saveState();
    }

    setFontFamily(fontFamily) {
        if (!this.selectedCells.length) return;

        this.selectedCells.forEach(cell => {
            cell.style.fontFamily = fontFamily;
        });

        this.saveState();
    }

    setFontSize(fontSize) {
        if (!this.selectedCells.length) return;

        this.selectedCells.forEach(cell => {
            cell.style.fontSize = fontSize;
        });

        this.saveState();
    }

    updateFormatButtons(cell = null) {
        const targetCell = cell || this.currentCell;
        if (!targetCell) return;

        const style = window.getComputedStyle(targetCell);

        // Update format buttons
        document.getElementById('boldBtn').classList.toggle('selected', style.fontWeight === 'bold' || style.fontWeight >= 600);
        document.getElementById('italicBtn').classList.toggle('selected', style.fontStyle === 'italic');
        document.getElementById('underlineBtn').classList.toggle('selected', style.textDecoration.includes('underline'));

        // Update color indicators
        document.getElementById('textColorIndicator').style.background = style.color;
        document.getElementById('bgColorIndicator').style.background = style.backgroundColor;
    }

    // Formula methods
    applyFormula(formula) {
        if (!this.currentCell) return;

        this.currentCell.textContent = formula;
        if (formula.startsWith('=')) {
            this.calculateFormula(this.currentCell);
        }
        this.saveState();
    }

    calculateFormula(cell) {
        const formula = cell.textContent;
        if (!formula.startsWith('=')) return;

        try {
            // Simple formula calculation (extend as needed)
            const expression = formula.substring(1).replace(/[A-Z]+\d+/g, (match) => {
                const refCell = document.querySelector(`[data-address="${match}"]`);
                return refCell ? (parseFloat(refCell.textContent) || 0) : 0;
            });

            const result = eval(expression);
            cell.setAttribute('data-formula', formula);
            cell.setAttribute('title', `Formula: ${formula}`);
            
            // Display result but keep formula in data attribute
            setTimeout(() => {
                cell.textContent = result;
            }, 100);
        } catch (error) {
            cell.textContent = '#ERROR';
        }
    }

    // Context menu methods
    showContextMenu(e) {
        const contextMenu = document.getElementById('contextMenu');
        contextMenu.style.display = 'block';
        contextMenu.style.left = e.pageX + 'px';
        contextMenu.style.top = e.pageY + 'px';

        // Add context menu item listeners
        document.querySelectorAll('.context-item').forEach(item => {
            item.onclick = (event) => {
                this.handleContextAction(event.currentTarget.dataset.action);
                this.hideContextMenu();
            };
        });
    }

    hideContextMenu() {
        document.getElementById('contextMenu').style.display = 'none';
    }

    handleContextAction(action) {
        switch(action) {
            case 'cut':
                this.cutSelection();
                break;
            case 'copy':
                this.copySelection();
                break;
            case 'paste':
                this.pasteSelection();
                break;
            case 'insert-row':
                this.insertRow();
                break;
            case 'insert-column':
                this.insertColumn();
                break;
            case 'delete-row':
                this.deleteRow();
                break;
            case 'delete-column':
                this.deleteColumn();
                break;
        }
    }

    // Clipboard methods
    cutSelection() {
        this.copySelection();
        this.deleteSelection();
    }

    copySelection() {
        if (!this.selectedCells.length) return;

        this.clipboard = this.selectedCells.map(cell => ({
            content: cell.textContent,
            style: cell.getAttribute('style') || ''
        }));
    }

    pasteSelection() {
        if (!this.clipboard || !this.currentCell) return;

        const startRow = parseInt(this.currentCell.dataset.row);
        const startCol = parseInt(this.currentCell.dataset.col);

        this.clipboard.forEach((data, index) => {
            const targetRow = startRow + Math.floor(index / Math.sqrt(this.clipboard.length));
            const targetCol = startCol + (index % Math.sqrt(this.clipboard.length));
            
            const targetCell = document.querySelector(`[data-row="${targetRow}"][data-col="${targetCol}"]`);
            if (targetCell) {
                targetCell.textContent = data.content;
                targetCell.setAttribute('style', data.style);
            }
        });

        this.saveState();
    }

    deleteSelection() {
        this.selectedCells.forEach(cell => {
            cell.textContent = '';
            cell.removeAttribute('style');
        });
        this.saveState();
    }

    // History methods
    saveState() {
        const state = this.getCurrentState();
        this.history = this.history.slice(0, this.historyIndex + 1);
        this.history.push(state);
        this.historyIndex++;

        // Limit history size
        if (this.history.length > 50) {
            this.history.shift();
            this.historyIndex--;
        }
    }

    getCurrentState() {
        const cells = {};
        document.querySelectorAll('.grid .cell').forEach(cell => {
            const address = cell.dataset.address;
            cells[address] = {
                content: cell.textContent,
                style: cell.getAttribute('style') || ''
            };
        });
        return { cells };
    }

    undo() {
        if (this.historyIndex > 0) {
            this.historyIndex--;
            this.restoreState(this.history[this.historyIndex]);
        }
    }

    redo() {
        if (this.historyIndex < this.history.length - 1) {
            this.historyIndex++;
            this.restoreState(this.history[this.historyIndex]);
        }
    }

    restoreState(state) {
        Object.keys(state.cells).forEach(address => {
            const cell = document.querySelector(`[data-address="${address}"]`);
            if (cell) {
                const cellData = state.cells[address];
                cell.textContent = cellData.content;
                if (cellData.style) {
                    cell.setAttribute('style', cellData.style);
                } else {
                    cell.removeAttribute('style');
                }
            }
        });
    }

    // Sheet methods
    addNewSheet() {
        const sheetsList = document.querySelector('.sheets-list');
        const sheetCount = sheetsList.children.length + 1;
        
        const newSheet = document.createElement('div');
        newSheet.className = 'sheet fade-in';
        newSheet.dataset.sheetIdx = sheetCount - 1;
        newSheet.innerHTML = `
            <span class="sheet-name">Sheet ${sheetCount}</span>
            <button class="sheet-close" title="Close sheet">
                <i class="fas fa-times"></i>
            </button>
        `;

        // Add event listeners
        newSheet.addEventListener('click', () => this.switchSheet(newSheet));
        newSheet.querySelector('.sheet-close').addEventListener('click', (e) => {
            e.stopPropagation();
            this.closeSheet(newSheet);
        });

        sheetsList.appendChild(newSheet);
        this.switchSheet(newSheet);
    }

    switchSheet(sheet) {
        document.querySelectorAll('.sheet').forEach(s => {
            s.classList.remove('active-sheet');
        });
        sheet.classList.add('active-sheet');
    }

    closeSheet(sheet) {
        if (document.querySelectorAll('.sheet').length > 1) {
            sheet.remove();
        }
    }

    // File operations
    newSpreadsheet() {
        if (confirm('Create a new spreadsheet? Unsaved changes will be lost.')) {
            document.querySelectorAll('.grid .cell').forEach(cell => {
                cell.textContent = '';
                cell.removeAttribute('style');
            });
            this.history = [];
            this.historyIndex = -1;
            this.saveState();
        }
    }

    saveSpreadsheet() {
        const data = this.getCurrentState();
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = 'spreadsheet.json';
        a.click();
        
        URL.revokeObjectURL(url);
        this.showNotification('Spreadsheet saved successfully!');
    }

    autoSave() {
        const data = this.getCurrentState();
        try {
            localStorage.setItem('spreadsheet_autosave', JSON.stringify(data));
        } catch (error) {
            console.warn('Auto-save failed:', error);
        }
    }

    // Utility methods
    handleKeyboardShortcuts(e) {
        if (e.ctrlKey || e.metaKey) {
            switch(e.key) {
                case 'z':
                    e.preventDefault();
                    e.shiftKey ? this.redo() : this.undo();
                    break;
                case 'c':
                    e.preventDefault();
                    this.copySelection();
                    break;
                case 'x':
                    e.preventDefault();
                    this.cutSelection();
                    break;
                case 'v':
                    e.preventDefault();
                    this.pasteSelection();
                    break;
                case 's':
                    e.preventDefault();
                    this.saveSpreadsheet();
                    break;
                case 'b':
                    e.preventDefault();
                    this.toggleFormat('bold');
                    break;
                case 'i':
                    e.preventDefault();
                    this.toggleFormat('italic');
                    break;
                case 'u':
                    e.preventDefault();
                    this.toggleFormat('underline');
                    break;
            }
        }
    }

    initTooltips() {
        // Add hover effects and tooltips
        document.querySelectorAll('[title]').forEach(element => {
            element.addEventListener('mouseenter', (e) => {
                // Custom tooltip implementation can be added here
            });
        });
    }

    showNotification(message, type = 'success') {
        const notification = document.createElement('div');
        notification.className = `notification ${type} fade-in`;
        notification.textContent = message;
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: ${type === 'success' ? '#4caf50' : '#f44336'};
            color: white;
            padding: 12px 20px;
            border-radius: 6px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 10000;
        `;

        document.body.appendChild(notification);

        setTimeout(() => {
            notification.remove();
        }, 3000);
    }

    // Row/Column operations
    insertRow() {
        // Implementation for inserting rows
        this.showNotification('Insert row functionality would be implemented here');
    }

    insertColumn() {
        // Implementation for inserting columns
        this.showNotification('Insert column functionality would be implemented here');
    }

    deleteRow() {
        // Implementation for deleting rows
        this.showNotification('Delete row functionality would be implemented here');
    }

    deleteColumn() {
        // Implementation for deleting columns
        this.showNotification('Delete column functionality would be implemented here');
    }
}

// Initialize the enhanced spreadsheet when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.spreadsheet = new EnhancedSpreadsheet();
});

// Export for use in other scripts
if (typeof module !== 'undefined' && module.exports) {
    module.exports = EnhancedSpreadsheet;
}