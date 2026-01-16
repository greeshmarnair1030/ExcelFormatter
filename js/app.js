// Excel Formatter Application
// Handles file upload, sorting, and download functionality

class ExcelFormatter {
    constructor() {
        this.workbook = null;
        this.sortedWorkbook = null;
        this.fileName = '';
        this.initializeElements();
        this.attachEventListeners();
    }

    initializeElements() {
        // Get DOM elements
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.browseBtn = document.getElementById('browseBtn');
        this.fileDetails = document.getElementById('fileDetails');
        this.fileNameDisplay = document.getElementById('fileName');
        this.rowCountDisplay = document.getElementById('rowCount');
        this.statusDisplay = document.getElementById('status');
        this.actionSection = document.getElementById('actionSection');
        this.sortBtn = document.getElementById('sortBtn');
        this.resetBtn = document.getElementById('resetBtn');
        this.downloadSection = document.getElementById('downloadSection');
        this.downloadBtn = document.getElementById('downloadBtn');
        this.message = document.getElementById('message');
        this.loading = document.getElementById('loading');
    }

    attachEventListeners() {
        // Browse button click
        this.browseBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.fileInput.click();
        });

        // Upload area click (but not on the button)
        this.uploadArea.addEventListener('click', (e) => {
            if (e.target !== this.browseBtn && !this.browseBtn.contains(e.target)) {
                this.fileInput.click();
            }
        });

        // File input change
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

        // Drag and drop events
        this.uploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
        this.uploadArea.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        this.uploadArea.addEventListener('drop', (e) => this.handleDrop(e));

        // Sort button
        this.sortBtn.addEventListener('click', () => this.sortExcelFile());

        // Reset button
        this.resetBtn.addEventListener('click', () => this.reset());

        // Download button
        this.downloadBtn.addEventListener('click', () => this.downloadFile());
    }

    handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadArea.classList.add('drag-over');
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadArea.classList.remove('drag-over');
    }

    handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadArea.classList.remove('drag-over');

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    handleFileSelect(e) {
        const files = e.target.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    processFile(file) {
        // Validate file type
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];
        const validExtensions = ['.xlsx', '.xls'];
        const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

        if (!validTypes.includes(file.type) && !validExtensions.includes(fileExtension)) {
            this.showMessage('Please upload a valid Excel file (.xlsx or .xls)', 'error');
            return;
        }

        // Validate file size (max 10MB)
        if (file.size > 10 * 1024 * 1024) {
            this.showMessage('File size exceeds 10MB. Please upload a smaller file.', 'error');
            return;
        }

        this.fileName = file.name;
        this.readExcelFile(file);
    }

    readExcelFile(file) {
        this.showLoading(true);
        this.hideMessage();

        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                this.workbook = XLSX.read(data, { type: 'array' });

                // Debug: Log all sheet names
                console.log('All sheets in workbook:', this.workbook.SheetNames);

                // Find "Bugs Reported" sheet (case-insensitive, space-tolerant)
                const bugsReportedSheet = this.findSheet(this.workbook.SheetNames, ['bugsreported', 'bugs reported']);
                
                if (!bugsReportedSheet) {
                    this.showMessage(
                        `Excel file must contain a sheet named "Bugs Reported". Found sheets: ${this.workbook.SheetNames.join(', ')}`,
                        'error'
                    );
                    this.showLoading(false);
                    return;
                }

                console.log('Using sheet:', bugsReportedSheet);
                const worksheet = this.workbook.Sheets[bugsReportedSheet];

                // Read with header option to preserve all columns including empty ones
                let jsonData = XLSX.utils.sheet_to_json(worksheet, {
                    defval: '',
                    raw: false
                });
                
                const rowCount = jsonData.length;

                // Validate required columns with flexible matching
                let hasNewExisting = null;
                let hasPriority = null;
                
                if (jsonData.length > 0) {
                    const firstRow = jsonData[0];
                    const allColumnNames = Object.keys(firstRow);
                    
                    console.log('All column names (including __EMPTY):', allColumnNames);
                    console.log('First row data:', firstRow);
                    
                    // Check all columns including __EMPTY ones
                    hasNewExisting = this.findColumn(allColumnNames, ['new/existing', 'existing/new', 'newexisting', 'existingnew']);
                    hasPriority = this.findColumn(allColumnNames, ['priority']);

                    // If New/Existing column not found by name, check if __EMPTY column contains New/Existing values
                    if (!hasNewExisting) {
                        const emptyColumns = allColumnNames.filter(col => col.includes('__EMPTY'));
                        for (const emptyCol of emptyColumns) {
                            const sampleValue = (firstRow[emptyCol] || '').toString().toLowerCase().trim();
                            if (sampleValue === 'new' || sampleValue === 'existing') {
                                hasNewExisting = emptyCol;
                                console.log(`Found New/Existing data in column: ${emptyCol}`);
                                break;
                            }
                        }
                    }

                    if (!hasNewExisting || !hasPriority) {
                        const missingCols = [];
                        if (!hasNewExisting) missingCols.push('"New/Existing"');
                        if (!hasPriority) missingCols.push('"Priority"');
                        
                        // Show all columns for debugging
                        const debugInfo = allColumnNames.map((col, idx) => {
                            const sampleValue = firstRow[col];
                            return `Column ${idx + 1}: "${col}" (sample: "${sampleValue}")`;
                        }).join(', ');
                        
                        this.showMessage(
                            `Missing columns: ${missingCols.join(', ')}. Debug info: ${debugInfo}`,
                            'error'
                        );
                        this.showLoading(false);
                        return;
                    }
                }

                // Store the sheet name and column mapping for later use
                this.bugsReportedSheetName = bugsReportedSheet;
                this.newExistingColumn = hasNewExisting;

                // Display file details
                this.displayFileDetails(rowCount);
                
                // Auto-sort the file immediately
                this.sortExcelFile();

            } catch (error) {
                console.error('Error reading Excel file:', error);
                console.error('Error stack:', error.stack);
                this.showMessage(`Error reading Excel file: ${error.message}. Check console for details.`, 'error');
                this.showLoading(false);
            }
        };

        reader.onerror = () => {
            this.showMessage('Error reading file. Please try again.', 'error');
            this.showLoading(false);
        };

        reader.readAsArrayBuffer(file);
    }

    // Helper method to find sheet with flexible matching (case-insensitive, space-tolerant)
    findSheet(sheetNames, searchTerms) {
        const normalizedSheets = sheetNames.map(sheet =>
            sheet.toLowerCase().replace(/\s+/g, '')
        );
        
        for (const term of searchTerms) {
            const normalizedTerm = term.toLowerCase().replace(/\s+/g, '');
            const index = normalizedSheets.indexOf(normalizedTerm);
            if (index !== -1) {
                return sheetNames[index]; // Return original sheet name
            }
        }
        return null;
    }

    // Helper method to find column with flexible matching (case-insensitive, space-tolerant)
    findColumn(columnNames, searchTerms) {
        const normalizedColumns = columnNames.map(col =>
            col.toLowerCase().replace(/\s+/g, '')
        );
        
        for (const term of searchTerms) {
            const normalizedTerm = term.toLowerCase().replace(/\s+/g, '');
            const index = normalizedColumns.indexOf(normalizedTerm);
            if (index !== -1) {
                return columnNames[index]; // Return original column name
            }
        }
        return null;
    }

    displayFileDetails(rowCount) {
        this.fileNameDisplay.textContent = this.fileName;
        this.rowCountDisplay.textContent = rowCount;
        this.statusDisplay.textContent = 'Ready to sort';
        this.statusDisplay.style.background = '#4caf50';

        // Hide upload area, show file details and action buttons
        this.uploadArea.style.display = 'none';
        this.fileDetails.style.display = 'block';
        this.actionSection.style.display = 'flex';
        this.downloadSection.style.display = 'none';
    }

    sortExcelFile() {
        if (!this.workbook) {
            this.showMessage('Please upload a file first.', 'error');
            return;
        }

        this.showLoading(true);
        this.hideMessage();

        // Use setTimeout to allow UI to update
        setTimeout(() => {
            try {
                // Get the "Bugs Reported" sheet
                const worksheet = this.workbook.Sheets[this.bugsReportedSheetName];

                // Convert to JSON with same options as validation
                let data = XLSX.utils.sheet_to_json(worksheet, {
                    defval: '',
                    raw: false
                });

                console.log('Data before sorting:', data.length, 'rows');
                console.log('Columns in data:', Object.keys(data[0]));

                // Sort the data
                data = this.applySorting(data);

                console.log('Data after sorting:', data.length, 'rows');

                // Rename __EMPTY column to "New/Existing" if needed
                if (this.newExistingColumn && this.newExistingColumn.includes('__EMPTY')) {
                    console.log('Renaming column:', this.newExistingColumn, '-> New/Existing');
                    data = data.map(row => {
                        const newRow = {};
                        for (const [key, value] of Object.entries(row)) {
                            if (key === this.newExistingColumn) {
                                newRow['New/Existing'] = value;
                            } else {
                                newRow[key] = value;
                            }
                        }
                        return newRow;
                    });
                }

                // Create new workbook with all original sheets
                this.sortedWorkbook = XLSX.utils.book_new();
                
                // Copy all sheets from original workbook
                for (const sheetName of this.workbook.SheetNames) {
                    if (sheetName === this.bugsReportedSheetName) {
                        // Use sorted data for "Bugs Reported" sheet
                        const newWorksheet = XLSX.utils.json_to_sheet(data);
                        XLSX.utils.book_append_sheet(this.sortedWorkbook, newWorksheet, sheetName);
                    } else {
                        // Copy other sheets as-is
                        const originalSheet = this.workbook.Sheets[sheetName];
                        XLSX.utils.book_append_sheet(this.sortedWorkbook, originalSheet, sheetName);
                    }
                }

                // Update status
                this.statusDisplay.textContent = 'Sorted successfully';
                this.statusDisplay.style.background = '#2196f3';

                // Hide action section, show download section
                this.actionSection.style.display = 'none';
                this.downloadSection.style.display = 'block';
                this.downloadSection.classList.add('fade-in');

                this.showLoading(false);
                this.showMessage('File sorted successfully! Click download to get your sorted file.', 'success');

            } catch (error) {
                console.error('Error sorting Excel file:', error);
                console.error('Error stack:', error.stack);
                this.showMessage(`Error sorting file: ${error.message}`, 'error');
                this.showLoading(false);
            }
        }, 100);
    }

    applySorting(data) {
        if (data.length === 0) return data;

        // Use stored column name (which might be __EMPTY)
        const newExistingCol = this.newExistingColumn;
        
        // Find other columns
        const columnNames = Object.keys(data[0]);
        const priorityCol = this.findColumn(columnNames, ['priority']);
        const bugsCol = this.findColumn(columnNames, ['bugs']);

        console.log('Sorting with columns:', {
            newExisting: newExistingCol,
            priority: priorityCol,
            bugs: bugsCol
        });

        // Define sort order with case-insensitive matching
        const newExistingOrder = { 'new': 0, 'existing': 1 };
        const priorityOrder = { 'blocker': 0, 'critical': 1, 'major': 2 };

        // Sort with three-level priority
        return data.sort((a, b) => {
            // Primary sort: New/Existing (New first)
            const typeA = (a[newExistingCol] || '').toString().toLowerCase().trim();
            const typeB = (b[newExistingCol] || '').toString().toLowerCase().trim();
            const typeOrderA = newExistingOrder[typeA] !== undefined ? newExistingOrder[typeA] : 999;
            const typeOrderB = newExistingOrder[typeB] !== undefined ? newExistingOrder[typeB] : 999;

            if (typeOrderA !== typeOrderB) {
                return typeOrderA - typeOrderB;
            }

            // Secondary sort: Priority (Blocker > Critical > Major > Others)
            const priorityA = (a[priorityCol] || '').toString().toLowerCase().trim();
            const priorityB = (b[priorityCol] || '').toString().toLowerCase().trim();
            const priorityOrderA = priorityOrder[priorityA] !== undefined ? priorityOrder[priorityA] : 999;
            const priorityOrderB = priorityOrder[priorityB] !== undefined ? priorityOrder[priorityB] : 999;

            if (priorityOrderA !== priorityOrderB) {
                return priorityOrderA - priorityOrderB;
            }

            // Tertiary sort: Bugs column by date (latest first - descending)
            if (bugsCol) {
                const dateA = this.extractDate(a[bugsCol]);
                const dateB = this.extractDate(b[bugsCol]);
                
                // Sort descending (latest date first)
                if (dateA && dateB) {
                    return dateB - dateA;
                } else if (dateA) {
                    return -1; // Items with dates come before items without
                } else if (dateB) {
                    return 1;
                }
            }

            return 0; // Keep original order if all criteria are equal
        });
    }

    // Extract date from text in format DD/MM (e.g., "09/01" or "16/01")
    extractDate(text) {
        if (!text) return null;
        
        const textStr = text.toString();
        // Match date pattern DD/MM
        const datePattern = /(\d{1,2})\/(\d{1,2})/;
        const match = textStr.match(datePattern);
        
        if (match) {
            const day = parseInt(match[1], 10);
            const month = parseInt(match[2], 10);
            
            // For dates spanning year boundary (Dec to Jan), we need to handle year correctly
            // Assume dates are recent - if month is 12 (Dec) and current month is 1-6, it's previous year
            const now = new Date();
            const currentYear = now.getFullYear();
            const currentMonth = now.getMonth() + 1; // getMonth() is 0-based
            
            let year = currentYear;
            // If we're in Jan-Jun and the date is in Dec, it's from previous year
            if (currentMonth <= 6 && month === 12) {
                year = currentYear - 1;
            }
            // If we're in Jul-Dec and the date is in Jan-Jun, it might be next year
            // But for bug tracking, assume it's current year
            
            const date = new Date(year, month - 1, day);
            
            console.log(`Extracted date: ${day}/${month} -> ${date.toISOString()} (timestamp: ${date.getTime()})`);
            
            return date.getTime();
        }
        
        return null;
    }

    downloadFile() {
        if (!this.sortedWorkbook) {
            this.showMessage('No sorted file available. Please sort the file first.', 'error');
            return;
        }

        try {
            // Generate file name
            const originalName = this.fileName.replace(/\.[^/.]+$/, '');
            const newFileName = `${originalName}_sorted.xlsx`;

            // Write workbook and trigger download
            XLSX.writeFile(this.sortedWorkbook, newFileName);

            this.showMessage('File downloaded successfully!', 'success');

        } catch (error) {
            console.error('Error downloading file:', error);
            this.showMessage('Error downloading file. Please try again.', 'error');
        }
    }

    reset() {
        // Reset all state
        this.workbook = null;
        this.sortedWorkbook = null;
        this.fileName = '';
        this.fileInput.value = '';

        // Reset UI
        this.uploadArea.style.display = 'block';
        this.fileDetails.style.display = 'none';
        this.actionSection.style.display = 'none';
        this.downloadSection.style.display = 'none';
        this.hideMessage();
    }

    showMessage(text, type = 'info') {
        this.message.textContent = text;
        this.message.className = `message ${type}`;
        this.message.style.display = 'block';

        // Auto-hide success messages after 5 seconds
        if (type === 'success') {
            setTimeout(() => this.hideMessage(), 5000);
        }
    }

    hideMessage() {
        this.message.style.display = 'none';
    }

    showLoading(show) {
        this.loading.style.display = show ? 'block' : 'none';
    }
}

// Initialize the application when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    new ExcelFormatter();
});

// Made with Bob
