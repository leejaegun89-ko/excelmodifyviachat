// DOM Elements
const excelFileInput = document.getElementById('excelFile');
const modificationRequest = document.getElementById('modificationRequest');
const sendRequestButton = document.getElementById('sendRequest');
const downloadButton = document.getElementById('downloadButton');
const chatHistory = document.getElementById('chatHistory');

// Global variables
let workbook = null;
let currentWorksheet = null;

// Handle Excel file upload
excelFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const data = await file.arrayBuffer();
        workbook = XLSX.read(data);
        currentWorksheet = workbook.Sheets[workbook.SheetNames[0]];
        
        addMessage('System', `"${file.name}" has been successfully uploaded.`);
        downloadButton.disabled = true;
    } catch (error) {
        addMessage('System', 'Error occurred while uploading the file.');
        console.error('Error reading Excel file:', error);
    }
});

// Handle modification request submission
sendRequestButton.addEventListener('click', () => {
    const request = modificationRequest.value.trim();
    if (!request) return;

    if (!workbook) {
        addMessage('System', 'Please upload an Excel file first.');
        return;
    }

    addMessage('User', request);
    processModificationRequest(request);
    modificationRequest.value = '';
});

// Handle Enter key press for message sending
modificationRequest.addEventListener('keypress', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        sendRequestButton.click();
    }
});

// Add chat message
function addMessage(sender, text) {
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${sender.toLowerCase()}-message`;
    messageDiv.textContent = `${sender}: ${text}`;
    chatHistory.appendChild(messageDiv);
    chatHistory.scrollTop = chatHistory.scrollHeight;
}

// Process modification request
function processModificationRequest(request) {
    try {
        if (!currentWorksheet) {
            throw new Error('No worksheet loaded');
        }

        console.log('Processing request:', request);
        
        // Parse the current worksheet data
        let wsData = XLSX.utils.sheet_to_json(currentWorksheet, { header: 1 });
        console.log('Original worksheet data:', wsData);

        // Get the header row
        const header = wsData[0];
        const dataRows = wsData.slice(1);

        // Function to get column index from letter
        const getColumnIndex = (colLetter) => {
            return colLetter.toUpperCase().charCodeAt(0) - 'A'.charCodeAt(0);
        };

        // Function to get column letter from index
        const getColumnLetter = (index) => {
            return String.fromCharCode(65 + index);
        };

        // Function to split a range (e.g., "xxx1-3" to ["xxx1", "xxx2", "xxx3"])
        const expandRange = (value) => {
            const match = value.match(/^(.+?)(\d+)-(\d+)$/);
            if (!match) return null;
            
            const [_, prefix, start, end] = match;
            const result = [];
            for (let i = parseInt(start); i <= parseInt(end); i++) {
                result.push(prefix + i);
            }
            return result;
        };

        // Function to split values by common separators
        const splitValues = (value) => {
            if (typeof value !== 'string') return [value];
            // Split by comma, slash, or space
            return value.split(/[,\/\s]+/).filter(v => v.trim());
        };

        // Process the modification request
        if (request.toLowerCase().includes('column')) {
            const colMatch = request.match(/column\s+([A-Z])/i);
            if (colMatch) {
                const targetCol = colMatch[1].toUpperCase();
                const colIndex = getColumnIndex(targetCol);
                
                // Create new rows array
                let newRows = [];
                
                // Process each data row
                dataRows.forEach(row => {
                    if (!row[colIndex]) {
                        newRows.push(row);
                        return;
                    }

                    const value = row[colIndex].toString();
                    let splitValues = [];

                    // Try to expand ranges first (e.g., xxx1-3)
                    const rangeExpanded = expandRange(value);
                    if (rangeExpanded) {
                        splitValues = rangeExpanded;
                    } else {
                        // Otherwise split by common separators
                        splitValues = value.split(/[,\/\s]+/).filter(v => v.trim());
                    }

                    // If we found multiple values, create new rows
                    if (splitValues.length > 1) {
                        splitValues.forEach(splitValue => {
                            const newRow = [...row];
                            newRow[colIndex] = splitValue;
                            newRows.push(newRow);
                        });
                    } else {
                        newRows.push(row);
                    }
                });

                // Update worksheet with new data
                wsData = [header, ...newRows];
                currentWorksheet = XLSX.utils.aoa_to_sheet(wsData);
                workbook.Sheets[workbook.SheetNames[0]] = currentWorksheet;
                
                console.log('Modified worksheet data:', wsData);
                addMessage('System', `Modified column ${targetCol}. Split ${newRows.length - dataRows.length} additional rows.`);
            }
        }

        // Enable download button after modification is complete
        downloadButton.disabled = false;
        addMessage('System', 'Modification complete. Click the download button to get your modified file.');
    } catch (error) {
        addMessage('System', 'Error occurred while processing the modification: ' + error.message);
        console.error('Error processing modification:', error);
    }
}

// Handle modified file download
downloadButton.addEventListener('click', () => {
    if (!workbook) return;

    try {
        // Save the modified workbook to a new file
        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        
        // Create and trigger download link
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'modified_excel_file.xlsx';
        a.click();
        window.URL.revokeObjectURL(url);
        
        addMessage('System', 'File has been successfully downloaded.');
    } catch (error) {
        addMessage('System', 'Error occurred while downloading the file.');
        console.error('Error downloading file:', error);
    }
}); 