document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const directoryPathInput = document.getElementById('directoryPath');
    const scanButton = document.getElementById('scanButton');
    let originalScanButtonText = scanButton.textContent; // Store initial text
    let currentSpinner = null; // To keep track of the added spinner
    const deleteButton = document.getElementById('deleteButton');
    const fileList = document.getElementById('fileList');
    const fileCount = document.getElementById('fileCount');
    const selectAllCheckbox = document.getElementById('selectAll');
    const statusMessage = document.getElementById('statusMessage');
    const darkModeToggle = document.getElementById('darkModeToggle');
    const wideModeToggle = document.getElementById('wideMode');
    
    // State
    let files = [];
    
    // Event Listeners
    scanButton.addEventListener('click', scanForFiles);
    deleteButton.addEventListener('click', deleteSelectedFiles);
    selectAllCheckbox.addEventListener('change', toggleSelectAll);
    
    // Only add dark mode event listener if the element exists
    if (darkModeToggle) {
        darkModeToggle.addEventListener('change', toggleDarkMode);
        
        // Initialize dark mode based on localStorage
        if (localStorage.getItem('darkMode') === 'enabled') {
            document.body.classList.add('dark-mode');
            darkModeToggle.checked = true;
        }
    }
    
    // Initialize wide mode based on localStorage
    if (wideModeToggle) {
        wideModeToggle.addEventListener('change', toggleWideMode);
        
        if (localStorage.getItem('wideMode') === 'enabled') {
            document.body.classList.add('wide-mode');
            wideModeToggle.checked = true;
        }
    }
    
    // Functions
    async function scanForFiles() {
        const directory = directoryPathInput.value.trim();
        
        if (!directory) {
            showStatus('Please enter a directory path', 'error');
            return;
        }
        
        setLoading(true);
        clearFileList();
        
        try {
            const response = await fetch('/api/scan', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ directory })
            });
            
            const data = await response.json();
            
            if (response.ok) {
                // Robust file processing matching CLI tool approach
                let processedFiles = [];
                
                if (data.files && Array.isArray(data.files)) {
                    processedFiles = data.files.filter(file => {
                        // Ensure file has required properties (matching CLI tool validation)
                        return file && 
                               typeof file === 'object' && 
                               file.FullName && 
                               file.Name;
                    }).map(file => ({
                        FullName: file.FullName,
                        Name: file.Name,
                        Directory: file.Directory || '',
                        LastModified: file.LastModified || '',
                        Size: file.Size || 0
                    }));
                } else {
                    // Fallback handling for unexpected data format
                    console.warn('Unexpected data format from server:', data);
                    processedFiles = [];
                }
                
                files = processedFiles;
                displayFiles(files);
                fileCount.textContent = files.length;
                
                if (files.length > 0) {
                    deleteButton.disabled = false;
                    showStatus('Found ' + files.length + ' temporary Visio file(s)', 'success');
                } else {
                    deleteButton.disabled = true;
                    showStatus('No temporary Visio files found', 'success');
                }
            } else {
                // Enhanced error handling matching CLI tool approach
                let errorMessage = 'Unknown error occurred';
                if (data.error) {
                    errorMessage = data.error;
                    if (data.details) {
                        errorMessage += ': ' + data.details;
                    }
                }
                showStatus('Error: ' + errorMessage, 'error');
            }
        } catch (error) {
            showStatus('Error: ' + error.message, 'error');
        } finally {
            setLoading(false);
        }
    }
    
    function displayFiles(filesToDisplay) { // Renamed parameter to avoid confusion with global 'files'
        clearFileList();

        if (filesToDisplay.length === 0) {
            const messageElement = document.createElement('li'); // Using 'li' to fit into the UL structure
            messageElement.className = 'empty-file-list-message';
            messageElement.textContent = 'No temporary Visio files found in this directory.';
            fileList.appendChild(messageElement);
        } else {
            filesToDisplay.forEach((fileData, index) => {
                const li = document.createElement('li');
            li.className = 'file-item';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = 'file-' + index;
            // Store the FullName for deletion, as that's what the backend expects
            checkbox.dataset.path = fileData.FullName; 
            checkbox.addEventListener('change', updateDeleteButton);
            
            const label = document.createElement('label');
            label.htmlFor = 'file-' + index;
            
            // Create a structure for title (Name) and subtitle (FullName)
            const titleSpan = document.createElement('span');
            titleSpan.className = 'file-title';
            titleSpan.textContent = fileData.Name; // Use Name for the title
            
            const pathSpan = document.createElement('span');
            pathSpan.className = 'file-path';
            pathSpan.textContent = fileData.FullName; // Use FullName for the subtitle
            pathSpan.setAttribute('title', fileData.FullName); // Add tooltip for full path
            
            label.appendChild(titleSpan);
            label.appendChild(pathSpan);
            
            li.appendChild(checkbox);
            li.appendChild(label);
                fileList.appendChild(li);
            });
        }
    }
    
    function toggleSelectAll() {
        const checkboxes = document.querySelectorAll('#fileList input[type="checkbox"]');
        checkboxes.forEach(checkbox => {
            checkbox.checked = selectAllCheckbox.checked;
        });
        
        updateDeleteButton();
    }
    
    function updateDeleteButton() {
        const anyChecked = document.querySelector('#fileList input[type="checkbox"]:checked');
        deleteButton.disabled = !anyChecked;
    }
    
    async function deleteSelectedFiles() {
        const checkboxes = document.querySelectorAll('#fileList input[type="checkbox"]:checked');
        const selectedFiles = Array.from(checkboxes).map(checkbox => checkbox.dataset.path);
        
        if (selectedFiles.length === 0) {
            showStatus('No files selected for deletion', 'error');
            return;
        }
        
        if (!confirm('Are you sure you want to delete ' + selectedFiles.length + ' file(s)?')) {
            return;
        }
        
        setLoading(true);
        
        try {
            const response = await fetch('/api/delete', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ files: selectedFiles })
            });
            
            const data = await response.json();
            
            if (response.ok) {
                showStatus(data.message, 'success');
                // Refresh the file list
                scanForFiles();
            } else {
                showStatus('Error: ' + data.error, 'error');
            }
        } catch (error) {
            showStatus('Error: ' + error.message, 'error');
        } finally {
            setLoading(false);
        }
    }
    
    function clearFileList() {
        fileList.innerHTML = '';
        deleteButton.disabled = true;
        selectAllCheckbox.checked = false;
    }
    
    function showStatus(message, type) {
        statusMessage.textContent = message;
        statusMessage.className = 'status-message ' + type;
        
        // Auto-hide success messages after 5 seconds
        if (type === 'success') {
            setTimeout(() => {
                statusMessage.textContent = '';
                statusMessage.className = 'status-message';
            }, 5000);
        }
    }
    
    function setLoading(isLoading) {
        if (isLoading) {
            document.body.classList.add('loading');
            // scanButton.disabled = true; // scanButton disabled state will be handled by its loading state visual
            deleteButton.disabled = true;
    
            // Enhance Scan Button
            originalScanButtonText = scanButton.textContent; // Save current text before changing
            scanButton.classList.add('button-loading');
            scanButton.disabled = true; // Also disable during loading
            scanButton.textContent = 'Scanning...'; // Change text
    
            // Create and append spinner
            if (!currentSpinner) { // Avoid adding multiple spinners
                currentSpinner = document.createElement('span');
                currentSpinner.className = 'spinner'; // Use the global spinner class
                scanButton.appendChild(currentSpinner);
            }
    
        } else {
            document.body.classList.remove('loading');
            // scanButton.disabled = false; // scanButton enabled state will be handled after loading
            updateDeleteButton(); // This already handles deleteButton's state
    
            // Restore Scan Button
            scanButton.classList.remove('button-loading');
            scanButton.disabled = false;
            scanButton.textContent = originalScanButtonText; // Restore original text
    
            // Remove spinner
            if (currentSpinner) {
                currentSpinner.remove();
                currentSpinner = null;
            }
        }
    }
    
    function toggleDarkMode() {
        if (darkModeToggle && darkModeToggle.checked) {
            document.body.classList.add('dark-mode');
            localStorage.setItem('darkMode', 'enabled');
        } else {
            document.body.classList.remove('dark-mode');
            localStorage.removeItem('darkMode');
        }
    }
    
    function toggleWideMode() {
        if (wideModeToggle.checked) {
            document.body.classList.add('wide-mode');
            localStorage.setItem('wideMode', 'enabled');
        } else {
            document.body.classList.remove('wide-mode');
            localStorage.removeItem('wideMode');
        }
    }
});
