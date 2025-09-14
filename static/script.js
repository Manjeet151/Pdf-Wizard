document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('uploadForm');
    const fileInput = document.getElementById('fileInput');
    const dropZone = document.getElementById('dropZone');
    const selectedFileContainer = document.getElementById('selectedFileContainer');
    const selectedFileName = document.getElementById('selectedFileName');
    const removeFileBtn = document.getElementById('removeFileBtn');
    const convertBtn = document.getElementById('convertBtn');
    const spinner = document.getElementById('spinner');
    const notification = document.getElementById('notification');
    const notificationMessage = document.getElementById('notificationMessage');

    let selectedFile = null;

    // Drag and drop functionality
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, unhighlight, false);
    });

    function highlight() {
        dropZone.classList.add('drag-over');
    }

    function unhighlight() {
        dropZone.classList.remove('drag-over');
    }

    dropZone.addEventListener('drop', handleDrop, false);

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length) {
            handleFiles(files);
        }
    }

    fileInput.addEventListener('change', function() {
        handleFiles(this.files);
    });

    function handleFiles(files) {
        if (files.length > 0) {
            selectedFile = files[0];
            updateSelectedFileUI();
        }
    }

    function updateSelectedFileUI() {
        if (selectedFile) {
            selectedFileName.textContent = selectedFile.name;
            selectedFileContainer.style.display = 'flex';
            convertBtn.disabled = false;
        } else {
            selectedFileContainer.style.display = 'none';
            convertBtn.disabled = true;
        }
    }

    removeFileBtn.addEventListener('click', function() {
        selectedFile = null;
        fileInput.value = '';
        updateSelectedFileUI();
    });

    form.addEventListener('submit', function(e) {
        e.preventDefault();
        
        if (!selectedFile) {
            showNotification('Please select a file first', 'error');
            return;
        }

        // Check file size (16MB limit)
        if (selectedFile.size > 16 * 1024 * 1024) {
            showNotification('File size exceeds 16MB limit', 'error');
            return;
        }

        // Get selected file type
        const fileType = document.querySelector('input[name="fileType"]:checked').value;
        
        // Start conversion
        convertFile(selectedFile, fileType);
    });

    function convertFile(file, fileType) {
        // Show loading state
        convertBtn.disabled = true;
        spinner.classList.remove('hidden');
        
        const formData = new FormData();
        formData.append('file', file);
        formData.append('type', fileType);
        
        fetch('/convert', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                return response.json().then(errorData => {
                    throw new Error(errorData.error || 'Conversion failed');
                });
            }
            return response.blob();
        })
        .then(blob => {
            // Create download link
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = file.name.split('.')[0] + '.pdf';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
            showNotification('File converted successfully!', 'success');
        })
        .catch(error => {
            console.error('Error:', error);
            showNotification(error.message || 'Conversion failed. Please try again.', 'error');
        })
        .finally(() => {
            // Reset loading state
            convertBtn.disabled = false;
            spinner.classList.add('hidden');
        });
    }

    function showNotification(message, type = 'error') {
        notificationMessage.textContent = message;
        notification.className = 'notification';
        
        if (type === 'success') {
            notification.classList.add('success');
        }
        
        // Show notification
        setTimeout(() => {
            notification.classList.remove('hidden');
        }, 10);
        
        // Hide after 5 seconds
        setTimeout(() => {
            notification.classList.add('hidden');
        }, 5000);
    }
});