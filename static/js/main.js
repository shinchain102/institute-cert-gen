document.addEventListener('DOMContentLoaded', function() {
    let templateVariables = [];
    let dataColumns = [];
    let templateFile = '';
    let dataFile = '';

    // Template Upload
    document.getElementById('uploadTemplate').addEventListener('click', function() {
        const fileInput = document.getElementById('templateFile');
        const file = fileInput.files[0];
        if (!file) {
            showAlert('Please select a template file', 'danger');
            return;
        }

        const formData = new FormData();
        formData.append('template', file);

        fetch('/upload-template', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.error) {
                throw new Error(data.error);
            }
            templateVariables = data.variables;
            templateFile = data.filename;  // Store the secure filename
            showAlert('Template uploaded successfully', 'success');
            displayTemplateVariables(templateVariables);
            updateMappingSection();
        })
        .catch(error => {
            showAlert(error.message, 'danger');
        });
    });

    // Data File Upload
    document.getElementById('uploadData').addEventListener('click', function() {
        const fileInput = document.getElementById('dataFile');
        const file = fileInput.files[0];
        if (!file) {
            showAlert('Please select a data file', 'danger');
            return;
        }

        const formData = new FormData();
        formData.append('data', file);

        fetch('/upload-data', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.error) {
                throw new Error(data.error);
            }
            dataColumns = data.columns;
            dataFile = data.filename;  // Store the secure filename
            showAlert('Data file uploaded successfully', 'success');
            displayDataColumns(dataColumns);
            updateMappingSection();
        })
        .catch(error => {
            showAlert(error.message, 'danger');
        });
    });

    // Generate Certificates
    document.getElementById('generateBtn').addEventListener('click', function() {
        if (!templateFile || !dataFile) {
            showAlert('Please upload both template and data files', 'danger');
            return;
        }

        const mapping = {};
        const mappingFields = document.querySelectorAll('.mapping-select');
        let isValid = true;

        mappingFields.forEach(field => {
            if (!field.value) {
                isValid = false;
                showAlert('Please map all template variables', 'danger');
                return;
            }
            mapping[field.dataset.variable] = field.value;
        });

        if (!isValid) return;

        // Show progress
        const progressBar = document.getElementById('progressBar');
        progressBar.style.display = 'block';
        const generateBtn = document.getElementById('generateBtn');
        generateBtn.disabled = true;

        fetch('/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                mapping: mapping,
                template_file: templateFile,
                data_file: dataFile
            })
        })
        .then(response => {
            if (!response.ok) {
                return response.json().then(data => {
                    throw new Error(data.error || 'Certificate generation failed');
                });
            }
            return response.blob();
        })
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'certificates.zip';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            showAlert('Certificates generated successfully', 'success');
        })
        .catch(error => {
            console.error('Generation error:', error);
            showAlert(error.message || 'Failed to generate certificates. Please try again.', 'danger');
        })
        .finally(() => {
            progressBar.style.display = 'none';
            generateBtn.disabled = false;
        });
    });

    function showAlert(message, type) {
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type} alert-dismissible fade show`;
        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;
        document.querySelector('.card-body').insertBefore(alertDiv, document.querySelector('.card-body').firstChild);
    }

    function displayTemplateVariables(variables) {
        const container = document.getElementById('templateVariables');
        container.innerHTML = `
            <div class="alert alert-info">
                <small>Template variables found: ${variables.join(', ')}</small>
            </div>
        `;
    }

    function displayDataColumns(columns) {
        const container = document.getElementById('dataColumns');
        container.innerHTML = `
            <div class="alert alert-info">
                <small>Data columns found: ${columns.join(', ')}</small>
            </div>
        `;
    }

    function updateMappingSection() {
        if (templateVariables.length > 0 && dataColumns.length > 0) {
            const mappingSection = document.getElementById('mappingSection');
            const mappingFields = document.getElementById('mappingFields');
            mappingSection.style.display = 'block';

            mappingFields.innerHTML = templateVariables.map(variable => `
                <div class="mb-3">
                    <label class="form-label">Map ${variable} to:</label>
                    <select class="form-select mapping-select" data-variable="${variable}">
                        <option value="">Select column</option>
                        ${dataColumns.map(column => `
                            <option value="${column}">${column}</option>
                        `).join('')}
                    </select>
                </div>
            `).join('');

            document.getElementById('generateBtn').disabled = false;
        }
    }
});