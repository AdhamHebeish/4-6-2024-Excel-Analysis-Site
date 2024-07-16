const submitButton = document.getElementById('submitButton');
const errorMessage = document.getElementById('error-message');

submitButton.addEventListener('click', async () => {
    const fileUpload = document.getElementById('fileUpload');
    const file = fileUpload.files[0];

    if (!file) {
        errorMessage.textContent = 'Please select a file to upload.';
        return;
    }

    errorMessage.textContent = '';  // Clear previous error messages

    try {
        const response = await fetch('/analyze_excel', {
            method: 'POST',
            body: new FormData(document.getElementById('uploadForm')),  // Use form ID
        });

        if (!response.ok) {
            const errorData = await response.json();
            errorMessage.textContent = errorData.error;
            return;
        }

        // No need for table population code here (handled by Python)
    } catch (error) {
        console.error(error);
        errorMessage.textContent = 'An error occurred during upload.';
    }
});