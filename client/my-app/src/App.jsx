import React, { useState } from 'react';

// The main App component for the file upload UI.
export default function App() {
  // State to hold the selected file.
  const [file, setFile] = useState(null);
  // State to hold the upload message (e.g., success, error).
  const [message, setMessage] = useState('');
  // State to track the status (success, error, or neutral) for styling.
  const [status, setStatus] = useState('neutral');

  /**
   * Handles the change event when a file is selected.
   * Updates the `file` state with the first selected file.
   * Clears the previous message and status.
   * @param {object} event - The file input change event.
   */
  const handleFileChange = (event) => {
    setFile(event.target.files[0]);
    setMessage('');
    setStatus('neutral');
  };

  /**
   * Handles the file upload process.
   * Sends a POST request to the specified API endpoint with the selected file.
   */
  const handleUpload = async () => {
    // Check if a file has been selected before attempting to upload.
    if (!file) {
      setMessage('Please select a file first.');
      setStatus('error');
      return;
    }

    // Use FormData to prepare the file for the POST request.
    const formData = new FormData();
    formData.append('file', file);

    try {
      // Perform the fetch request to the Spring Boot backend.
      const response = await fetch('http://localhost:8080/api/excel/upload', {
        method: 'POST',
        body: formData,
      });

      // Check the response status. A 200 status indicates success.
      if (response.ok) { // `response.ok` is true for 200-299 status codes.
        setMessage('Success! File uploaded successfully.');
        setStatus('success');
      } else {
        // Handle non-200 responses, providing a specific error message.
        const errorText = await response.text();
        setMessage(`Error: ${response.status} - ${errorText}`);
        setStatus('error');
      }
    } catch (error) {
      // Handle network errors or other exceptions during the fetch call.
      setMessage('An error occurred. Please check your network connection.');
      setStatus('error');
    }
  };

  return (
    // Tailwind classes for a centered, full-screen layout.
    <div className="flex items-center justify-center min-h-screen bg-gray-100 p-4">
      <div className="w-full max-w-md p-8 bg-white rounded-xl shadow-xl">
        <h1 className="text-2xl font-bold text-center mb-6 text-gray-800">
          File Uploader
        </h1>

        {/* File input and upload button container */}
        <div className="flex flex-col space-y-4">
          <label className="block text-sm font-medium text-gray-700">
            Select a file to upload:
          </label>
          <input
            type="file"
            onChange={handleFileChange}
            className="block w-full text-sm text-gray-500
              file:mr-4 file:py-2 file:px-4
              file:rounded-full file:border-0
              file:text-sm file:font-semibold
              file:bg-blue-50 file:text-blue-700
              hover:file:bg-blue-100 cursor-pointer"
          />

          {/* Display the selected file name */}
          {file && (
            <p className="text-sm text-gray-600">
              Selected file: <span className="font-semibold">{file.name}</span>
            </p>
          )}

          <button
            onClick={handleUpload}
            className={`
              w-full px-4 py-2 mt-4 text-white font-semibold rounded-lg shadow-md
              transition-colors duration-200 ease-in-out
              ${!file ? 'bg-gray-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700'}
            `}
            disabled={!file}
          >
            Upload File
          </button>
        </div>

        {/* Message area to display upload status */}
        {message && (
          <div
            className={`p-4 mt-6 rounded-lg text-center font-medium
              ${status === 'success' ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'}
            `}
          >
            {message}
          </div>
        )}
      </div>
    </div>
  );
}

