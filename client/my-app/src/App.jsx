import React, { useState } from 'react';

// The main App component for the file upload UI.
export default function App() {
  const [file, setFile] = useState(null);
  const [message, setMessage] = useState('');
  const [status, setStatus] = useState('neutral');

  const handleFileChange = (event) => {
    setFile(event.target.files[0]);
    setMessage('');
    setStatus('neutral');
  };

  const handleUpload = async () => {
    if (!file) {
      setMessage('Please select a file first.');
      setStatus('error');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);

    // Use relative path so Nginx (in frontend container) can proxy to backend:
    const targetUrl = `/api/excel/upload`;

    try {
      const response = await fetch(targetUrl, {
        method: 'POST',
        body: formData,
      });

      if (response.ok) {
        setMessage('Success! File uploaded successfully.');
        setStatus('success');
      } else {
        const errText = await response.text().catch(() => '');
        setMessage(`Error: ${response.status}${errText ? ' - ' + errText : ''}`);
        setStatus('error');
      }
    } catch (error) {
      setMessage('An error occurred. Please check your network connection.');
      setStatus('error');
      // console.error('Upload error', error);
    }
  };

  return (
    <div className="flex items-center justify-center min-h-screen bg-gray-100 p-4">
      <div className="w-full max-w-md p-8 bg-white rounded-xl shadow-xl">
        <h1 className="text-2xl font-bold text-center mb-6 text-gray-800">
          File Uploader
        </h1>

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
