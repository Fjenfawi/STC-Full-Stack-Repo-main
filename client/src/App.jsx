import React, { useState } from 'react';
import ExcelUploader from './components/ExcelUploader';
import './App.css';

function App() {
  const [department, setDepartment] = useState('');
  const [location, setLocation] = useState('');
  const [excelData, setExcelData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [submitted, setSubmitted] = useState(false); // ✅ To prevent re-submission

  const handleSubmit = async (e) => {
    e.preventDefault();

    if (!department || !location || !excelData || !excelData.file) {
      alert('Please fill all fields and upload a valid Excel file.');
      return;
    }

    setLoading(true);
    setSubmitted(true); // ✅ Prevent further submissions

    const formData = new FormData();
    formData.append('file', excelData.file);
    formData.append('department', department);
    formData.append('location', location);

    try {
      const response = await fetch('http://localhost:5001/api/excel/modify', {
        method: 'POST',
        body: formData,
      });

      if (response.ok) {
        const blob = await response.blob();
        const downloadUrl = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.download = `${location} ${department} VA Trackersheet.xlsx`;
        document.body.appendChild(link);
        link.click();
        link.remove();

        // ✅ Reset form after successful download
        setDepartment('');
        setLocation('');
        setExcelData(null);
        setSubmitted(false); // Allow new uploads after this one
      } else {
        alert('Server error.');
        setSubmitted(false); // Allow retry
      }
    } catch (err) {
      console.error('Submission failed', err);
      alert('Submission failed.');
      setSubmitted(false); // Allow retry
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app-container">
      <div id="header"><div id="stc">STC</div></div>
      <pre>This is a tool to automate the creation and styling of VA trackersheet. This is the first sprint we'll enhance it later.</pre>
      <div className="form-container">
        <form onSubmit={handleSubmit}>
          <div className="form-group">
            <label htmlFor="department">Choose Department:</label>
            <select
              id="department"
              value={department}
              onChange={(e) => setDepartment(e.target.value)}
              disabled={loading}
            >
              <option value="">Select...</option>
              <option value="IT">IT</option>
              <option value="Network">Network</option>
            </select>
          </div>

          <div className="form-group">
            <label htmlFor="location">Choose Location:</label>
            <select
              id="location"
              value={location}
              onChange={(e) => setLocation(e.target.value)}
              disabled={loading}
            >
              <option value="">Select...</option>
              <option value="External">External</option>
              <option value="Internal">Internal</option>
            </select>
          </div>

          <ExcelUploader onDataParsed={setExcelData} disabled={loading} />

          <button type="submit" style={{ marginTop: '20px' }} disabled={loading || submitted}>
            {loading ? 'Processing...' : 'Submit'}
          </button>

          {loading && (
            <div className="spinner-container">
              <div className="spinner"></div>
              <p>Processing file, please wait...</p>
            </div>
          )}
        </form>
      </div>
    </div>
  );
}

export default App;
