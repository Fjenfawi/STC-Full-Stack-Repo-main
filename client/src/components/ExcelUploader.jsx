import React, { useCallback, useState } from 'react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';

const ExcelUploader = ({ onDataParsed }) => {
  const [fileName, setFileName] = useState('');

  const onDrop = useCallback((acceptedFiles) => {
    const file = acceptedFiles[0];
    if (!file) return;
  
    setFileName(file.name);
  
    const reader = new FileReader();
    reader.onload = (e) => {
      const binaryStr = e.target.result;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  
      // Pass both the file and the parsed data to the parent
      onDataParsed({
        file,      // File object for backend submission
        data: jsonData // Parsed data for preview if needed
      });
    };
  
    reader.readAsBinaryString(file);
  }, [onDataParsed]);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: ['.xlsx', '.xls']
  });

  return (
    <div>
      <div {...getRootProps()} style={{
        border: '2px dashed #ccc',
        padding: '20px',
        marginTop: '20px',
        textAlign: 'center',
        background: isDragActive ? '#f0f8ff' : '#fafafa',
        cursor: 'pointer'
      }}>
        <input {...getInputProps()} />
        <p>{isDragActive ? 'Drop the Excel file here...' : 'Click or drag an Excel file'}</p>
      </div>

      {fileName && (
        <div style={{ marginTop: '10px', display: 'flex', alignItems: 'center' }}>
          <img
            src="https://cdn-icons-png.flaticon.com/512/732/732220.png"
            alt="Excel Icon"
            style={{ width: 24, height: 24, marginRight: 10 }}
          />
          <span>{fileName}</span>
        </div>
      )}
    </div>
  );
};

export default ExcelUploader;
