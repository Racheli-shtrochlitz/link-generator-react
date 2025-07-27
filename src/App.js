import React, { useState } from 'react';
import { Button, Container, Typography, Box, TextField, MenuItem, Select, InputLabel, FormControl, Paper, CircularProgress, Alert } from '@mui/material';
import * as XLSX from 'xlsx';

function App() {
  const [file, setFile] = useState(null);
  const [originalFileName, setOriginalFileName] = useState('');
  const [columns, setColumns] = useState([]);
  const [sheetData, setSheetData] = useState([]);
  const [folderCol, setFolderCol] = useState('');
  const [fileCol, setFileCol] = useState('');
  const [rootPath, setRootPath] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState('');

  const handleFileUpload = (e) => {
    setError('');
    setDownloadUrl('');
    const f = e.target.files[0];
    if (!f) {
      console.log('No file selected.');
      return;
    }
    setFile(f);
    setOriginalFileName(f.name);
    console.log('Selected file:', f.name);
  
    const reader = new FileReader();
  
    reader.onload = (evt) => {
      try {
        const result = evt.target.result;
        console.log('FileReader result type:', typeof result);
        const data = new Uint8Array(result);
        console.log('Parsed Uint8Array from file:', data.slice(0, 10));
  
        const workbook = XLSX.read(data, { type: 'array' });
        console.log('Workbook parsed:', workbook);
  
        const ws = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
  
        if (json.length === 0) {
          setError('הקובץ ריק או לא תקין.');
          console.warn('Empty or invalid file.');
          return;
        }
  
        console.log('First row (columns):', json[0]);
        setColumns(json[0]);
        setSheetData(json);
      } catch (err) {
        console.error('Error while reading Excel file:', err);
        setError('שגיאה בקריאת הקובץ: ' + err.message);
      }
    };
  
    reader.onerror = (err) => {
      console.error('FileReader failed:', err);
      setError('שגיאה בקריאת הקובץ.');
    };
  
    reader.readAsArrayBuffer(f);
  };
  

  const handleGenerate = () => {
    setError('');
    setDownloadUrl('');
    if (!file || !folderCol || !fileCol || !rootPath) {
      setError('אנא מלא את כל השדות.');
      return;
    }
  
    setLoading(true);
    setTimeout(() => {
      try {
        const totalCols = columns.length;
        const colIdxFolder = totalCols - 1 - columns.indexOf(folderCol);
        const colIdxFile = totalCols - 1 - columns.indexOf(fileCol);
  
        if (colIdxFolder === -1 || colIdxFile === -1) {
          setError('בחירת עמודות לא תקינה.');
          setLoading(false);
          return;
        }
  
        const newData = sheetData.map(row => [...row].reverse());
  
        const ws = XLSX.utils.aoa_to_sheet(newData);
  
        for (let i = 1; i < newData.length; i++) {
          const row = newData[i];
          if (!row[colIdxFolder] || !row[colIdxFile]) continue;
  
          const folderName = row[colIdxFolder];          
          const fullFolderPath = `קלסר ${folderName}`;    
          let filename = row[colIdxFile];
          let fullFileName = filename;
          fullFileName = `${folderName}${filename}`;
  
          if (!/\.[a-zA-Z0-9]+$/.test(filename)) {
            fullFileName +='.pdf';
          }
  
          const link = `file://${rootPath.replace(/\\/g, '/')}/${fullFolderPath}/${fullFileName}`;
  
          const cellAddr = XLSX.utils.encode_cell({ r: i, c: colIdxFile });
  
          if (!ws[cellAddr]) ws[cellAddr] = { t: 's', v: filename };
          ws[cellAddr].l = { Target: link };
        }
  
        ws['!sheetViews'] = [{ rightToLeft: true }];
        const wb = XLSX.utils.book_new();
        wb.Workbook = { Views: [{ RTL: true }] };
        XLSX.utils.book_append_sheet(wb, ws, 'Links');
  
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        setDownloadUrl(url);
      } catch (e) {
        setError('שגיאה ביצירת הקובץ: ' + e.message);
      }
      setLoading(false);
    }, 300);
  };
  

  const generatedFileName = originalFileName
    ? originalFileName.replace(/\.[^/.]+$/, '') + '_עם קישורים.xlsx'
    : 'קובץ_עם_לינקים.xlsx';

  return (
    <Container maxWidth="sm" sx={{ mt: 4 }}>
      <Paper elevation={3} sx={{ p: 3, mt: 4 }}>
        <Typography variant="h5" gutterBottom align="center">
          מחולל לינקים לאקסל
        </Typography>
        <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <Button variant="contained" component="label">
            העלאת קובץ אקסל
            <input type="file" accept=".xlsx,.xls" hidden onChange={handleFileUpload} />
          </Button>
          {columns.length > 0 && (
            <>
              <FormControl fullWidth>
                <InputLabel>בחר עמודת שם תיקייה</InputLabel>
                <Select
                  value={folderCol}
                  label="בחר עמודת שם תיקייה"
                  onChange={e => setFolderCol(e.target.value)}
                >
                  {columns.map((col, idx) => (
                    <MenuItem key={idx} value={col}>{col}</MenuItem>
                  ))}
                </Select>
              </FormControl>
              <FormControl fullWidth>
                <InputLabel>בחר עמודת שם קובץ</InputLabel>
                <Select
                  value={fileCol}
                  label="בחר עמודת שם קובץ"
                  onChange={e => setFileCol(e.target.value)}
                >
                  {columns.map((col, idx) => (
                    <MenuItem key={idx} value={col}>{col}</MenuItem>
                  ))}
                </Select>
              </FormControl>
              <TextField
                label="נתיב תיקייה ראשית (למשל: C:/Files)"
                value={rootPath}
                onChange={e => setRootPath(e.target.value)}
                fullWidth
              />
              <Button variant="contained" color="primary" onClick={handleGenerate} disabled={loading}>
                צור קובץ עם לינקים
              </Button>
            </>
          )}
          {loading && <Box sx={{ display: 'flex', justifyContent: 'center' }}><CircularProgress /></Box>}
          {error && <Alert severity="error">{error}</Alert>}
          {downloadUrl && (
            <Button
              variant="outlined"
              color="success"
              href={downloadUrl}
              download={generatedFileName}
              sx={{ mt: 2 }}
            >
              הורדת קובץ אקסל עם קישורים
            </Button>
          )}
        </Box>
      </Paper>
    </Container>
  );
}

export default App;
