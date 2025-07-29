import React, { useState } from 'react';
import { Button, Container, Typography, Box, TextField, MenuItem, Select, InputLabel, FormControl, Paper, CircularProgress, Alert } from '@mui/material';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

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

  const handleFileUpload = async (e) => {
    setError('');
    setDownloadUrl('');
    const f = e.target.files[0];

    if (!f) {
      console.warn('❗ לא נבחר קובץ');
      return;
    }

    if (!(f instanceof Blob)) {
      setError('הקובץ אינו תקין (לא Blob)');
      return;
    }

    if (!f.name.endsWith('.xlsx') && !f.name.endsWith('.xls')) {
      setError('נא לבחור קובץ Excel בלבד (סיומת .xlsx או .xls)');
      return;
    }

    setFile(f);
    setOriginalFileName(f.name);
    console.log('📄 קובץ שנבחר:', f.name, '| גודל:', f.size, '| סוג:', f.type);

    try {
      const arrayBuffer = await f.arrayBuffer();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      const worksheet = workbook.worksheets[0];
      if (!worksheet) {
        setError('הגיליון ריק או לא נמצא.');
        return;
      }

      const json = [];
      worksheet.eachRow((row, rowNumber) => {
        const rowValues = row.values;
        json.push(rowValues.slice(1));
      });

      if (json.length === 0) {
        setError('הקובץ ריק או לא תקין.');
        return;
      }

      setColumns(json[0]);
      setSheetData(json);
      console.log('✅ ExcelJS נטען בהצלחה, כותרות:', json[0]);
    } catch (e) {
      console.error('❌ שגיאה בקריאת הקובץ עם ExcelJS:', e);
      setError('שגיאה בקריאת קובץ ה־Excel: ' + e.message);
    }
  };

  const handleGenerate = async () => {
    setError('');
    setDownloadUrl('');
    if (!file || !folderCol || !fileCol || !rootPath) {
      setError('אנא מלא את כל השדות.');
      return;
    }

    setLoading(true);
    setTimeout(async () => {
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

        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Links', { properties: { tabColor: { argb: 'FFC0000' } } });

        newData.forEach((row, idx) => {
          ws.addRow(row);
        });

        for (let i = 1; i < newData.length; i++) {
          const row = newData[i];
          if (!row[colIdxFolder] || !row[colIdxFile]) continue;

          const folderName = row[colIdxFolder];
          const fullFolderPath = `קלסר ${folderName}`;
          let filename = row[colIdxFile];
          let fullFileName = filename;
          fullFileName = `${folderName}${filename}`;

          if (!/\.[a-zA-Z0-9]+$/.test(filename)) {
            fullFileName += '.pdf';
          }

          const link = `file://${rootPath.replace(/\\/g, '/')}/${fullFolderPath}/${fullFileName}`;

          const excelRow = ws.getRow(i + 1); 
          const cell = excelRow.getCell(colIdxFile + 1);

          cell.value = {
            text: filename,
            hyperlink: link,
          };
        }

        ws.views = [{ rightToLeft: true }];

        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/octet-stream' });
        saveAs(blob, generatedFileName);

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
          {originalFileName && (
            <Typography variant="body2" sx={{ mt: 1 }}>
              קובץ נבחר: {originalFileName}
            </Typography>
          )}
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
