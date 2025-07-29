import React, { useState } from 'react';
import { Button, Container } from '@mui/material';
import ExcelJS from 'exceljs';

function App() {
  const [file, setFile] = useState(null);

  const handleFileUpload = async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);

    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(await f.arrayBuffer());
      console.log('✅ Loaded!');
    } catch (err) {
      console.error('❌ ExcelJS failed to load file:', err);
    }
  };

  return (
    <Container sx={{ mt: 4 }}>
      <Button variant="contained" component="label">
        העלאת קובץ
        <input type="file" hidden onChange={handleFileUpload} accept=".xlsx,.xls" />
      </Button>
    </Container>
  );
}

export default App;
