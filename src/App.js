import React from 'react';
import { Button, Container } from '@mui/material';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

function App() {
  const handleDownload = async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'בדיקה';

    try {
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      saveAs(blob, 'בדיקה.xlsx');
      console.log('✅ Created & Downloaded!');
    } catch (err) {
      console.error('❌ יצירת קובץ נכשלה:', err);
    }
  };

  return (
    <Container sx={{ mt: 4 }}>
      <Button variant="contained" onClick={handleDownload}>
        הורד קובץ בדיקה
      </Button>
    </Container>
  );
}

export default App;
