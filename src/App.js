import React from 'react';
import ExcelJS from 'exceljs';

export default function App() {
  const handleDownload = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');

    worksheet.addRow(['Name', 'Age']);
    worksheet.addRow(['Alice', 30]);
    worksheet.addRow(['Bob', 25]);

    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'example.xlsx';
    a.click();
    window.URL.revokeObjectURL(url);
  };

  return (
    <div style={{ padding: 20 }}>
      <button onClick={handleDownload}>הורד קובץ Excel</button>
    </div>
  );
}
