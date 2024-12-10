import React, { useState } from 'react';
import { 
  Box, 
  Container, 
  Typography, 
  Button, 
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  TextField
} from '@mui/material';
import * as XLSX from 'xlsx';

interface TableData {
  headers: string[];
  rows: any[][];
}

function App() {
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [tableData, setTableData] = useState<TableData | null>(null);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [images, setImages] = useState<string[]>([]);

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      
      setSheetNames(workbook.SheetNames);
      if (workbook.SheetNames.length > 0) {
        handleSheetSelect(workbook.SheetNames[0], workbook);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleSheetSelect = (sheetName: string, workbook?: XLSX.WorkBook) => {
    const wb = workbook || XLSX.read(sheetName, { type: 'string' });
    const worksheet = wb.Sheets[sheetName];
    
    // Convert sheet data to array
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (jsonData.length > 0) {
      setTableData({
        headers: jsonData[0] as string[],
        rows: jsonData.slice(1) as any[][]
      });
    }
    
    setSelectedSheet(sheetName);
  };

  const filterData = (data: TableData) => {
    if (!searchTerm) return data;

    const filteredRows = data.rows.filter(row =>
      row.some(cell => 
        String(cell).toLowerCase().includes(searchTerm.toLowerCase())
      )
    );

    return {
      headers: data.headers,
      rows: filteredRows
    };
  };

  return (
    <Container maxWidth="lg">
      <Box sx={{ my: 4 }}>
        <Typography variant="h4" component="h1" gutterBottom>
          Excel Viewer
        </Typography>

        <Button
          variant="contained"
          component="label"
          sx={{ mb: 2 }}
        >
          Upload Excel File
          <input
            type="file"
            hidden
            accept=".xlsx"
            onChange={handleFileUpload}
          />
        </Button>

        {sheetNames.length > 0 && (
          <Box sx={{ mb: 2 }}>
            <Typography variant="h6">Select Sheet:</Typography>
            {sheetNames.map(name => (
              <Button
                key={name}
                onClick={() => handleSheetSelect(name)}
                variant={selectedSheet === name ? "contained" : "outlined"}
                sx={{ mr: 1 }}
              >
                {name}
              </Button>
            ))}
          </Box>
        )}

        {tableData && (
          <>
            <TextField
              fullWidth
              label="Search"
              variant="outlined"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              sx={{ mb: 2 }}
            />

            <TableContainer component={Paper}>
              <Table>
                <TableHead>
                  <TableRow>
                    {tableData.headers.map((header, index) => (
                      <TableCell key={index}>{header}</TableCell>
                    ))}
                  </TableRow>
                </TableHead>
                <TableBody>
                  {filterData(tableData).rows.map((row, rowIndex) => (
                    <TableRow key={rowIndex}>
                      {row.map((cell, cellIndex) => (
                        <TableCell key={cellIndex}>
                          {String(cell)}
                        </TableCell>
                      ))}
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </TableContainer>
          </>
        )}
      </Box>
    </Container>
  );
}

export default App; 