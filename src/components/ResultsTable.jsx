import React from 'react';
import { Paper, Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Box } from '@mui/material';
import ResultsBarChart from './ResultsBarChart';

const ResultsTable = ({ results }) => {
  const formatNumber = (num) => {
    if (num === null || num === undefined) return 'N/A';
    return Number(num).toFixed(4);
  };

  // Calculate averages
  const validResults = results.filter(result => 
    result.extractableOilReserves !== null && 
    result.remainingOilReserves !== null
  );

  const averageExtrOil = validResults.length > 0 
    ? validResults.reduce((sum, r) => sum + r.extractableOilReserves, 0) / 4 
    : null;

  const averageRemOil = validResults.length > 0 
    ? validResults.reduce((sum, r) => sum + r.remainingOilReserves, 0) / 4 
    : null;

  return (
    <Box>
      <TableContainer component={Paper} sx={{ mb: 2 }}>
        <Table>
          <TableHead>
            <TableRow>
              <TableCell>Метод</TableCell>
              <TableCell align="right">V извлекаемые</TableCell>
              <TableCell align="right">V остаточные</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {results.map((result, index) => (
              <TableRow key={index}>
                <TableCell>{result.method}</TableCell>
                <TableCell align="right">{formatNumber(result.extractableOilReserves)}</TableCell>
                <TableCell align="right">{formatNumber(result.remainingOilReserves)}</TableCell>
              </TableRow>
            ))}
            <TableRow sx={{ backgroundColor: '#f5f5f5' }}>
              <TableCell sx={{ fontWeight: 'bold' }}>V среднее</TableCell>
              <TableCell align="right" sx={{ fontWeight: 'bold' }}>{formatNumber(averageExtrOil)}</TableCell>
              <TableCell align="right" sx={{ fontWeight: 'bold' }}>{formatNumber(averageRemOil)}</TableCell>
            </TableRow>
          </TableBody>
        </Table>
      </TableContainer>
      
      <Paper sx={{ p: 2 }}>
        <ResultsBarChart results={results} />
      </Paper>
    </Box>
  );
};

export default ResultsTable;
