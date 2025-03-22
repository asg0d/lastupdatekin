import React from 'react';
import { Box, Typography, Paper } from '@mui/material';

const ConsoleOutput = ({ result }) => {
  const formatNumber = (num) => {
    if (num === null || num === undefined) return 'N/A';
    return Number(num).toFixed(4);
  };

  const formatCoefficient = (num) => {
    if (num === null || num === undefined) return 'N/A';
    const formatted = Number(num).toFixed(12);
    // If the number is 0, return with 12 zeros after decimal
    if (Number(formatted) === 0) {
      return '0.000000000000';
    }
    return formatted;
  };

  if (!result) return null;

  return (
    <Box>
      <Typography variant="h6" sx={{ mb: 2, color: '#1976d2' }}>
        {result.method}
      </Typography>

      {/* Входные данные */}
      <Paper sx={{ p: 2, mb: 2 }}>
        <Typography variant="subtitle1" sx={{ mb: 1, fontWeight: 'bold' }}>
          Входные данные:
        </Typography>
        <Box sx={{ fontFamily: 'monospace' }}>
          <div>X: {result.xDescription}</div>
          <div>Y: {result.yDescription}</div>
        </Box>
      </Paper>

      {/* Суммы */}
      <Paper sx={{ p: 2, mb: 2 }}>
        <Typography variant="subtitle1" sx={{ mb: 1, fontWeight: 'bold' }}>
          Суммы:
        </Typography>
        <Box sx={{ fontFamily: 'monospace' }}>
          <div>Σx = {formatNumber(result.sums.sumX)}</div>
          <div>Σy = {formatNumber(result.sums.sumY)}</div>
          <div>Σxy = {formatNumber(result.sums.sumXY)}</div>
          <div>Σx² = {formatNumber(result.sums.sumX2)}</div>
          <div>(Σx)² = {formatNumber(result.sums.sumXSquared)}</div>
        </Box>
      </Paper>

      {/* Коэффициенты */}
      <Paper sx={{ p: 2, mb: 2 }}>
        <Typography variant="subtitle1" sx={{ mb: 1, fontWeight: 'bold' }}>
          Коэффициенты:
        </Typography>
        <Box sx={{ fontFamily: 'monospace' }}>
          <div>A = {formatCoefficient(result.coefficients.A)}</div>
          <div>B = {formatCoefficient(result.coefficients.B)}</div>
        </Box>
      </Paper>

      {/* Результаты */}
      <Paper sx={{ p: 2, mb: 2 }}>
        <Typography variant="subtitle1" sx={{ mb: 1, fontWeight: 'bold' }}>
          Результаты:
        </Typography>
        <Box sx={{ fontFamily: 'monospace' }}>
          <div>V извлекаемые = {formatNumber(result.extractableOilReserves)}</div>
          <div>V остаточные = {formatNumber(result.remainingOilReserves)}</div>
        </Box>
      </Paper>

      {/* Таблица результатов */}
      {result.results && (
        <Paper sx={{ p: 2 }}>
          <Typography variant="subtitle1" sx={{ mb: 1, fontWeight: 'bold' }}>
            Таблица результатов:
          </Typography>
          <Box sx={{ overflowX: 'auto' }}>
            <table style={{ 
              width: '100%', 
              borderCollapse: 'collapse', 
              fontFamily: 'monospace'
            }}>
              <thead>
                <tr>
                  <th style={{ padding: '8px', textAlign: 'right', borderBottom: '2px solid #e0e0e0' }}>Год</th>
                  <th style={{ padding: '8px', textAlign: 'right', borderBottom: '2px solid #e0e0e0' }}>X</th>
                  <th style={{ padding: '8px', textAlign: 'right', borderBottom: '2px solid #e0e0e0' }}>Y</th>
                  <th style={{ padding: '8px', textAlign: 'right', borderBottom: '2px solid #e0e0e0' }}>XY</th>
                  <th style={{ padding: '8px', textAlign: 'right', borderBottom: '2px solid #e0e0e0' }}>X²</th>
                </tr>
              </thead>
              <tbody>
                {result.results.map((row, index) => {
                  const x = parseFloat(row.x) || 0;
                  const y = parseFloat(row.y) || 0;
                  const xy = x * y;
                  const x2 = x * x;
                  return (
                    <tr key={index}>
                      <td style={{ padding: '8px', textAlign: 'right', borderBottom: '1px solid #e0e0e0' }}>{row.year}</td>
                      <td style={{ padding: '8px', textAlign: 'right', borderBottom: '1px solid #e0e0e0' }}>{formatNumber(x)}</td>
                      <td style={{ padding: '8px', textAlign: 'right', borderBottom: '1px solid #e0e0e0' }}>{formatNumber(y)}</td>
                      <td style={{ padding: '8px', textAlign: 'right', borderBottom: '1px solid #e0e0e0' }}>{formatNumber(xy)}</td>
                      <td style={{ padding: '8px', textAlign: 'right', borderBottom: '1px solid #e0e0e0' }}>{formatNumber(x2)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </Box>
        </Paper>
      )}
    </Box>
  );
};

export default ConsoleOutput;
