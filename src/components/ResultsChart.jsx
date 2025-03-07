import React, { useState } from 'react';
import { Dialog, DialogTitle, DialogContent, IconButton, Box, FormControl, InputLabel, Select, MenuItem } from '@mui/material';
import CloseIcon from '@mui/icons-material/Close';
import { ScatterChart, Scatter, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';

const COLORS = [
  '#1976d2', // blue
  '#dc004e', // red
  '#388e3c', // green
  '#f57c00', // orange
  '#6d1b7b', // purple
  '#1565c0'  // dark blue
];

const CustomTooltip = ({ active, payload }) => {
  if (active && payload && payload.length) {
    const data = payload[0].payload;
    return (
      <div style={{ 
        backgroundColor: 'white', 
        padding: '10px', 
        border: '1px solid #ccc',
        borderRadius: '4px'
      }}>
        <p><strong>Год:</strong> {data.year}</p>
        <p><strong>X:</strong> {data.x.toFixed(4)}</p>
        <p><strong>Y:</strong> {data.y.toFixed(4)}</p>
      </div>
    );
  }
  return null;
};

const ResultsChart = ({ open, onClose, results }) => {
  const [selectedMethod, setSelectedMethod] = useState('all');

  if (!results || !results.length) return null;

  const methodsWithData = results.filter(method => 
    method.results && method.results.length > 0
  );

  const displayedMethods = selectedMethod === 'all' 
    ? methodsWithData 
    : methodsWithData.filter(m => m.method === selectedMethod);

  return (
    <Dialog 
      open={open} 
      onClose={onClose}
      maxWidth="lg"
      fullWidth
      PaperProps={{
        sx: {
          minHeight: '80vh',
          maxHeight: '90vh'
        }
      }}
    >
      <DialogTitle sx={{ 
        m: 0, 
        p: 2, 
        display: 'flex', 
        justifyContent: 'space-between', 
        alignItems: 'center',
        bgcolor: '#f8f9fa',
        borderBottom: '1px solid #e0e0e0'
      }}>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 2, flex: 1 }}>
          <span>График X-Y</span>
          <FormControl size="small" sx={{ minWidth: 200 }}>
            <InputLabel>Метод</InputLabel>
            <Select
              value={selectedMethod}
              label="Метод"
              onChange={(e) => setSelectedMethod(e.target.value)}
            >
              <MenuItem value="all">Все методы</MenuItem>
              {methodsWithData.map((method, index) => (
                <MenuItem key={method.method} value={method.method}>
                  {method.method}
                </MenuItem>
              ))}
            </Select>
          </FormControl>
        </Box>
        <IconButton
          aria-label="close"
          onClick={onClose}
          sx={{
            color: '#666'
          }}
        >
          <CloseIcon />
        </IconButton>
      </DialogTitle>
      <DialogContent sx={{ p: 3 }}>
        <Box sx={{ height: '70vh' }}>
          <ResponsiveContainer width="100%" height="100%">
            <ScatterChart
              margin={{
                top: 20,
                right: 30,
                left: 20,
                bottom: 20,
              }}
            >
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis 
                type="number"
                dataKey="x"
                name="X"
                label={{ 
                  value: 'X', 
                  position: 'bottom', 
                  offset: -5 
                }}
              />
              <YAxis 
                type="number"
                dataKey="y"
                name="Y"
                label={{ 
                  value: 'Y', 
                  angle: -90, 
                  position: 'insideLeft',
                  offset: -5
                }}
              />
              <Tooltip content={<CustomTooltip />} />
              <Legend />
              {displayedMethods.map((method, index) => (
                <Scatter
                  key={method.method}
                  name={method.method}
                  data={method.results.map(r => ({
                    x: r.x,
                    y: r.y,
                    year: r.year
                  }))}
                  fill={COLORS[index % COLORS.length]}
                  line
                  shape="circle"
                />
              ))}
            </ScatterChart>
          </ResponsiveContainer>
        </Box>
      </DialogContent>
    </Dialog>
  );
};

export default ResultsChart;
