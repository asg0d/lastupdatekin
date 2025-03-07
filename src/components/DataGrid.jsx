import React from 'react';
import { 
  Box, 
  Typography, 
  TextField, 
  IconButton, 
  TableContainer, 
  Table, 
  TableHead, 
  TableBody, 
  TableRow, 
  TableCell, 
  Paper, 
  Checkbox
} from '@mui/material';
import DeleteIcon from '@mui/icons-material/Delete';
import AddIcon from '@mui/icons-material/Add';

const DataGrid = ({ data, setData }) => {
  const handleChange = (index, field, value) => {
    const newData = [...data];
    newData[index] = {
      ...newData[index],
      [field]: value
    };

    // If oil or liquid changes, recalculate water and water cut
    if (field === 'oil' || field === 'liquid') {
      const oil = parseFloat(field === 'oil' ? value : newData[index].oil) || 0;
      const liquid = parseFloat(field === 'liquid' ? value : newData[index].liquid) || 0;
      const water = Math.max(0, liquid - oil);
      newData[index].water = water.toFixed(2);
      
      // Calculate water cut using year-over-year change
      if (index > 0) {
        const prevRow = newData[index - 1];
        const prevLiquid = parseFloat(prevRow.liquid) || 0;
        const prevOil = parseFloat(prevRow.oil) || 0;
        const prevWater = Math.max(0, prevLiquid - prevOil);
        
        const waterChange = water - prevWater;
        const liquidChange = liquid - prevLiquid;
        
        if (liquidChange !== 0) {
          newData[index].waterCut = ((waterChange / liquidChange) * 100).toFixed(2);
        }
      }
    }
    
    setData(newData);
  };

  const addRow = () => {
    setData([...data, { year: '', water: '', oil: '', liquid: '', active: false, waterCut: '' }]);
  };

  const deleteRow = (index) => {
    const newData = data.filter((_, i) => i !== index);
    setData(newData);
  };

  const toggleActive = (index) => {
    const newData = [...data];
    newData[index] = { ...newData[index], active: !newData[index].active };
    setData(newData);
  };

  return (
    <Box sx={{ width: '100%', overflow: 'auto' }}>
      <Typography 
        variant="h6" 
        sx={{ 
          mb: 2, 
          color: '#1976d2',
          fontWeight: 500
        }}
      >
        Входные данные
      </Typography>
      <TableContainer component={Paper} sx={{ height: '100%', overflow: 'auto' }}>
        <Table 
          size="small" 
          stickyHeader 
          sx={{
            '& .MuiTableCell-head': {
              textAlign: 'right',
              padding: '8px 16px',
              borderBottom: '2px solid #e0e0e0',
              whiteSpace: 'nowrap',
              fontWeight: 'bold',
              color: '#1976d2'
            },
            '& .MuiTableCell-root': {
              padding: '4px 8px',
              whiteSpace: 'nowrap',
              borderColor: '#e0e0e0'
            },
            '& .MuiTableRow-root:hover': {
              bgcolor: '#f5f5f5'
            }
          }}
        >
          <TableHead>
            <TableRow>
              <TableCell align="center" sx={{ width: '50px' }}>№</TableCell>
              <TableCell align="center" sx={{ width: '80px' }}>Годы</TableCell>
              <TableCell align="center" sx={{ minWidth: '120px' }}>Нефти</TableCell>
              <TableCell align="center" sx={{ minWidth: '120px' }}>Жидкости</TableCell>
              <TableCell align="center" sx={{ minWidth: '120px' }}>Воды</TableCell>
              <TableCell align="center" sx={{ minWidth: '120px' }}>Обводненность</TableCell>
              <TableCell align="center" sx={{ width: '100px' }}>Active points</TableCell>
              <TableCell align="center" sx={{ width: '50px' }}></TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {data.map((row, index) => (
              <TableRow key={index}>
                <TableCell align="center">
                  {index + 1}
                </TableCell>
                <TableCell>
                  <TextField
                    variant="outlined"
                    size="small"
                    value={row.year || ''}
                    onChange={(e) => handleChange(index, 'year', e.target.value)}
                    sx={{ width: '80px' }}
                  />
                </TableCell>
                <TableCell>
                  <TextField
                    variant="outlined"
                    size="small"
                    type="number"
                    value={row.oil || ''}
                    onChange={(e) => handleChange(index, 'oil', e.target.value)}
                    sx={{ width: '120px' }}
                  />
                </TableCell>
                <TableCell>
                  <TextField
                    variant="outlined"
                    size="small"
                    type="number"
                    value={row.liquid || ''}
                    onChange={(e) => handleChange(index, 'liquid', e.target.value)}
                    sx={{ width: '120px' }}
                  />
                </TableCell>
                <TableCell align="center">
                  {row.water || '0'}
                </TableCell>
                <TableCell align="center">
                  {row.waterCut || '0'}
                </TableCell>
                <TableCell align="center">
                  <Checkbox
                    checked={row.active || false}
                    disabled
                  />
                </TableCell>
                <TableCell align="center">
                  <IconButton
                    size="small"
                    onClick={() => deleteRow(index)}
                    sx={{ color: '#d32f2f' }}
                  >
                    <DeleteIcon fontSize="small" />
                  </IconButton>
                </TableCell>
              </TableRow>
            ))}
            <TableRow>
              <TableCell colSpan={8} sx={{ border: 'none' }}>
                <TextField
                  fullWidth
                  size="small"
                  placeholder="Нажмите Enter для добавления строки"
                  onKeyPress={(e) => {
                    if (e.key === 'Enter') {
                      e.preventDefault();
                      addRow();
                    }
                  }}
                  variant="outlined"
                  sx={{
                    '& .MuiOutlinedInput-root': {
                      bgcolor: '#f8f9fa',
                      '& fieldset': {
                        borderColor: '#e0e0e0',
                      },
                      '&:hover fieldset': {
                        borderColor: '#1976d2',
                      }
                    }
                  }}
                />
              </TableCell>
            </TableRow>
          </TableBody>
        </Table>
      </TableContainer>
      <Box sx={{ mt: 2, display: 'flex', justifyContent: 'flex-end' }}>
        <IconButton
          onClick={addRow}
          sx={{
            bgcolor: '#1976d2',
            color: 'white',
            '&:hover': {
              bgcolor: '#1565c0'
            }
          }}
        >
          <AddIcon />
        </IconButton>
      </Box>
    </Box>
  );
};

export default DataGrid;
