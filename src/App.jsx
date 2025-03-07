import React, { useState } from 'react';
import {
  Box,
  Button,
  Paper,
  Typography,
  Container,
  Tabs,
  Tab,
  TextField,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Grid,
} from '@mui/material';
import DataGrid from './components/DataGrid';
import ConsoleOutput from './components/ConsoleOutput';
import ResultsChart from './components/ResultsChart';
import ResultsTable from './components/ResultsTable';
import { calculateAll } from './services/calculationService';
import { importExcel, exportExcel } from './services/excelService';

const methodNames = [
  'nazarov-sipachev',
  'sipachev-posevich',
  'maksimov',
  'sazonov',
  'pirverdyan',
  'kambarov'
];

function App() {
  const [activeTab, setActiveTab] = useState(0);
  const [data, setData] = useState([]);
  const [calculationResults, setCalculationResults] = useState({});
  const [defaultN, setDefaultN] = useState(11);
  const [fnLimit, setFnLimit] = useState(0.15);
  const [feLimit, setFeLimit] = useState(0.85);
  const [error, setError] = useState(null);
  const [showCharts, setShowCharts] = useState(false);
  const [showResults, setShowResults] = useState(false);

  const handleFileImport = async (event) => {
    const file = event.target.files[0];
    if (file) {
      try {
        const importedData = await importExcel(file);
        setData(importedData);
        setCalculationResults({});
        setError(null);
      } catch (error) {
        console.error('Error importing file:', error);
        setError('Error importing file: ' + error.message);
      }
    }
  };

  const handleNChange = (newN) => {
    setDefaultN(newN);
    const sortedData = [...data].sort((a, b) => parseInt(a.year) - parseInt(b.year));
    const updatedData = sortedData.map((row, index) => ({
      ...row,
      active: newN === 0 || index >= sortedData.length - newN
    }));
    setData(updatedData);
  };

  const handleCalculate = () => {
    if (!data.length) {
      setError('No data to calculate');
      return;
    }

    const sortedData = [...data].sort((a, b) => parseInt(a.year) - parseInt(b.year));
    const activeData = defaultN === 0 ? sortedData : sortedData.slice(-defaultN);

    try {
      const allResults = calculateAll(activeData, {
        defaultN,
        fnLimit,
        feLimit
      });
      setCalculationResults(allResults);
      setError(null);
    } catch (err) {
      console.error('Error in calculation:', err);
      setError('Ошибка при вычислении: ' + err.message);
    }
  };

  const handleExport = async () => {
    try {
      await exportExcel(data, calculationResults);
    } catch (error) {
      console.error('Error exporting file:', error);
      setError('Error exporting file: ' + error.message);
    }
  };

  const handleShowCharts = () => {
    if (!calculationResults[methodNames[activeTab]]) {
      setError('Сначала выполните расчет');
      return;
    }
    setShowCharts(true);
  };

  const handleShowResults = () => {
    setShowResults(!showResults);
  };

  const currentMethodResults = calculationResults[methodNames[activeTab]];

  return (
    <Box sx={{ 
      height: '100vh',
      display: 'flex',
      flexDirection: 'column',
      overflow: 'hidden'
    }}>
      {/* Header */}
      <Typography variant="h4" component="h1" sx={{ p: 3, pb: 0, color: '#1976d2', fontWeight: 500 }}>
        Инженерный калькулятор
      </Typography>

      {/* Top Controls */}
      <Paper sx={{ m: 3, mb: 0, p: 2, display: 'flex', gap: 2, alignItems: 'center' }}>
        <Button variant="contained" component="label">
          Импорт
          <input type="file" hidden accept=".xlsx,.xls" onChange={handleFileImport} />
        </Button>

        <FormControl sx={{ minWidth: 120 }}>
          <InputLabel>n по умолчанию</InputLabel>
          <Select
            value={defaultN}
            label="n по умолчанию"
            onChange={(e) => handleNChange(e.target.value)}
            size="small"
          >
            {[...Array(21)].map((_, i) => (
              <MenuItem key={i} value={i}>{i}</MenuItem>
            ))}
          </Select>
        </FormControl>

        <TextField
          label="fn. пред."
          value={fnLimit}
          onChange={(e) => setFnLimit(e.target.value)}
          size="small"
          sx={{ width: 100 }}
        />

        <TextField
          label="fe. пред."
          value={feLimit}
          onChange={(e) => setFeLimit(e.target.value)}
          size="small"
          sx={{ width: 100 }}
        />

        <Box sx={{ flex: 1 }} />

        <Button variant="contained" onClick={handleCalculate}>
          Вычислить
        </Button>
        <Button variant="outlined" onClick={handleShowResults}>
          Результаты
        </Button>
        <Button variant="outlined" onClick={handleShowCharts}>
          Графики
        </Button>
        <Button variant="outlined" onClick={handleExport}>
          Экспорт
        </Button>
      </Paper>

      {/* Main Content */}
      <Box sx={{ flex: 1, m: 3, mt: 2, display: 'flex', gap: 2, minHeight: 0 }}>
        {/* Left side - Data Grid */}
        <Paper sx={{ flex: 1, overflow: 'hidden', borderRadius: 2 }}>
          <Box sx={{ height: '100%', overflow: 'auto' }}>
            <DataGrid data={data} setData={setData} />
          </Box>
        </Paper>

        {/* Right side - Results */}
        <Paper sx={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', borderRadius: 2 }}>
          <Tabs 
            value={activeTab} 
            onChange={(e, v) => setActiveTab(v)}
            variant="scrollable"
            sx={{ borderBottom: 1, borderColor: 'divider' }}
          >
            <Tab label="Назаров-Сипачев" />
            <Tab label="Сипачев-Посевич" />
            <Tab label="Максимов" />
            <Tab label="Сазонов" />
            <Tab label="Пирвердян" />
            <Tab label="Камбаров" />
          </Tabs>

          <Box sx={{ flex: 1, overflow: 'auto', p: 2 }}>
            {currentMethodResults && !showResults && (
              <ConsoleOutput result={currentMethodResults} />
            )}
            {showResults && Object.keys(calculationResults).length > 0 && (
              <ResultsTable results={Object.values(calculationResults)} />
            )}
          </Box>
        </Paper>
      </Box>

      <ResultsChart
        open={showCharts}
        onClose={() => setShowCharts(false)}
        results={Object.values(calculationResults)}
      />

      {error && (
        <Typography color="error" sx={{ mx: 3, p: 2, bgcolor: '#ffebee', borderRadius: 1 }}>
          {error}
        </Typography>
      )}
    </Box>
  );
}

export default App;
