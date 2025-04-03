import React, { useState } from 'react';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  BarController,
  Title,
  Tooltip,
  Legend,
  LineElement,
  LineController,
  PointElement
} from 'chart.js';
import { Bar } from 'react-chartjs-2';
import { FormGroup, FormControlLabel, Checkbox, Box, Button } from '@mui/material';
import FileDownloadIcon from '@mui/icons-material/FileDownload';
import { exportExcel } from '../services/excelService';

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  BarController,
  LineElement,
  LineController,
  PointElement,
  Title,
  Tooltip,
  Legend
);

const ResultsBarChart = ({ results }) => {
  const [geologicalReserves, setGeologicalReserves] = useState('');
  const [selectedMethods, setSelectedMethods] = useState({
    'Максимов': true,
    'Назаров-Сипачев': true,
    'Сипачев-Посевич': true,
    'Сазонов': true,
    'Пирвердян': true,
    'Камбаров': true
  });
  
  const handleMethodChange = (method) => {
    setSelectedMethods(prev => ({
      ...prev,
      [method]: !prev[method]
    }));
  };
  
  // Filter results based on selected methods
  const validResults = results
    .filter(result => result.remainingOilReserves !== null)
    .filter(result => selectedMethods[result.method]);

  const averageValue = validResults.length > 0 
    ? validResults.reduce((sum, r) => sum + r.remainingOilReserves, 0) / validResults.length 
    : null;

  const labels = validResults.map(result => result.method);
  
  // Fixed values for ORC calculation
  const cumulativeOilProduction = 46564979.630; // Накопленные добыча нефть
  const remainingOilReserves = 11166666.667;    // V остаточные
  const totalNumerator = cumulativeOilProduction + remainingOilReserves;
  
  const orc = geologicalReserves && geologicalReserves > 0 
    ? (totalNumerator / parseFloat(geologicalReserves)).toFixed(3)
    : null;

  const data = {
    labels,
    datasets: [
      {
        label: 'V остаточные',
        data: validResults.map(result => result.remainingOilReserves),
        backgroundColor: 'rgba(25, 118, 210, 0.6)',
        borderColor: 'rgba(25, 118, 210, 1)',
        borderWidth: 1,
      }
    ]
  };

  if (averageValue !== null) {
    data.datasets.push({
      type: 'line',
      label: 'V среднее',
      data: Array(labels.length).fill(averageValue),
      borderColor: 'rgba(255, 99, 132, 1)',
      borderWidth: 2,
      borderDash: [5, 5],
      pointStyle: false,
      fill: false
    });
  }

  const options = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        display: true,
        position: 'top'
      },
      title: {
        display: true,
        text: 'Сравнение V остаточные по методам',
        font: {
          size: 16
        }
      }
    },
    scales: {
      y: {
        beginAtZero: true,
        title: {
          display: true,
          text: 'V остаточные'
        }
      }
    }
  };

  const handleExport = () => {
    const exportData = {
      chartData: validResults.map(result => ({
        method: result.method,
        remainingOilReserves: result.remainingOilReserves,
        extractableOilReserves: result.extractableOilReserves,
        coefficients: result.coefficients
      })),
      average: {
        remainingOilReserves: averageValue
      },
      orcCalculation: {
        geologicalReserves: parseFloat(geologicalReserves) || 0,
        cumulativeOilProduction,
        remainingOilReserves,
        totalNumerator,
        orc: orc || 0
      }
    };

    exportExcel(exportData);
  };

  return (
    <div>
      <Box sx={{ mb: 2, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <FormGroup row sx={{ flexWrap: 'wrap', gap: 2 }}>
          {Object.keys(selectedMethods).map(method => (
            <FormControlLabel
              key={method}
              control={
                <Checkbox
                  checked={selectedMethods[method]}
                  onChange={() => handleMethodChange(method)}
                />
              }
              label={method}
            />
          ))}
        </FormGroup>
        {/* <Button
          variant="contained"
          startIcon={<FileDownloadIcon />}
          onClick={handleExport}
          sx={{ ml: 2 }}
        >
          Export to Excel
        </Button> */}
      </Box>
      <div style={{ marginBottom: '20px' }}>
        <label style={{ marginRight: '10px' }}>
          Q geological reserves:
          <input
            type="number"
            value={geologicalReserves}
            onChange={(e) => setGeologicalReserves(e.target.value)}
            style={{ marginLeft: '10px' }}
          />
        </label>
        {orc !== null && (
          <div style={{ marginTop: '10px' }}>
            <div>Накопленные добыча нефть: {cumulativeOilProduction.toLocaleString('en-US', { minimumFractionDigits: 3, maximumFractionDigits: 3 })}</div>
            <div>V остаточные: {remainingOilReserves.toLocaleString('en-US', { minimumFractionDigits: 3, maximumFractionDigits: 3 })}</div>
            <div>Total: {totalNumerator.toLocaleString('en-US', { minimumFractionDigits: 3, maximumFractionDigits: 3 })}</div>
            <div>ORC (КИН) = {orc}</div>
          </div>
        )}
      </div>
      <div style={{ height: '300px' }}>
        <Bar data={data} options={options} />
      </div>
    </div>
  );
};

export default ResultsBarChart;
