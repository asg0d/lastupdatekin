import { Paper, Table, TableBody, TableCell, TableContainer, TableHead, TableRow } from '@mui/material';

export default function ResultsPanel({ results }) {
  if (!results) return null;

  return (
    <TableContainer component={Paper} sx={{ height: '100%', overflow: 'auto' }}>
      <Table size="small">
        <TableHead>
          <TableRow>
            <TableCell align="center" sx={{ width: '50px' }}>№</TableCell>
            <TableCell align="center" sx={{ width: '80px' }}>Годы</TableCell>
            <TableCell align="center" sx={{ width: '120px' }}>X=Ve</TableCell>
            <TableCell align="center" sx={{ width: '120px' }}>Y=Vж/Vн</TableCell>
            <TableCell align="center" sx={{ width: '120px' }}>XY</TableCell>
            <TableCell align="center" sx={{ width: '120px' }}>X^2</TableCell>
          </TableRow>
        </TableHead>
        <TableBody>
          {results.map((row, index) => (
            <TableRow key={index}>
              <TableCell align="center">{index + 1}</TableCell>
              <TableCell align="center">{row.year}</TableCell>
              <TableCell align="right">{row.x?.toFixed(4) || '0.0000'}</TableCell>
              <TableCell align="right">{row.y?.toFixed(4) || '0.0000'}</TableCell>
              <TableCell align="right">{row.xy?.toFixed(4) || '0.0000'}</TableCell>
              <TableCell align="right">{row.x2?.toFixed(4) || '0.0000'}</TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </TableContainer>
  );
}
