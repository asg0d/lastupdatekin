import { calculateLinearRegression, calculateRSquared } from './utils';

export const calculate = (data, options = {}) => {
  // First pass: calculate x, y values
  const results = data.map(row => {
    const liquid = parseFloat(row.liquid) || 0;
    const x = liquid > 0 ? Math.pow(liquid, -0.5) : 0;  // X = Vж^(-1/2) (liquid volume to power -1/2)
    const y = parseFloat(row.oil) || 0;  // Y = Vн (oil volume)
    const xy = x * y;
    const x2 = x * x;
    
    return {
      year: row.year,
      x,
      y,
      xy,
      x2,
      method: 'Пирвердян',
      xDescription: 'V воды',
      yDescription: 'V жидкости / V нефти'
    };
  });

  // Calculate sums for regression coefficients
  const sumX = results.reduce((sum, p) => sum + p.x, 0);
  const sumY = results.reduce((sum, p) => sum + p.y, 0);
  const sumXY = results.reduce((sum, p) => sum + p.xy, 0);
  const sumX2 = results.reduce((sum, p) => sum + p.x2, 0);

  // B = (N * ∑xiyi - ∑xi * ∑yi) / (N * ∑xi² - ∑xi * ∑xi)
  const numeratorB = data.length * sumXY - sumX * sumY;
  const denominatorB = data.length * sumX2 - sumX * sumX;
  const B = denominatorB !== 0 ? numeratorB / denominatorB : 0;

  // A = (∑yi - B * ∑xi) / N
  const A = (sumY - B * sumX) / data.length;

  // Calculate R²
  const N = data.length;
  const yMean = sumY / N;
  const ssTotal = results.reduce((sum, p) => sum + Math.pow(p.y - yMean, 2), 0);
  const ssResidual = results.reduce((sum, p) => {
    const yPredicted = A * p.x + B;
    return sum + Math.pow(p.y - yPredicted, 2);
  }, 0);
  const R2 = 1 - (ssResidual / ssTotal);

  // Calculate extractable and remaining oil reserves
  const extractableOilReserves = A;
  const lastYearOil = parseFloat(data[data.length - 1].oil) || 0;
  const remainingOilReserves = extractableOilReserves - lastYearOil;

  return {
    results,
    coefficients: {
      A,
      B,
      R2
    },
    sums: {
      sumX,
      sumY,
      sumXY,
      sumX2,
      sumXSquared: sumX * sumX
    },
    extractableOilReserves,
    remainingOilReserves,
    method: 'Пирвердян',
    xDescription: 'V воды',
    yDescription: 'V жидкости / V нефти'
  };
};
