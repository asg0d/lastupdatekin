import { calculateRSquared } from './utils';

/**
 * Calculates Nazarov-Sipachev method results.
 * 
 * @param {Array} data - Input data.
 * @param {Object} options - Options.
 * @returns {Object} Results.
 */
export const calculate = (data, options = {}) => {
  // First pass: calculate x, y values
  const results = data.map(row => {
    const x = parseFloat(row.water) || 0;  // X = Vв (water volume)
    const oil = parseFloat(row.oil) || 0;
    const liquid = parseFloat(row.liquid) || 0;
    const y = oil !== 0 ? liquid / oil : 0;  // Y = Vж/Vн (liquid/oil ratio)
    const xy = x * y;
    const x2 = x * x;
    
    return {
      year: row.year,
      x,
      y,
      xy,
      x2,
      method: 'Назаров-Сипачев',
      xDescription: 'V воды',
      yDescription: 'V жидкости / V нефти'
    };
  });

  // Calculate regression coefficients directly
  const N = results.length;
  const sumX = results.reduce((sum, p) => sum + p.x, 0);
  const sumY = results.reduce((sum, p) => sum + p.y, 0);
  const sumXY = results.reduce((sum, p) => sum + p.xy, 0);
  const sumX2 = results.reduce((sum, p) => sum + p.x2, 0);

  // A = (N * ∑xiyi - ∑xi * ∑yi) / (N * ∑xi² - ∑xi * ∑xi)
  const numeratorA = N * sumXY - sumX * sumY;
  const denominatorA = N * sumX2 - sumX * sumX;
  const A = denominatorA !== 0 ? numeratorA / denominatorA : 0;

  // B = (∑yi - A * ∑xi) / N
  const B = (sumY - A * sumX) / N;

  // Calculate R²
  const yMean = sumY / N;
  const ssTotal = results.reduce((sum, p) => sum + Math.pow(p.y - yMean, 2), 0);
  const ssResidual = results.reduce((sum, p) => {
    const yPredicted = A * p.x + B;
    return sum + Math.pow(p.y - yPredicted, 2);
  }, 0);
  const R2 = 1 - (ssResidual / ssTotal);

  // Calculate extractable and remaining oil reserves
  const extractableOilReserves = A !== 0 ? 1 / A : null;
  const lastYearOil = parseFloat(data[data.length - 1].oil) || 0;
  const remainingOilReserves = extractableOilReserves !== null ? extractableOilReserves - lastYearOil : null;

  // Calculate Vmax = 1/A
  const vMax = A !== 0 ? 1 / A : null;

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
    vMax,
    method: 'Назаров-Сипачев',
    xDescription: 'V воды',
    yDescription: 'V жидкости / V нефти'
  };
};
