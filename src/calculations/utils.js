// Helper functions for calculations
export const calculateLinearRegression = (points) => {
  const N = points.length;
  
  // Calculate sums
  let sumX = 0;
  let sumY = 0;
  let sumXY = 0;
  let sumX2 = 0;

  points.forEach(point => {
    sumX += point.vw;
    sumY += point.vlvo;
    sumXY += point.vw * point.vlvo;
    sumX2 += point.vw * point.vw;
  });

  // Calculate A using the formula:
  // A = (N * ∑xiyi - ∑xi * ∑yi) / (N * ∑xi² - ∑xi * ∑xi)
  const numeratorA = N * sumXY - sumX * sumY;
  const denominatorA = N * sumX2 - sumX * sumX;
  const A = denominatorA !== 0 ? numeratorA / denominatorA : 0;

  // Calculate B using the formula:
  // B = (∑yi - A * ∑xi) / N
  const B = (sumY - A * sumX) / N;

  return { A, B };
};

export const calculateRSquared = (points, A, B) => {
  const yMean = points.reduce((sum, p) => sum + p.vlvo, 0) / points.length;
  
  const ssTotal = points.reduce((sum, p) => {
    return sum + Math.pow(p.vlvo - yMean, 2);
  }, 0);
  
  const ssResidual = points.reduce((sum, p) => {
    const yPredicted = A * p.vw + B;
    return sum + Math.pow(p.vlvo - yPredicted, 2);
  }, 0);
  
  return 1 - (ssResidual / ssTotal);
};
