import { calculate as nazarovSipachev } from '../calculations/nazarovSipachev';
import { calculate as sipachevPosevich } from '../calculations/sipachevPosevich';
import { calculate as maksimov } from '../calculations/maksimov';
import { calculate as sazonov } from '../calculations/sazonov';
import { calculate as pirverdyan } from '../calculations/pirverdyan';
import { calculate as kambarov } from '../calculations/kambarov';

// Calculation types for UI
export const calculationTypes = {
  NAZAROV_SIPACHEV: 'nazarov-sipachev',
  SIPACHEV_POSEVICH: 'sipachev-posevich',
  MAKSIMOV: 'maksimov',
  SAZONOV: 'sazonov',
  PIRVERDYAN: 'pirverdyan',
  KAMBAROV: 'kambarov'
};

// Method names used for calculation results
export const METHOD_NAMES = {
  'nazarov-sipachev': 'Назаров-Сипачев',
  'sipachev-posevich': 'Сипачев-Посевич',
  'maksimov': 'Максимов',
  'sazonov': 'Сазонов',
  'pirverdyan': 'Пирвердян',
  'kambarov': 'Камбаров'
};

const validateData = (data) => {
  if (!Array.isArray(data)) {
    throw new Error('Data must be an array');
  }
  if (data.length === 0) {
    throw new Error('Data array is empty');
  }
  return data.map(row => ({
    year: row.year || '',
    water: parseFloat(row.water) || 0,
    oil: parseFloat(row.oil) || 0,
    liquid: parseFloat(row.liquid) || 0
  }));
};

export const calculateAll = (data, options = {}) => {
  try {
    const validatedData = validateData(data);
    
    const results = {
      'nazarov-sipachev': nazarovSipachev(validatedData, options),
      'sipachev-posevich': sipachevPosevich(validatedData, options),
      'maksimov': maksimov(validatedData, options),
      'sazonov': sazonov(validatedData, options),
      'pirverdyan': pirverdyan(validatedData, options),
      'kambarov': kambarov(validatedData, options)
    };

    // Add method names to results
    Object.keys(results).forEach(key => {
      if (results[key]) {
        results[key].method = METHOD_NAMES[key];
      }
    });

    return results;
  } catch (error) {
    console.error('Error in calculateAll:', error);
    throw error;
  }
};

export const calculate = (data, method, options = {}) => {
  try {
    const validatedData = validateData(data);
    
    let result;
    switch (method) {
      case 'nazarov-sipachev':
        result = nazarovSipachev(validatedData, options);
        break;
      case 'sipachev-posevich':
        result = sipachevPosevich(validatedData, options);
        break;
      case 'maksimov':
        result = maksimov(validatedData, options);
        break;
      case 'sazonov':
        result = sazonov(validatedData, options);
        break;
      case 'pirverdyan':
        result = pirverdyan(validatedData, options);
        break;
      case 'kambarov':
        result = kambarov(validatedData, options);
        break;
      default:
        throw new Error(`Unknown calculation method: ${method}`);
    }

    if (result) {
      result.method = METHOD_NAMES[method];
    }

    return result;
  } catch (error) {
    console.error(`Error calculating ${method}:`, error);
    throw error;
  }
};
