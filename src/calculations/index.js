import { calculate as nazarovSipachev } from './nazarovSipachev';
import { calculate as sipachevPosevich } from './sipachevPosevich';
import { calculate as maksimov } from './maksimov';
import { calculate as sazonov } from './sazonov';
import { calculate as pirverdyan } from './pirverdyan';
import { calculate as kambarov } from './kambarov';

const calculationMethods = {
  'nazarov-sipachev': nazarovSipachev,
  'sipachev-posevich': sipachevPosevich,
  'maksimov': maksimov,
  'sazonov': sazonov,
  'pirverdyan': pirverdyan,
  'kambarov': kambarov
};

export { calculationMethods };
