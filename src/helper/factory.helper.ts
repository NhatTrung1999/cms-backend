export const getFactory = (factory: string) => {
  switch (factory) {
    case 'LYV':
      return '樂億 - LYV';
    case 'LHG':
      return '樂億II - LHG';
    case 'LVL':
      return '億春B - LVL';
    case 'LYM':
      return '昌億 - LYM';
    case 'LYF':
      return '億福 - LYF';
    case 'JAZ':
      return 'Jiazhi-1';
    case 'JZS':
      return 'Jiazhi-2';
    default:
      return '';
  }
};
