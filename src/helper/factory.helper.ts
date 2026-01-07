const FACTORY_LABEL_MAP: Record<string, string> = {
  LYV: '樂億 - LYV',
  LHG: '樂億II - LHG',
  LVL: '億春B - LVL',
  LYM: '昌億 - LYM',
  LYF: '億福 - LYF',
  JAZ: 'Jiazhi-1',
  JZS: 'Jiazhi-2',
};

export type FactoryCode = keyof typeof FACTORY_LABEL_MAP;

export const getFactory = (factory: FactoryCode): string => {
  return FACTORY_LABEL_MAP[factory] ?? '';
};

export const FACTORY_LIST: FactoryCode[] = [
  'LYV',
  'LHG',
  'LVL',
  'LYM',
  'LYF',
  'JAZ',
  'JZS',
];
