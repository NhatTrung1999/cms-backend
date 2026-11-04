import { getOwn } from './object.helper';

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
  return getOwn(FACTORY_LABEL_MAP, factory) ?? '';
};

const FACTORY_CODE_BY_LABEL = new Map(
  Object.entries(FACTORY_LABEL_MAP).map(([code, label]) => [
    label.trim().toLowerCase(),
    code,
  ]),
);

/** Reverse of getFactory: '樂億 - LYV' -> 'LYV'. Returns '' for an unknown label. */
export const getFactoryCodeByLabel = (label: string): string => {
  return FACTORY_CODE_BY_LABEL.get(label.trim().toLowerCase()) ?? '';
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
