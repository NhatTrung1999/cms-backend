const CAT6_AIRPORT_NAMES: Record<string, string> = {
  NUE: 'Nuremberg Airport (Albrecht Durer Airport Nurnberg)',
  CAN: 'Guangzhou Baiyun International Airport',
  RGN: 'Yangon International Airport',
  SUB: 'Juanda International Airport',
  NKG: 'Nanjing Lukou International Airport',
  HKG: 'Hong Kong International Airport',
  SRG: 'Achmad Yani International Airport',
  SZX: "Shenzhen Bao'an International Airport",
  LAX: 'Los Angeles International Airport',
  TPE: 'Taiwan Taoyuan International Airport',
  SGN: 'Tan Son Nhat International Airport',
  KUL: 'Kuala Lumpur International Airport',
  PVG: 'Shanghai Pudong International Airport',
  RMQ: 'Taichung International Airport',
  BKK: 'Suvarnabhumi Airport',
};

export const getCat6AirportName = (value?: string | null) => {
  const code = value?.trim().toUpperCase() ?? '';
  if (!code) return '';
  return CAT6_AIRPORT_NAMES[code] ?? value?.trim() ?? '';
};
