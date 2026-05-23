export const getCat6AssistedIds = (value: unknown): string[] => {
  if (typeof value !== 'string' || !value.trim()) {
    return [];
  }

  return value
    .split('$')
    .map((item) => item.trim())
    .filter(Boolean);
};

const getAlphabetSuffix = (index: number): string => {
  let value = index;
  let suffix = '';

  do {
    suffix = String.fromCharCode(65 + (value % 26)) + suffix;
    value = Math.floor(value / 26) - 1;
  } while (value >= 0);

  return suffix;
};

export const appendCat6DocumentSuffix = (
  documentNumber: string,
  index: number,
): string => {
  const base = documentNumber?.trim() ?? '';
  if (!base) return '';
  return `${base}-${getAlphabetSuffix(index)}`;
};
