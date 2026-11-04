/** Sorts `data` in place by one field ('asc' = ascending, anything else = descending). */
export const sortByField = <T extends object>(
  data: T[],
  sortField: string,
  sortOrder: string,
) => {
  data.sort((a, b) => {
    const aValue = (a as Record<string, string | number>)[sortField];
    const bValue = (b as Record<string, string | number>)[sortField];
    if (sortOrder === 'asc') {
      return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
    } else {
      return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
    }
  });
};
