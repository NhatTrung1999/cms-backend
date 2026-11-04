/** A plain column name (letters, digits, `_`), so a sort field from the request cannot carry SQL. */
const COLUMN_NAME = /^[A-Za-z_][A-Za-z0-9_]*$/;

/**
 * The `ORDER BY` term of a sort field and order sent by the client: `<sortField> ASC|DESC` (ASC only for 'asc').
 * Throws when the sort field is not a plain column name (such a value always made the query fail before).
 */
export const orderByClause = (sortField: string, sortOrder: string) => {
  if (!COLUMN_NAME.test(String(sortField))) {
    throw new Error(`Invalid sort field: ${sortField}`);
  }
  return `${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}`;
};
