/** WHERE clause (+ positional parameters) shared by every log query. */
export const buildLogFilter = (
  dateFrom: string,
  dateTo: string,
  factory: string,
) => {
  let where: string = 'WHERE 1=1';
  const replacements: any[] = [];

  if (dateFrom && dateTo) {
    where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
    replacements.push(dateFrom, dateTo);
  }

  if (factory) {
    where += ' AND Fac LIKE ?';
    replacements.push(`%${factory}%`);
  }

  return { where, replacements };
};
