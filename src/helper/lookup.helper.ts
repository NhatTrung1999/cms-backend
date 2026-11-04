/**
 * Wraps `lookup` so it runs once per distinct key: every caller with the same key shares the same (pending) result.
 * Use it where many rows ask the database the same question (one query per row would queue on the pool).
 */
export const memoizeByKey = <V>(lookup: (key: string) => Promise<V>) => {
  const results = new Map<string, Promise<V>>();
  return (key: string): Promise<V> => {
    let result = results.get(key);
    if (!result) {
      result = lookup(key);
      results.set(key, result);
    }
    return result;
  };
};
