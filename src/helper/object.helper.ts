/**
 * `map[key]`, but only for keys the map really has: 'constructor', 'toString', '__proto__'...
 * (properties every object inherits) give undefined instead of a function or an object.
 */
export const getOwn = <T>(
  map: Record<string, T>,
  key: string,
): T | undefined => (Object.hasOwn(map, key) ? map[key] : undefined);
