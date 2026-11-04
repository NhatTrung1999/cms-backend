import dayjs from 'dayjs';

/** The day after `dateTo` (YYYY-MM-DD): the exclusive end of an inclusive date range. */
export const getDayAfter = (dateTo: string) =>
  dayjs(dateTo).add(1, 'day').startOf('day').format('YYYY-MM-DD');

/**
 * Date filter of the HRIS / CMS queries: `dateToExclusive` is the day after `dateTo` ('' without `dateTo`),
 * `hasDate` says whether both dates were given.
 */
export const buildReplacements = (
  dateFrom: string,
  dateTo: string,
): {
  dateToExclusive: string;
  hasDate: boolean;
} => {
  return {
    dateToExclusive: dateTo
      ? dayjs(dateTo).add(1, 'day').format('YYYY-MM-DD')
      : '',
    hasDate: !!(dateFrom && dateTo),
  };
};
