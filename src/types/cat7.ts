export const ACTIVITY_TYPES = ['3.3'] as const;
export type ActivityType = (typeof ACTIVITY_TYPES)[number];
