export const parse: <T>(
  s: 'xlsb' | 'xlsm' | 'xlsx' | 'xls',
  b: Buffer,
  headers: [string, keyof T][],
  sheet_name?: string | number,
) => T[];
