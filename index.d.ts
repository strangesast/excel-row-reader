export const parse: <T>(
  s: 'xlsb' | 'xlsm' | 'xlsx' | 'xls',
  b: Buffer,
  sheet_name: string,
  headers: [string, keyof T][],
) => T[];
