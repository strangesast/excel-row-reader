export type Fmt = 'xlsb' | 'xlsm' | 'xlsx' | 'xls';
export const parse: <T>(s: Fmt, b: Buffer, headers: [string, keyof T][], sheet_name?: string | number) => T[];

export const dump: <T>(s: Fmt, b: Buffer, headers: [string, keyof T][], sheet_name?: string | number) => Buffer;
