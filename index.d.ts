export const sync: (input: number) => number;
// sleep [duration] ms, return Promise which resolved 2 * duration
export const sleep: (duration: number) => Promise<number>;

export const world: (b: Buffer, sheet_name: string, headers: string[]) => string[][];
