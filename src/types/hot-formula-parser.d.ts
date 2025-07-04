declare module "hot-formula-parser" {
  export class Parser {
    constructor();
    parse(formula: string): {
      result?: string | number;
      error?: string;
    };
    on(
      event: "callCellValue",
      callback: (
        cellRef: { row: number; column: string },
        done: (value: number | string) => void
      ) => void
    ): void;
    on(
      event: "callRangeValue",
      callback: (
        startCellRef: { row: number; column: string },
        endCellRef: { row: number; column: string },
        done: (values: (number | string)[][]) => void
      ) => void
    ): void;
  }
}
