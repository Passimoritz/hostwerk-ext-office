/*
 * Spreadsheet error values (#DIV/0!, #N/A …). One shared instance per code, so `v === ERR.NA` works,
 * but always test with isErr(v) – values from files may be fresh instances.
 */
export class FErr {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}

export const ERR = {
  NULL: new FErr('#NULL!'),
  DIV0: new FErr('#DIV/0!'),
  VALUE: new FErr('#VALUE!'),
  REF: new FErr('#REF!'),
  NAME: new FErr('#NAME?'),
  NUM: new FErr('#NUM!'),
  NA: new FErr('#N/A'),
  SPILL: new FErr('#SPILL!'),
  CALC: new FErr('#CALC!'),
  CIRC: new FErr('#CIRC!'), // circular reference (not an Excel code – saved as 0 like Excel shows it)
  GETTING_DATA: new FErr('#GETTING_DATA'),
};

const BY_CODE = new Map(Object.values(ERR).map((e) => [e.code, e]));
export const ERROR_CODES = [...BY_CODE.keys()];

export const isErr = (v) => v instanceof FErr;
/** '#N/A' → ERR.NA (unknown codes → a new FErr) */
export const errOf = (code) => BY_CODE.get(String(code).toUpperCase()) || new FErr(String(code));
/** Excel's ERROR.TYPE numbers */
export const ERROR_TYPE = { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#GETTING_DATA': 8, '#SPILL!': 9, '#CALC!': 14 };
