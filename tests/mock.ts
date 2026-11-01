// Libro de Excel simulado para probar los Office Scripts fuera de Excel.
// Implementa solo los métodos usados por los scripts (ver excelscript.d.ts).
type V = string | number | boolean;
interface DatosTabla { hoja: string; headers: string[]; rows: V[][]; formulas: string[][]; }

class MockRange {
  constructor(private getter: () => V[][], private setter?: (r: number, c: number, v: V) => void,
    private formulas?: () => string[][], private r0 = 0, private c0 = 0) {}
  getValues(): V[][] { return this.getter(); }
  getFormulas(): string[][] {
    const f = this.formulas ? this.formulas() : [];
    return this.getter().map((row, i) => row.map((v, j) => (f[i] && f[i][j]) || String(v)));
  }
  getCell(row: number, column: number): MockRange {
    return new MockRange(() => [[this.getter()[row][column]]], (_r, _c, v) => this.setter!(this.r0 + row, this.c0 + column, v),
      () => [[(this.formulas ? this.formulas()[row] || [] : [])[column] || ""]], this.r0 + row, this.c0 + column);
  }
  getRow(row: number): MockRange {
    return new MockRange(() => [this.getter()[row]], this.setter, () => [(this.formulas ? this.formulas()[row] : []) || []], this.r0 + row, this.c0);
  }
  setValue(value: V): void { this.setter!(this.r0, this.c0, value); }
  getFormat() { return { getFont: () => ({ setSize: (_: number) => {}, setBold: (_: boolean) => {} }) }; }
}

class MockTable {
  escrituras = 0;
  constructor(public nombre: string, public d: DatosTabla) {}
  getHeaderRowRange() { return new MockRange(() => [this.d.headers.slice()]); }
  getRowCount() { return this.d.rows.length; }
  getRangeBetweenHeaderAndTotal() {
    return new MockRange(() => this.d.rows.map((r) => r.slice()), (r, c, v) => {
      if (this.d.formulas[r] && this.d.formulas[r][c]) throw new Error(`Se intentó sobrescribir la fórmula ${this.nombre}[${this.d.headers[c]}]`);
      this.d.rows[r][c] = typeof v === "string" && v.startsWith("'") ? v.slice(1) : v;
      this.escrituras++;
    }, () => this.d.formulas);
  }
  addRow(index?: number, values?: V[]) {
    if (!values || values.length !== this.d.headers.length) throw new Error(`addRow ${this.nombre}: ${values?.length} valores, se esperaban ${this.d.headers.length}`);
    const f = values.map((v) => (typeof v === "string" && v.startsWith("=") ? v : ""));
    const r = values.map((v) => (typeof v === "string" && v.startsWith("=") ? "" : typeof v === "string" && v.startsWith("'") ? v.slice(1) : v));
    this.d.rows.push(r);
    this.d.formulas.push(f);
    this.escrituras++;
  }
  deleteRowsAt(index: number, count = 1) { this.d.rows.splice(index, count); this.d.formulas.splice(index, count); }
}

class MockWorkbook {
  tablas = new Map<string, MockTable>();
  celdas = new Map<string, V>();
  constructor(json: { tablas: { [k: string]: DatosTabla }; hojas: string[] }, public hojas: string[] = json.hojas) {
    for (const k of Object.keys(json.tablas)) this.tablas.set(k, new MockTable(k, JSON.parse(JSON.stringify(json.tablas[k]))));
  }
  getTable(k: string) { return this.tablas.get(k); }
  getWorksheet(n: string) {
    if (this.hojas.indexOf(n) < 0) return undefined;
    return {
      getRange: (a: string) => new MockRange(() => [[this.celdas.get(`${n}!${a}`) ?? ""]], (_r, _c, v) => { this.celdas.set(`${n}!${a}`, v); }),
      getCharts: () => [],
    };
  }
  getApplication() { return { calculate: (_: unknown) => {} }; }
  getPivotTables() { return []; }
  refreshAllPivotTables() {}
  filas(t: string): { [col: string]: V }[] {
    const tb = this.tablas.get(t)!;
    return tb.d.rows.map((r) => Object.fromEntries(tb.d.headers.map((h, i) => [h, r[i]])));
  }
  fila(t: string, col: string, valor: V) { return this.filas(t).find((r) => r[col] === valor); }
}

// ------------------------------------------------------------- aserciones ---
let __ok = 0;
let __fallas = 0;
function verificar(cond: boolean, msg: string) {
  if (cond) { __ok++; } else { __fallas++; console.log("  ✗ " + msg); }
}
function igual(a: unknown, b: unknown, msg: string) {
  verificar(JSON.stringify(a) === JSON.stringify(b), `${msg} — esperado ${JSON.stringify(b)}, obtenido ${JSON.stringify(a)}`);
}
function finPruebas(nombre: string) {
  console.log(`${nombre}: ${__ok} verificaciones OK, ${__fallas} fallas`);
  if (__fallas > 0) (globalThis as unknown as { process: { exitCode: number } }).process.exitCode = 1;
}
declare function require(m: string): any;
function cargarLibro(ruta: string): MockWorkbook {
  const fs = require("fs");
  return new MockWorkbook(JSON.parse(fs.readFileSync(ruta, "utf-8")));
}
function wbx(m: MockWorkbook): ExcelScript.Workbook { return m as unknown as ExcelScript.Workbook; }
