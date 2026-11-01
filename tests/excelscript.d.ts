// Subconjunto de la API de Office Scripts (ExcelScript) usado por los scripts.
// Las firmas siguen la referencia oficial: https://learn.microsoft.com/javascript/api/office-scripts/excelscript
// Sirve para verificar tipos y ejecutar las pruebas con un libro simulado.
declare namespace ExcelScript {
  const enum CalculationType { recalculate = "Recalculate", full = "Full", fullRebuild = "FullRebuild" }
  const enum AggregationFunction { unknown = "Unknown", automatic = "Automatic", sum = "Sum", count = "Count", average = "Average" }

  interface Workbook {
    getTable(key: string): Table | undefined;
    getWorksheet(key: string): Worksheet | undefined;
    addWorksheet(name?: string): Worksheet;
    getApplication(): Application;
    getPivotTables(): PivotTable[];
    getPivotTable(name: string): PivotTable | undefined;
    refreshAllPivotTables(): void;
    getSlicers(): Slicer[];
    addSlicer(slicerSource: string | PivotTable | Table, sourceField: string | PivotField | number | TableColumn,
      slicerDestination?: string | Worksheet): Slicer;
  }
  interface Application { calculate(calculationType: CalculationType): void; }
  interface Worksheet {
    getRange(address?: string): Range;
    getCharts(): Chart[];
    addPivotTable(name: string, source: Range | string | Table, destination: Range | string): PivotTable;
  }
  interface Table {
    getHeaderRowRange(): Range;
    getRangeBetweenHeaderAndTotal(): Range;
    getRowCount(): number;
    addRow(index?: number, values?: (boolean | string | number)[]): void;
    deleteRowsAt(index: number, count?: number): void;
  }
  interface TableColumn { getName(): string; }
  interface Range {
    getValues(): (string | number | boolean)[][];
    getFormulas(): string[][];
    getCell(row: number, column: number): Range;
    getRow(row: number): Range;
    setValue(value: any): void;
    getFormat(): RangeFormat;
  }
  interface RangeFormat { getFont(): RangeFont; }
  interface RangeFont { setSize(size: number): void; setBold(bold: boolean): void; }
  interface Chart { getTitle(): ChartTitle; getSeries(): ChartSeries[]; }
  interface ChartTitle { getText(): string; }
  interface ChartSeries { setValues(sourceData: Range): void; setXAxisValues(sourceData: Range): void; }
  interface PivotTable {
    getHierarchy(name: string): PivotHierarchy | undefined;
    addRowHierarchy(pivotHierarchy: PivotHierarchy): RowColumnPivotHierarchy;
    addColumnHierarchy(pivotHierarchy: PivotHierarchy): RowColumnPivotHierarchy;
    addDataHierarchy(pivotHierarchy: PivotHierarchy): DataPivotHierarchy;
  }
  interface PivotHierarchy { getName(): string; }
  interface PivotField { getName(): string; }
  interface RowColumnPivotHierarchy { getName(): string; }
  interface DataPivotHierarchy { setSummarizeBy(summarizeBy: AggregationFunction): void; }
  interface Slicer {
    getCaption(): string; setCaption(caption: string): void; setLeft(left: number): void; setTop(top: number): void;
    setWidth(width: number): void; setHeight(height: number): void;
  }
}
declare const console: { log(...args: unknown[]): void };
