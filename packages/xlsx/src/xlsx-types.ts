export type SheetCellValue = string | number | boolean | Date | null | undefined;

export type ReadSheet<ParsedNumber = number> = {
  sheet: string;
  data: Array<Array<string | ParsedNumber | boolean | Date | null>>;
};

export type SchemaValueType<Value> = StringConstructor | NumberConstructor | BooleanConstructor | DateConstructor | ((value: SheetCellValue) => Value | undefined);

type SchemaEntryForValue<Key extends keyof Object, Object extends object, TopLevel extends object, ColumnTitle extends string> = {
  column: ColumnTitle;
  type?: SchemaValueType<Object[Key]>;
  oneOf?: Object[Key][];
  required?: boolean | ((row: TopLevel) => boolean);
  validate?(value: Object[Key]): void;
};

type SchemaEntryRecursive<Key extends keyof Object, Object extends object, TopLevel extends object, ColumnTitle extends string> = {
  schema: Object[Key] extends object ? Schema<Object[Key], ColumnTitle> : never;
  required?: boolean | ((row: TopLevel) => boolean);
};

export type Schema<Object extends object, ColumnTitle extends string = string> = {
  [Key in keyof Object]: SchemaEntryForValue<Key, Object, Object, ColumnTitle> | SchemaEntryRecursive<Key, Object, Object, ColumnTitle>;
};

export interface ParseSheetDataError<ColumnTitle extends string = string> {
  row: number;
  column: ColumnTitle;
  columnIndex: number;
  error: string;
  reason: string | undefined;
  value: SheetCellValue;
  type?: SchemaValueType<unknown>;
}

export type ParseSheetDataResult<Object extends object, ColumnTitle extends string = string, Error extends ParseSheetDataError<ColumnTitle> = ParseSheetDataError<ColumnTitle>> =
  | { objects: Object[]; errors: undefined }
  | { objects: undefined; errors: Error[] };

export interface WriterCell {
  value?: string | number | boolean | Date;
  type?: StringConstructor | NumberConstructor | BooleanConstructor | DateConstructor | "Formula";
  format?: string;
  align?: "left" | "center" | "right";
  alignVertical?: "top" | "center" | "bottom";
  backgroundColor?: string;
  textColor?: string;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: "bold";
  fontStyle?: "italic";
  textDecoration?: { underline?: true; doubleUnderline?: true; strikethrough?: boolean };
  wrap?: boolean;
  indent?: number;
  height?: number;
  columnSpan?: number;
  rowSpan?: number;
  leftBorderColor?: string;
  leftBorderStyle?: string;
  rightBorderColor?: string;
  rightBorderStyle?: string;
  topBorderColor?: string;
  topBorderStyle?: string;
  bottomBorderColor?: string;
  bottomBorderStyle?: string;
}

export type WriterCellValue = WriterCell | string | number | boolean | Date | null | undefined;

export interface WriterColumn<Object> {
  width?: number;
  header?: WriterCellValue;
  cell: (object: Object, objectIndex: number) => WriterCellValue;
}
