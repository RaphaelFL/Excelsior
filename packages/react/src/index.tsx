import { useEffect, useRef, type CSSProperties } from "react";
import { createSpreadsheet, type CreateSpreadsheetOptions } from "@excelsior/vanilla";

export interface SpreadsheetProps extends CreateSpreadsheetOptions {
  className?: string;
  style?: CSSProperties;
}

export const Spreadsheet = ({ className, style, ...options }: SpreadsheetProps) => {
  const hostRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!hostRef.current) {
      return undefined;
    }

    const instance = createSpreadsheet(hostRef.current, options);
    return () => instance.destroy();
  }, [options]);

  return <div ref={hostRef} className={className} style={style} />;
};