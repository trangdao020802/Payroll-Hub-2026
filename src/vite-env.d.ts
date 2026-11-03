/// <reference types="vite/client" />
declare module '*.png' {
  const value: string;
  export default value;
}
declare module "d3" {
  export function csvParseRows<T>(text: string, convert: (row: string[]) => T | null): T[];
}
