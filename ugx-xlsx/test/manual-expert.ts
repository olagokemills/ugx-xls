import { exportJsonToExcel } from "../src/browser-helper";

const data = [
  { Name: "Alice", Score: 95 },
  { Name: "Bob", Score: 88 },
];

exportJsonToExcel(data, "SampleExport");
