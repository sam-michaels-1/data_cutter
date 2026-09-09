import ExcelJS from 'exceljs';
import { readFileSync, writeFileSync } from 'fs';
import { detectTableColumns } from './src/engine/detect';
import { buildEngineConfig } from './src/engine/config_builder';
import { generateDataPack } from './src/engine/generator';

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(readFileSync('/home/ubuntu/repos/data-cutter/frontend/public/sample-data.xlsx'));

  const sheetName = wb.worksheets[0].name;
  const detect = detectTableColumns(wb, sheetName);

  const config = buildEngineConfig({
    wb,
    sheetName,
    dataType: 'arr',
    dateCol: detect.date_columns[0],
    customerIdCol: detect.customer_id_col,
    arrCol: detect.date_columns[0],
    attributes: detect.attribute_cols.map((a: { header: string; letter: string }) => ({ display_name: a.header, letter: a.letter })),
    outputGranularities: ['annual', 'quarterly'],
    fiscalYearEndMonth: 12,
    rowCount: detect.row_count,
    scaleFactor: detect.auto_scale_factor,
    dataFrequency: detect.detected_frequency,
    inputFormat: 'cleaned',
    dateColumns: detect.date_columns,
    headerRow: detect.header_row,
    dateHeaderRow: detect.date_header_row,
  });

  const out = await generateDataPack(config, wb, (msg: string) => console.log(msg));
  const buf = await out.xlsx.writeBuffer();
  writeFileSync('/home/ubuntu/summary_test_fixed.xlsx', Buffer.from(buf));
  console.log('written /home/ubuntu/summary_test_fixed.xlsx');
}

main().catch(console.error);
