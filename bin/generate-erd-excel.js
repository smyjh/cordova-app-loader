const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const ERD_PATH = path.resolve('/workspace/docs/erd.mmd');
const OUT_XLSX = path.resolve('/workspace/docs/erd.xlsx');

function readFileText(filePath) {
  return fs.readFileSync(filePath, 'utf8');
}

function mapAliasToFullTableName(alias, allTableNames) {
  // Try exact match first
  if (allTableNames.includes(alias)) return alias;
  // Try endsWith matching like COMP -> G5_SCH_INFO_STAT_COMP
  const endsWithMatch = allTableNames.find((t) => t.endsWith(`_${alias}`) || t.endsWith(alias));
  return endsWithMatch || alias;
}

function parseMermaidERD(text) {
  const lines = text.split(/\r?\n/);

  const relationships = []; // {fromTable, toTable, label, raw}
  const tables = {}; // name -> { name, columns: [{name,type,pk,fk,fkTargetTable,fkTargetColumn,unique,raw}] }

  // First pass: collect table names and relationship lines
  const relationshipLineRegex = /^\s*([A-Z0-9_]+)\s+[^ ]+\s+([A-Z0-9_]+)\s*:\s*(.+)$/;
  const tableHeaderRegex = /^\s*([A-Z0-9_]+)\s*\{\s*$/;

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const relMatch = line.match(relationshipLineRegex);
    if (relMatch) {
      relationships.push({ fromTable: relMatch[1], toTable: relMatch[2], label: relMatch[3].trim(), raw: line });
      i += 1;
      continue;
    }
    const tableMatch = line.match(tableHeaderRegex);
    if (tableMatch) {
      const tableName = tableMatch[1];
      if (!tables[tableName]) tables[tableName] = { name: tableName, columns: [] };
      i += 1;
      // read until closing brace
      while (i < lines.length && !/^\s*}\s*$/.test(lines[i])) {
        const colLine = lines[i].trim();
        if (colLine.length > 0) {
          const col = parseColumnLine(colLine);
          if (col) tables[tableName].columns.push(col);
        }
        i += 1;
      }
      // consume the closing brace
      i += 1;
      continue;
    }
    i += 1;
  }

  // Resolve FK alias targets using full table names
  const allTableNames = Object.keys(tables);
  for (const t of allTableNames) {
    for (const col of tables[t].columns) {
      if (col.fk && col.fkTargetTable) {
        col.fkTargetTable = mapAliasToFullTableName(col.fkTargetTable, allTableNames);
      }
    }
  }

  // Also convert relationship short names to actual table names
  const normalizedRelationships = relationships.map((r) => ({
    fromTable: mapAliasToFullTableName(r.fromTable, allTableNames),
    toTable: mapAliasToFullTableName(r.toTable, allTableNames),
    label: r.label,
    raw: r.raw,
  }));

  return { tables, relationships: normalizedRelationships };
}

function parseColumnLine(colLine) {
  // Expected patterns like:
  //   int wr_id PK
  //   int wr_no FK "-> COMP.wr_id"
  //   varchar wr_subject
  //   longtext wr_content
  //   enum wr_join_type
  //   int wr_parent FK "self"  (we treat self FK without target column)
  const m = colLine.match(/^(\w+)\s+(\w+)(?:\s+(.*))?$/);
  if (!m) return null;
  const dataType = m[1];
  const name = m[2];
  const rest = m[3] ? m[3].trim() : '';
  let pk = false;
  let unique = false;
  let fk = false;
  let fkTargetTable = '';
  let fkTargetColumn = '';

  if (/(?:^|\s)PK(?:\s|$)/.test(rest)) pk = true;
  if (/(?:^|\s)UNIQUE(?:\s|$)/.test(rest)) unique = true;
  if (/(?:^|\s)FK(?:\s|$)/.test(rest)) fk = true;

  const arrowMatch = rest.match(/->\s*([A-Z0-9_]+)(?:\.(\w+))?/i);
  if (arrowMatch) {
    fkTargetTable = arrowMatch[1];
    fkTargetColumn = arrowMatch[2] || '';
  } else if (/\bself\b/i.test(rest)) {
    fkTargetTable = 'self';
  }

  return { name, type: dataType, pk, unique, fk, fkTargetTable, fkTargetColumn, raw: colLine };
}

async function writeExcel({ tables, relationships }) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'ERD Generator';
  wb.created = new Date();

  // Sheet 1: Tables
  const wsTables = wb.addWorksheet('Tables');
  wsTables.columns = [
    { header: 'Table', key: 'table', width: 34 },
    { header: 'Column', key: 'column', width: 28 },
    { header: 'Type', key: 'type', width: 16 },
    { header: 'PK', key: 'pk', width: 6 },
    { header: 'FK', key: 'fk', width: 6 },
    { header: 'FK Target Table', key: 'fkTable', width: 34 },
    { header: 'FK Target Column', key: 'fkColumn', width: 24 },
    { header: 'Unique', key: 'unique', width: 8 },
  ];
  Object.values(tables)
    .sort((a, b) => a.name.localeCompare(b.name))
    .forEach((t) => {
      t.columns.forEach((c) => {
        wsTables.addRow({
          table: t.name,
          column: c.name,
          type: c.type,
          pk: c.pk ? 'Y' : '',
          fk: c.fk ? 'Y' : '',
          fkTable: c.fk ? c.fkTargetTable : '',
          fkColumn: c.fk ? c.fkTargetColumn : '',
          unique: c.unique ? 'Y' : '',
        });
      });
    });

  // Sheet 2: Relationships
  const wsRels = wb.addWorksheet('Relationships');
  wsRels.columns = [
    { header: 'From Table', key: 'from', width: 34 },
    { header: 'To Table', key: 'to', width: 34 },
    { header: 'Label', key: 'label', width: 24 },
    { header: 'Source', key: 'source', width: 12 },
  ];

  // From diagram edges
  const seen = new Set();
  relationships.forEach((r) => {
    const sig = `edge|${r.fromTable}|${r.toTable}|${r.label}`;
    if (seen.has(sig)) return;
    seen.add(sig);
    wsRels.addRow({ from: r.fromTable, to: r.toTable, label: r.label, source: 'edge' });
  });

  // From column FK annotations
  Object.values(tables).forEach((t) => {
    t.columns.forEach((c) => {
      if (c.fk && c.fkTargetTable) {
        const toTable = c.fkTargetTable === 'self' ? t.name : c.fkTargetTable;
        const label = c.fkTargetColumn ? `${c.name} -> ${toTable}.${c.fkTargetColumn}` : `${c.name} -> ${toTable}`;
        const sig = `col|${t.name}|${toTable}|${label}`;
        if (seen.has(sig)) return;
        seen.add(sig);
        wsRels.addRow({ from: t.name, to: toTable, label, source: 'column' });
      }
    });
  });

  // Basic header styling
  [wsTables, wsRels].forEach((ws) => {
    ws.getRow(1).font = { bold: true };
  });

  await wb.xlsx.writeFile(OUT_XLSX);
}

async function main() {
  const text = readFileText(ERD_PATH);
  const parsed = parseMermaidERD(text);
  await writeExcel(parsed);
  console.log(`Wrote Excel: ${OUT_XLSX}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});