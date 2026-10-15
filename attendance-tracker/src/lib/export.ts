import type { ReportTable } from './reports';

function csvCell(v: string | number): string {
  const s = String(v);
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

export function tablesToCsv(tables: ReportTable[]): string {
  const lines: string[] = [];
  for (const t of tables) {
    if (lines.length) lines.push('');
    lines.push(csvCell(t.title), csvCell(t.subtitle), '');
    lines.push(t.headers.map(csvCell).join(','));
    for (const r of t.rows) lines.push(r.map(csvCell).join(','));
    if (t.totals.length) {
      lines.push('');
      for (const [k, v] of t.totals) lines.push(`${csvCell(k)},${csvCell(v)}`);
    }
  }
  return lines.join('\r\n');
}

function download(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** CSV with a UTF-8 BOM so Excel opens it with the right encoding. */
export function exportCsv(tables: ReportTable[], filename: string) {
  download(new Blob(['﻿' + tablesToCsv(tables)], { type: 'text/csv;charset=utf-8' }), filename);
}

export async function exportPdf(tables: ReportTable[], filename: string) {
  const [{ jsPDF }, { default: autoTable }] = await Promise.all([import('jspdf'), import('jspdf-autotable')]);
  const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
  tables.forEach((t, i) => {
    if (i > 0) doc.addPage();
    doc.setFontSize(14);
    doc.text(t.title, 40, 40);
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(t.subtitle, 40, 58);
    doc.setTextColor(0);
    autoTable(doc, {
      startY: 72,
      head: [t.headers],
      body: t.rows.map((r) => r.map(String)),
      styles: { fontSize: 8, cellPadding: 3 },
      headStyles: { fillColor: [30, 64, 120] },
      margin: { left: 40, right: 40 },
    });
    if (t.totals.length) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const y = (doc as any).lastAutoTable.finalY + 16;
      autoTable(doc, {
        startY: y,
        head: [['Monthly totals', '']],
        body: t.totals,
        styles: { fontSize: 9 },
        headStyles: { fillColor: [30, 64, 120] },
        tableWidth: 260,
        margin: { left: 40 },
      });
    }
  });
  doc.save(filename);
}
