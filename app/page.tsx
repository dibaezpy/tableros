import fs from "node:fs";
import path from "node:path";
import * as XLSX from "xlsx";

const excelPath = path.join(process.cwd(), "base", "costo historico.xlsx");
const sheetName = "costo";
const maxRows = 200;
const maxColumns = 30;

function readExcel() {
  const fileName = path.basename(excelPath);

  if (!fs.existsSync(excelPath)) {
    return {
      fileName,
      rows: [] as string[][],
      totalRows: 0,
      availableSheets: [] as string[],
      error: `No se encontro el archivo ${fileName}.`,
    };
  }

  try {
    const fileBuffer = fs.readFileSync(excelPath);
    const workbook = XLSX.read(fileBuffer, { type: "buffer", cellDates: true });
    const availableSheets = workbook.SheetNames;
    const selectedSheet = availableSheets.find(
      (availableSheet) => availableSheet.toLowerCase() === sheetName,
    );

    if (!selectedSheet) {
      return {
        fileName,
        rows: [] as string[][],
        totalRows: 0,
        availableSheets,
        error: `No se encontro la hoja ${sheetName}.`,
      };
    }

    const rows = XLSX.utils.sheet_to_json<string[]>(workbook.Sheets[selectedSheet], {
      header: 1,
      defval: "",
      raw: false,
    });

    return {
      fileName,
      rows: rows
        .slice(0, maxRows)
        .map((row) => row.slice(0, maxColumns).map((cell) => String(cell ?? ""))),
      totalRows: rows.length,
      availableSheets,
      error: null,
    };
  } catch (error) {
    return {
      fileName,
      rows: [] as string[][],
      totalRows: 0,
      availableSheets: [] as string[],
      error: error instanceof Error ? error.message : "No se pudo leer el archivo.",
    };
  }
}

export default function Home() {
  const sheetData = readExcel();
  const header = sheetData.rows[0] ?? [];
  const dataRows = sheetData.rows.slice(1);

  return (
    <main className="shell">
      <style>{`
        :root {
          --background: #f5f7fa;
          --foreground: #1f2937;
          --muted: #667085;
          --line: #d8dee8;
          --panel: #ffffff;
          --primary: #0f766e;
          --warning-bg: #fff7e6;
          --warning-line: #f3c36b;
          --warning-text: #7a4a05;
        }

        * { box-sizing: border-box; }
        html, body { min-height: 100%; margin: 0; }
        body {
          background: var(--background);
          color: var(--foreground);
          font-family: Arial, Helvetica, sans-serif;
        }
        .shell { min-height: 100vh; }
        .topbar {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 24px;
          padding: 18px 28px;
          background: var(--panel);
          border-bottom: 1px solid var(--line);
        }
        .topbar h1 { margin: 0; font-size: 22px; line-height: 1.2; }
        .eyebrow {
          margin: 0 0 4px;
          color: var(--muted);
          font-size: 12px;
          font-weight: 700;
          text-transform: uppercase;
        }
        .file-pill {
          border: 1px solid var(--line);
          background: #f8fafc;
          color: var(--foreground);
          padding: 8px 12px;
          border-radius: 999px;
          font-size: 13px;
          white-space: nowrap;
        }
        .content {
          width: min(1180px, calc(100% - 32px));
          margin: 0 auto;
          padding: 28px 0 44px;
        }
        .tab {
          display: inline-flex;
          align-items: center;
          height: 40px;
          border-bottom: 2px solid var(--primary);
          color: var(--foreground);
          font-weight: 700;
          padding: 0 12px;
          margin-bottom: 18px;
        }
        .panel {
          background: var(--panel);
          border: 1px solid var(--line);
          border-radius: 8px;
          padding: 20px;
        }
        .panel-head { margin-bottom: 16px; }
        .panel-head h2 { margin: 0; font-size: 20px; }
        .panel-head p {
          margin: 6px 0 0;
          color: var(--muted);
          font-size: 13px;
        }
        .notice {
          border: 1px solid var(--warning-line);
          background: var(--warning-bg);
          color: var(--warning-text);
          border-radius: 8px;
          margin-bottom: 16px;
          padding: 14px;
        }
        .notice p { margin: 6px 0 0; }
        .table-wrap {
          overflow: auto;
          max-height: 620px;
          border: 1px solid var(--line);
          border-radius: 8px;
        }
        table {
          width: 100%;
          min-width: 920px;
          border-collapse: collapse;
          font-size: 13px;
        }
        th, td {
          padding: 10px 12px;
          border-bottom: 1px solid var(--line);
          text-align: left;
          white-space: nowrap;
        }
        th {
          position: sticky;
          top: 0;
          background: #f8fafc;
          color: var(--muted);
          font-size: 12px;
          text-transform: uppercase;
          z-index: 1;
        }
        @media (max-width: 900px) {
          .topbar { align-items: stretch; flex-direction: column; }
        }
      `}</style>

      <header className="topbar">
        <div>
          <p className="eyebrow">Archivo Excel</p>
          <h1>Vista de hoja costo</h1>
        </div>
        <div className="file-pill">{sheetData.fileName}</div>
      </header>

      <section className="content">
        <div className="tab">Costo</div>

        <section className="panel">
          <div className="panel-head">
            <div>
              <h2>Hoja costo</h2>
              <p>
                {sheetData.error
                  ? "No se pudo cargar la hoja solicitada."
                  : `${sheetData.totalRows.toLocaleString("es-PY")} filas leidas.`}
              </p>
            </div>
          </div>

          {sheetData.error ? (
            <div className="notice">
              <strong>{sheetData.error}</strong>
              {sheetData.availableSheets.length > 0 ? (
                <p>Hojas disponibles: {sheetData.availableSheets.join(", ")}</p>
              ) : null}
            </div>
          ) : null}

          {sheetData.rows.length > 0 ? (
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    {header.map((cell, index) => (
                      <th key={`${cell}-${index}`}>{cell || `Columna ${index + 1}`}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {dataRows.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {header.map((_, cellIndex) => (
                        <td key={cellIndex}>{row[cellIndex] ?? ""}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : null}
        </section>
      </section>
    </main>
  );
}
