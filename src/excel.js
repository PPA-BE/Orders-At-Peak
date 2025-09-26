// src/excel.js
import { money, parseNum } from "./state.js"; // TAX_RATE now lives in state.js as well

const TEMPLATE_PATH = "po-template-new.xlsx";
const EXCEL_CELLS = {
  sheetName: "Req",
  // Header fields
  requisitioner: "D3",     // REQUISITIONER value (C/D/E merged)
  date: "I3",
  department: "E4",
  supplier: "C6",
  epicorSupplierNo: "D7",
  // Address block (supplier address)
  addr_street: "F8",
  addr_city: "F10",
  addr_province: "F11",
  addr_country: "F12",
  addr_postal: "F13",
  // Contact block (optional - not specified here but kept for future)
  contact_name: "B14",
  contact_email: "B15",
  contact_phone: "B16",
  // Currency (unchanged)
  currencyCad: "C18",
  currencyOther: "E18",
  // Totals
  subTotalCell: "B36"
};

const EXCEL_TABLE = {
  startRow: 20,
  columns: [
    { key: "partNumber",   col: "A" }, // supplierItem or peakPart
    { key: "description",  col: "C" }, // assumption: Description column is C
    { key: "qty",          col: "F" },
    { key: "unitPrice",    col: "G" },
    { key: "uom",          col: "H" },
    { key: "total",        col: "I" },
  ],
  grandTotalCell: null,
};

function a1(col, row){ return `${col}${row}`; }

export async function exportExcelUsingTemplate(payload, items) {
  const resp = await fetch(TEMPLATE_PATH, { cache: "no-store" });
  if (!resp.ok) throw new Error("Template not found: " + TEMPLATE_PATH);

  const ab = await resp.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(ab);
  const ws = EXCEL_CELLS.sheetName ? wb.getWorksheet(EXCEL_CELLS.sheetName) : wb.worksheets[0];

  // Header mapping
  if (EXCEL_CELLS.requisitioner) ws.getCell(EXCEL_CELLS.requisitioner).value = payload.createdBy || payload.user?.name || "";
  if (EXCEL_CELLS.department)    ws.getCell(EXCEL_CELLS.department).value    = payload.department || payload.user?.department || "";
  if (EXCEL_CELLS.date)          { const d = payload.date ? new Date(payload.date) : new Date(); ws.getCell(EXCEL_CELLS.date).value = d; ws.getCell(EXCEL_CELLS.date).numFmt = "yyyy-mm-dd"; }
  if (EXCEL_CELLS.supplier)      ws.getCell(EXCEL_CELLS.supplier).value      = (payload.vendor?.name || payload.vendor?.id || "");
  if (EXCEL_CELLS.epicorSupplierNo) ws.getCell(EXCEL_CELLS.epicorSupplierNo).value = payload.vendor?.referenceNo || "";

  // Address mapping (best effort based on available fields)
  if (EXCEL_CELLS.addr_street)   ws.getCell(EXCEL_CELLS.addr_street).value   = payload.vendor?.address1 || "";
  if (EXCEL_CELLS.addr_city)     ws.getCell(EXCEL_CELLS.addr_city).value     = payload.vendor?.city || "";
  if (EXCEL_CELLS.addr_province) ws.getCell(EXCEL_CELLS.addr_province).value = payload.vendor?.state || "";
  if (EXCEL_CELLS.addr_country)  ws.getCell(EXCEL_CELLS.addr_country).value  = payload.vendor?.country || "";
  if (EXCEL_CELLS.addr_postal)   ws.getCell(EXCEL_CELLS.addr_postal).value   = payload.vendor?.zip || "";

  // Currency mapping
  if (payload.currency && payload.currency.toUpperCase() === "CAD") {
    if (EXCEL_CELLS.currencyCad)   ws.getCell(EXCEL_CELLS.currencyCad).value = "CAD";
    if (EXCEL_CELLS.currencyOther) ws.getCell(EXCEL_CELLS.currencyOther).value = "";
  } else {
    if (EXCEL_CELLS.currencyCad)   ws.getCell(EXCEL_CELLS.currencyCad).value = "";
    if (EXCEL_CELLS.currencyOther) ws.getCell(EXCEL_CELLS.currencyOther).value = (payload.currency || "");
  }



  let row = EXCEL_TABLE.startRow;
  (items || []).forEach((r, i) => {
    EXCEL_TABLE.columns.forEach((c) => {
      const cell = ws.getCell(a1(c.col, row));
      let val = r[c.key];

      // Computed/compat fields
      if (c.key === "line") val = i + 1;
      if (c.key === "partNumber") val = r.supplierItem || r.peakPart || r.partNumber || "";
      if (c.key === "total") {
        // Write a formula, but also backfill a numeric value for safety
        try { cell.value = { formula: `F${row}*G${row}` }; } catch(_) {}
        val = (Number(r.qty || 0) * Number(r.unitPrice || 0)) || 0;
      }
      const isNumeric = c.key === "qty" || c.key === "unitPrice" || c.key === "total";
      if (cell.value && typeof cell.value === 'object' && cell.value.formula) {
        // Already set as a formula; still set a number format if needed
        if (c.key === "unitPrice" || c.key === "total") cell.numFmt = '#,##0.00';
      } else {
        cell.value = isNumeric ? Number(val || 0) : (val ?? "");
      }
    });
    row++;
  });

  
  // Subtotal at B36 = SUM of totals column I from startRow to last row used
  try {
    const firstRow = EXCEL_TABLE.startRow;
    const lastRow = row - 1 >= firstRow ? row - 1 : firstRow;
    const subCell = ws.getCell(EXCEL_CELLS.subTotalCell || "B36");
    subCell.value = { formula: `SUM(I${firstRow}:I${lastRow})` };
  } catch(_) {}
const filename = (payload.poId || "PO") + ".xlsx";
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
}

// ---- HTML preview (includes UOM) ----
function escapeHtml(s=""){ return String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }

export function buildHtmlPreview(po = {}) {
  const items = Array.isArray(po.items) ? po.items : [];

  const subtotal = items.reduce((acc, r) => acc + (parseNum(r?.qty) * parseNum(r?.unitPrice)), 0);
  const tax = +(subtotal * 0.13).toFixed(2);
  const grand = subtotal + tax;

  const rowsHtml = items.map((r, i) => `
    <tr>
      <td class="py-1 px-2 border border-slate-200 text-right">${i + 1}</td>
      <td class="py-1 px-2 border border-slate-200">${escapeHtml(r?.supplierItem || "")}</td>
      <td class="py-1 px-2 border border-slate-200">${escapeHtml(r?.peakPart || "")}</td>
      <td class="py-1 px-2 border border-slate-200">${escapeHtml(r?.description || "")}</td>
      <td class="py-1 px-2 border border-slate-200 text-right">${parseNum(r?.qty) || 0}</td>
      <td class="py-1 px-2 border border-slate-200 text-center w-16">${escapeHtml(r?.uom || "")}</td>
      <td class="py-1 px-2 border border-slate-200 text-right">${money(parseNum(r?.unitPrice))}</td>
      <td class="py-1 px-2 border border-slate-200 text-right font-medium">${money(parseNum(r?.qty) * parseNum(r?.unitPrice))}</td>
    </tr>
  `).join("");

  return `
    <div class="text-sm text-slate-700 space-y-3">
      <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
        <div class="border rounded p-3">
          <div class="text-xs font-semibold mb-1">Vendor</div>
          <div class="break-anywhere">${escapeHtml(po?.vendor?.name || "")}</div>
          <div class="text-slate-500 break-anywhere">${escapeHtml(po?.vendor?.address1 || "")}</div>
          <div class="text-slate-500 break-anywhere">
            ${escapeHtml(po?.vendor?.city || "")}
            ${po?.vendor?.state ? ", " + escapeHtml(po.vendor.state) : ""}
            ${po?.vendor?.zip ? " " + escapeHtml(po.vendor.zip) : ""}
          </div>
        </div>
        <div class="border rounded p-3">
          <div class="text-xs font-semibold mb-1">PO</div>
          <div>PO ID: <span class="font-medium break-anywhere">${escapeHtml(po?.poId || "")}</span></div>
          <div>Date: ${escapeHtml(po?.date || "")}</div>
          <div>Total: <span class="font-medium">${money(grand)}</span></div>
        </div>
        <div class="border rounded p-3">
          <div class="text-xs font-semibold mb-1">Ship To</div>
          <div>Peak Processing Solutions</div>
          <div>2065 Solar Crescent</div>
          <div>Oldcastle, ON, Canada</div>
          <div>N0R1L0</div>
        </div>
      </div>

      <div class="overflow-x-auto">
        <table class="min-w-full border-separate border-spacing-0">
          <thead>
            <tr class="bg-slate-50 text-xs text-slate-600 uppercase">
              <th class="py-2 px-2 border border-slate-200 text-right">#</th>
              <th class="py-2 px-2 border border-slate-200">Supplier Item #</th>
              <th class="py-2 px-2 border border-slate-200">Peak Part #</th>
              <th class="py-2 px-2 border border-slate-200">Description</th>
              <th class="py-2 px-2 border border-slate-200 text-right">Qty</th>
              <th class="py-2 px-2 border border-slate-200 text-center w-16">UOM</th>
              <th class="py-2 px-2 border border-slate-200 text-right">Unit Price</th>
              <th class="py-2 px-2 border border-slate-200 text-right">Line Total</th>
            </tr>
          </thead>
          <tbody>${rowsHtml || `<tr><td colspan="8" class="text-center text-slate-500 py-4 border border-slate-200">No items</td></tr>`}</tbody>
          <tfoot>
            <tr>
              <td colspan="7" class="text-right pr-3 py-2 border border-slate-200">Subtotal</td>
              <td class="text-right pr-2 py-2 border border-slate-200">${money(subtotal)}</td>
            </tr>
            <tr>
              <td colspan="7" class="text-right pr-3 py-2 border border-slate-200">HST (13%)</td>
              <td class="text-right pr-2 py-2 border border-slate-200">${money(tax)}</td>
            </tr>
            <tr>
              <td colspan="7" class="text-right pr-3 py-2 border border-slate-200 font-medium">Grand Total</td>
              <td class="text-right pr-2 py-2 border border-slate-200 font-semibold">${money(grand)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  `;
}
