import { getSql, json, handleOptions } from './db.js';

export default async (event) => {
  const method = event?.httpMethod || event?.request?.method || 'GET';
  if (method === 'OPTIONS' || method === 'HEAD') return handleOptions();
  if (method !== 'GET') return json({ error: 'Method not allowed' }, 405);

  try {
    const sql = getSql();
    const params = event?.queryStringParameters || {};

    const page = Math.max(1, parseInt(params.page || '1', 10));
    const pageSize = Math.max(1, Math.min(500, parseInt(params.pageSize || '500', 10)));
    const offset = (page - 1) * pageSize;

    // MODIFIED: This query now joins payment data to calculate paid_total and remaining for each PO.
    // This allows the UI to correctly display the "(Partially Paid)" status.
    const rows = await sql`
      SELECT
        po.id,
        po.created_at,
        po.created_by,
        po.department,
        po.vendor_name,
        po.subtotal,
        po.tax,
        po.total,
        po.status,
        po.paid_at,
        (po.status || CASE WHEN po.paid_at IS NOT NULL THEN ' (Paid)' ELSE '' END) AS status_label,
        po.po_number,
        po.meta,
        COUNT(poi.id)::int AS line_items,
        COALESCE(payments.paid_total, 0)::numeric AS paid_total,
        GREATEST(0, po.total::numeric - COALESCE(payments.paid_total, 0))::numeric AS remaining
      FROM purchase_orders po
      LEFT JOIN purchase_order_items poi ON po.id = poi.po_id
      LEFT JOIN (
        SELECT po_id, SUM(amount) AS paid_total
        FROM po_payments
        GROUP BY po_id
      ) AS payments ON po.id = payments.po_id
      GROUP BY po.id, payments.paid_total
      ORDER BY po.created_at DESC
      LIMIT ${pageSize} OFFSET ${offset}
    `;

    const [{ count }] = await sql`SELECT COUNT(*)::int AS count FROM purchase_orders`;

    return json({ ok: true, page, pageSize, count, rows });
  } catch (err) {
    console.error(err);
    return json({ error: err.message || String(err) }, 500);
  }
};