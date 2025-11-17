import { TradeRecord } from '../types/trade';

const hsnMapping: Record<string, { description: string; category: string }> = {
  '73239990': { description: 'Household articles of iron or steel', category: 'Steel' },
  '73239900': { description: 'Table, kitchen household articles', category: 'Steel' },
  '73211900': { description: 'Cooking appliances and plate warmers', category: 'Steel' },
  '73239300': { description: 'Kitchen or tableware', category: 'Steel' },
};

export function parseGoodsDescription(description: string = ''): Partial<TradeRecord> {
  const parsed: Partial<TradeRecord> = {};
  const clean = description.trim();

  const qtyMatch = clean.match(/QTY[:\s]*(\d+)/i);
  if (qtyMatch) parsed.quantity = Number(qtyMatch[1]);

  const priceMatch = clean.match(/USD[:\s]*(\d+(\.\d+)?)/i);
  if (priceMatch) parsed.unit_price_usd = Number(priceMatch[1]);

  const modelMatch = clean.match(/MODEL[:\s]*([A-Z0-9-]+)/i);
  if (modelMatch) parsed.model_name = modelMatch[1];

  const lower = clean.toLowerCase();
  if (lower.includes('scrubber')) parsed.sub_category = 'Scrubber';
  else if (lower.includes('container')) parsed.sub_category = 'Container';
  else if (lower.includes('basket')) parsed.sub_category = 'Basket';
  else if (lower.includes('lunch box')) parsed.sub_category = 'Lunch Box';
  else if (lower.includes('cutlery')) parsed.sub_category = 'Cutlery';
  else parsed.sub_category = 'Others';

  return parsed;
}

export function enrichTradeRecord(record: TradeRecord): TradeRecord {
  const enriched: TradeRecord = { ...record } as TradeRecord;
  const hsn = hsnMapping[record.hs_code];
  enriched.hsn_description = hsn?.description ?? 'Unknown';
  enriched.main_category = hsn?.category ?? 'Others';
  enriched.grand_total_inr = Number(record.total_value_inr || 0) + Number(record.duty_paid_inr || 0);
  if (record.date) enriched.year = new Date(record.date).getFullYear();
  Object.assign(enriched, parseGoodsDescription(record.goods_description || ''));
  return enriched;
}

export function parseCSVRow(row: string[]): Partial<TradeRecord> {
  return {
    port_code: row[0] || '',
    date: row[1] || '',
    iec: row[2] || '',
    hs_code: row[3] || '',
    goods_description: row[4] || '',
    model_name: row[6] || '',
    model_number: row[7] || '',
    capacity: row[8] || '',
    quantity: Number(row[13] || 0),
    unit: row[14] || '',
    unit_price_inr: Number(row[15] || 0),
    total_value_inr: Number(row[16] || 0),
    unit_price_usd: Number(row[17] || 0),
    total_value_usd: Number(row[18] || 0),
    duty_paid_inr: Number(row[19] || 0),
  };
}
