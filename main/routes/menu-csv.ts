import { Router, Request, Response } from 'express';
import { getDatabase, now, generateShortId } from '../db';
import { v4 as uuidv4 } from 'uuid';

const router = Router();

// ─── CSV helpers ─────────────────────────────────────────────────────────────

function parseCSV(text: string): string[][] {
  const rows: string[][] = [];
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  for (const line of lines) {
    if (!line.trim()) continue;
    const fields: string[] = [];
    let i = 0;
    while (i <= line.length) {
      if (i === line.length) { fields.push(''); break; }
      if (line[i] === '"') {
        let val = '';
        i++;
        while (i < line.length) {
          if (line[i] === '"' && line[i + 1] === '"') { val += '"'; i += 2; }
          else if (line[i] === '"') { i++; break; }
          else { val += line[i++]; }
        }
        if (i < line.length && line[i] === ',') i++;
        fields.push(val);
      } else {
        const end = line.indexOf(',', i);
        if (end === -1) { fields.push(line.slice(i)); break; }
        fields.push(line.slice(i, end));
        i = end + 1;
      }
    }
    rows.push(fields);
  }
  return rows;
}

function toObjects(rows: string[][]): Record<string, string>[] {
  if (rows.length < 2) return [];
  const headers = rows[0].map((h) => h.trim().toLowerCase());
  return rows.slice(1).map((row) => {
    const obj: Record<string, string> = {};
    headers.forEach((h, i) => { obj[h] = (row[i] ?? '').trim(); });
    return obj;
  });
}

function toCsvRow(fields: (string | number | null | undefined)[]): string {
  return fields
    .map((f) => {
      const s = String(f ?? '');
      return s.includes(',') || s.includes('"') || s.includes('\n')
        ? '"' + s.replace(/"/g, '""') + '"'
        : s;
    })
    .join(',');
}

function isTruthy(v: string) {
  return ['yes', 'true', '1'].includes((v || '').toLowerCase());
}

// ─── CSV pre-import cleaner ───────────────────────────────────────────────────
// Runs on the raw CSV string BEFORE parsing. Handles messy Excel/Sheets exports.
function cleanCsv(raw: string): string {
  // 1. Strip UTF-8 BOM (Excel adds this)
  let text = raw.replace(/^\uFEFF/, '');

  // 2. Normalise line endings
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  const lines = text.split('\n');
  if (lines.length === 0) return text;

  // 3. Extract header to know which columns are numeric / boolean
  const headerLine = lines[0];
  const headers = headerLine.split(',').map((h) => h.trim().toLowerCase().replace(/^"|"$/g, ''));

  const NUMERIC_COLS  = new Set(['price', 'cost', 'tax_rate', 'cashback_percent', 'sort_order',
                                  'group_min_select', 'group_max_select']);
  const BOOLEAN_COLS  = new Set(['is_active', 'group_required']);

  const cleanedLines: string[] = [headerLine]; // keep original header casing

  for (let li = 1; li < lines.length; li++) {
    const line = lines[li];
    if (!line.trim()) continue; // drop fully empty lines

    // Split respecting quoted fields
    const fields = line.split(',');
    const cleaned: string[] = [];
    let fieldIdx = 0;

    for (let fi = 0; fi < fields.length; fi++) {
      let val = fields[fi];

      // Re-join fields that were split inside quotes
      while (val.startsWith('"') && !val.endsWith('"') && fi + 1 < fields.length) {
        fi++;
        val += ',' + fields[fi];
      }

      // Strip surrounding quotes
      if (val.startsWith('"') && val.endsWith('"')) {
        val = val.slice(1, -1).replace(/""/g, '"');
      }

      // Trim whitespace
      val = val.trim();

      const col = headers[fieldIdx] ?? '';

      // Numeric columns: strip currency symbols, commas used as thousands separators
      if (NUMERIC_COLS.has(col)) {
        val = val.replace(/[₹$€£¥,]/g, '').trim();
        if (val === '' || val === '-') val = '0';
      }

      // Boolean columns: normalise to yes/no
      if (BOOLEAN_COLS.has(col)) {
        const lv = val.toLowerCase();
        if (['true', '1', 'y', 'yes'].includes(lv))        val = 'yes';
        else if (['false', '0', 'n', 'no', ''].includes(lv)) val = 'no';
      }

      // Re-quote if value contains comma or quote
      if (val.includes(',') || val.includes('"') || val.includes('\n')) {
        val = '"' + val.replace(/"/g, '""') + '"';
      }

      cleaned.push(val);
      fieldIdx++;
    }

    // Skip rows where every field is empty (all whitespace/commas)
    if (cleaned.every((v) => v === '' || v === '""')) continue;

    cleanedLines.push(cleaned.join(','));
  }

  return cleanedLines.join('\n');
}

// Canonicalise a tag from CSV so variants like "Non-Veg", "nonveg", "NON VEG"
// are all stored as the standard key (e.g. "non_veg").
function normalizeTag(raw: string): string {
  const s = raw.toLowerCase().replace(/[\s\-_]+/g, '');
  if (s === 'nonveg' || s === 'nonvegetarian') return 'non_veg';
  if (s === 'veg'    || s === 'vegetarian')    return 'veg';
  if (s === 'vegan')                           return 'vegan';
  if (s === 'egg'    || s === 'eggetarian')    return 'egg';
  if (s === 'spicy'  || s === 'hot')           return 'spicy';
  if (s === 'containsnuts' || s === 'nuts')    return 'contains_nuts';
  if (s === 'glutenfree'   || s === 'gf')      return 'gluten_free';
  if (s === 'dairyfree'    || s === 'df')      return 'dairy_free';
  if (s === 'newarrival'   || s === 'new')     return 'new_arrival';
  if (s === 'bestseller'   || s === 'best')    return 'bestseller';
  if (s === 'organic')                         return 'organic';
  if (s === 'fragrancefree')                   return 'fragrance_free';
  if (s === 'limited')                         return 'limited';
  return raw.toLowerCase().replace(/[\s\-]+/g, '_');
}

// ─── Templates ───────────────────────────────────────────────────────────────

const TEMPLATES: Record<string, string> = {
  categories: [
    'name,description,color,icon,sort_order',
    'Beverages,Hot and cold drinks,blue,☕,1',
    'Food,Snacks and meals,green,🍔,2',
    'Desserts,Sweet treats,pink,🍰,3',
    'Combos,Meal deals and bundles,amber,🎁,4',
  ].join('\n'),

  products: [
    'id,sku,name,category,price,description,cost,tax_type,tax_rate,cashback_percent,tags,is_active',
    ',,Cappuccino,Beverages,150,Rich espresso with steamed milk,50,inclusive,5,0,"veg,bestseller",yes',
    ',,Espresso,Beverages,100,,40,inclusive,5,0,veg,yes',
    ',,Cold Coffee,Beverages,130,Chilled blended coffee,45,inclusive,5,0,"veg,new_arrival",yes',
    ',,Classic Burger,Food,250,Juicy patty with lettuce and tomato,100,exclusive,5,0,non_veg,yes',
    ',,Veg Sandwich,Food,180,Fresh vegetables in toasted bread,60,none,0,0,"veg,new_arrival",yes',
    ',,Chocolate Cake,Desserts,120,Rich chocolate slice,,none,0,0,veg,yes',
  ].join('\n'),

  addons: [
    'group_name,addon_name,price,group_required,group_min_select,group_max_select',
    'Size,Small,0,no,1,1',
    'Size,Regular,20,no,1,1',
    'Size,Large,40,no,1,1',
    'Milk Type,Full Cream,0,yes,1,1',
    'Milk Type,Oat Milk,30,yes,1,1',
    'Milk Type,Almond Milk,40,yes,1,1',
    'Extras,Extra Shot,30,no,0,3',
    'Extras,Extra Sugar,0,no,0,3',
    'Temperature,Hot,0,yes,1,1',
    'Temperature,Cold (Iced),10,yes,1,1',
  ].join('\n'),
};

router.get('/template/:type', (req: Request, res: Response) => {
  const { type } = req.params;
  const csv = TEMPLATES[type];
  if (!csv) return res.status(404).json({ error: 'Unknown template type' });
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', `attachment; filename="${type}-template.csv"`);
  res.send(csv);
});

// ─── Export ──────────────────────────────────────────────────────────────────

router.get('/export/categories', (_req: Request, res: Response) => {
  try {
    const db = getDatabase();
    const rows = db
      .prepare('SELECT * FROM categories WHERE deleted_at IS NULL ORDER BY sort_order, name')
      .all() as any[];
    const lines = ['name,description,color,icon,sort_order'];
    for (const c of rows)
      lines.push(toCsvRow([c.name, c.description, c.color, c.icon, c.sort_order]));
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename="categories-export.csv"');
    res.send(lines.join('\n'));
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

router.get('/export/products', (_req: Request, res: Response) => {
  try {
    const db = getDatabase();
    const rows = db
      .prepare(
        `SELECT p.*, c.name AS category_name
         FROM products p
         LEFT JOIN categories c ON p.category_id = c.id
         WHERE p.deleted_at IS NULL
         ORDER BY c.sort_order, p.sort_order, p.name`
      )
      .all() as any[];
    const lines = ['id,sku,name,category,price,description,cost,tax_type,tax_rate,cashback_percent,tags,is_active'];
    for (const p of rows) {
      let tags = '';
      if (p.tags) {
        try { const t = JSON.parse(p.tags); tags = Array.isArray(t) ? t.join(',') : p.tags; }
        catch { tags = p.tags; }
      }
      lines.push(
        toCsvRow([p.id, p.sku, p.name, p.category_name, p.price, p.description, p.cost,
          p.tax_type, p.tax_rate, p.cb_percent ?? 0, tags, p.is_active ? 'yes' : 'no'])
      );
    }
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename="products-export.csv"');
    res.send(lines.join('\n'));
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

router.get('/export/addons', (_req: Request, res: Response) => {
  try {
    const db = getDatabase();
    const groups = db
      .prepare('SELECT * FROM addon_groups WHERE is_active = 1 ORDER BY sort_order, name')
      .all() as any[];
    const lines = ['group_name,addon_name,price,group_required,group_min_select,group_max_select'];
    for (const g of groups) {
      const addons = db
        .prepare('SELECT * FROM addons WHERE addon_group_id = ? AND is_active = 1 ORDER BY sort_order, name')
        .all(g.id) as any[];
      for (const a of addons)
        lines.push(toCsvRow([g.name, a.name, a.price, g.is_required ? 'yes' : 'no', g.min_selection, g.max_selection]));
    }
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename="addons-export.csv"');
    res.send(lines.join('\n'));
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Import ──────────────────────────────────────────────────────────────────

router.post('/import/categories', (req: Request, res: Response) => {
  try {
    const { csv } = req.body as { csv: string };
    if (!csv) return res.status(400).json({ error: 'No CSV data provided' });

    const rows = toObjects(parseCSV(cleanCsv(csv)));
    if (!rows.length) return res.status(400).json({ error: 'CSV has no data rows' });

    const db = getDatabase();
    let created = 0, skipped = 0;
    const errors: string[] = [];
    const warnings: string[] = [];

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r.name) {
        r.name = `Unnamed Category ${i + 1}`;
        warnings.push(`Row ${i + 2}: missing name — imported as "${r.name}"`);
      }

      const exists = db
        .prepare('SELECT id FROM categories WHERE name = ? AND deleted_at IS NULL')
        .get(r.name);
      if (exists) { skipped++; continue; }

      const slug = r.name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
      db.prepare(
        `INSERT INTO categories (id, name, slug, description, color, icon, sort_order, is_active, created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?, ?)`
      ).run(uuidv4(), r.name, slug, r.description || null, r.color || null, r.icon || null,
        parseInt(r.sort_order) || 0, now(), now());
      created++;
    }

    res.json({ created, skipped, errors, warnings });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

router.post('/import/products', (req: Request, res: Response) => {
  try {
    const { csv } = req.body as { csv: string };
    if (!csv) return res.status(400).json({ error: 'No CSV data provided' });

    const rows = toObjects(parseCSV(cleanCsv(csv)));
    if (!rows.length) return res.status(400).json({ error: 'CSV has no data rows' });

    const db = getDatabase();

    const catRows = db
      .prepare('SELECT id, name FROM categories WHERE deleted_at IS NULL')
      .all() as any[];
    const catMap: Record<string, string> = {};
    for (const c of catRows) catMap[c.name.toLowerCase()] = c.id;

    let created = 0, updated = 0, skipped = 0;
    const errors: string[] = [];
    const warnings: string[] = [];

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r.name) {
        r.name = `Unnamed Product ${i + 1}`;
        warnings.push(`Row ${i + 2}: missing name — imported as "${r.name}"`);
      }

      let price = parseFloat(r.price);
      if (isNaN(price)) {
        warnings.push(`Row ${i + 2} (${r.name}): invalid price "${r.price}" — defaulted to 0`);
        price = 0;
      }

      let categoryId: string | null = null;
      if (r.category) {
        categoryId = catMap[r.category.toLowerCase()] ?? null;
        if (!categoryId) {
          warnings.push(`Row ${i + 2} (${r.name}): category "${r.category}" not found — imported without category`);
        }
      }

      let tagsJson: string | null = null;
      if (r.tags) {
        const arr = r.tags.split(',').map((t: string) => normalizeTag(t.trim())).filter(Boolean);
        if (arr.length) tagsJson = JSON.stringify(arr);
      }

      const taxType = ['none', 'inclusive', 'exclusive'].includes(r.tax_type) ? r.tax_type : 'none';
      const isActive = !r.is_active || isTruthy(r.is_active) ? 1 : 0;
      const cost = parseFloat(r.cost) || 0;
      const taxRate = parseFloat(r.tax_rate) || 0;
      const cbPercent = parseFloat(r.cashback_percent) || 0;
      const sku = r.sku || null;

      // If an id is provided, try to update the existing product
      if (r.id) {
        const existing = db
          .prepare('SELECT id FROM products WHERE id = ? AND deleted_at IS NULL')
          .get(r.id);
        if (!existing) {
          errors.push(`Row ${i + 2} (${r.name}): id "${r.id}" not found — leave id blank to create a new item`);
          continue;
        }
        db.prepare(
          `UPDATE products SET name=?, category_id=?, price=?, description=?, cost=?,
           tax_type=?, tax_rate=?, cb_percent=?, tags=?, is_active=?, sku=?, updated_at=?
           WHERE id=?`
        ).run(r.name, categoryId, price, r.description || null, cost,
          taxType, taxRate, cbPercent, tagsJson, isActive, sku, now(), r.id);
        updated++;
        continue;
      }

      // No id — insert as new, skip if name+category duplicate
      const exists = db
        .prepare('SELECT id FROM products WHERE name = ? AND category_id IS ? AND deleted_at IS NULL')
        .get(r.name, categoryId);
      if (exists) { skipped++; continue; }

      db.prepare(
        `INSERT INTO products (id, name, category_id, price, description, cost, tax_type, tax_rate,
         cb_percent, tags, is_active, sku, sort_order, created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?)`
      ).run(generateShortId('products'), r.name, categoryId, price, r.description || null,
        cost, taxType, taxRate, cbPercent, tagsJson, isActive, sku, now(), now());
      created++;
    }

    res.json({ created, updated, skipped, errors, warnings });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

router.post('/import/addons', (req: Request, res: Response) => {
  try {
    const { csv } = req.body as { csv: string };
    if (!csv) return res.status(400).json({ error: 'No CSV data provided' });

    const rows = toObjects(parseCSV(cleanCsv(csv)));
    if (!rows.length) return res.status(400).json({ error: 'CSV has no data rows' });

    const db = getDatabase();
    let groupsCreated = 0, addonsCreated = 0, skipped = 0;
    const errors: string[] = [];
    const warnings: string[] = [];
    const groupCache: Record<string, string> = {};

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r.group_name && !r.addon_name) {
        warnings.push(`Row ${i + 2}: missing group_name and addon_name — row skipped`);
        continue;
      }
      if (!r.group_name) {
        r.group_name = 'General';
        warnings.push(`Row ${i + 2}: missing group_name — defaulted to "General"`);
      }
      if (!r.addon_name) {
        r.addon_name = `Addon ${i + 1}`;
        warnings.push(`Row ${i + 2}: missing addon_name — imported as "${r.addon_name}"`);
      }

      let price = parseFloat(r.price);
      if (isNaN(price)) {
        warnings.push(`Row ${i + 2} (${r.group_name}/${r.addon_name}): invalid price "${r.price}" — defaulted to 0`);
        price = 0;
      }

      const key = r.group_name.toLowerCase();
      let groupId = groupCache[key];
      if (!groupId) {
        const existing = db.prepare('SELECT id FROM addon_groups WHERE name = ?').get(r.group_name) as any;
        if (existing) {
          groupId = existing.id;
        } else {
          groupId = uuidv4();
          db.prepare(
            `INSERT INTO addon_groups (id, name, is_required, min_selection, max_selection, is_active, sort_order, created_at, updated_at)
             VALUES (?, ?, ?, ?, ?, 1, 0, ?, ?)`
          ).run(groupId, r.group_name, isTruthy(r.group_required) ? 1 : 0,
            parseInt(r.group_min_select) || 0, parseInt(r.group_max_select) || 1, now(), now());
          groupsCreated++;
        }
        groupCache[key] = groupId;
      }

      const addonExists = db
        .prepare('SELECT id FROM addons WHERE addon_group_id = ? AND name = ?')
        .get(groupId, r.addon_name);
      if (addonExists) { skipped++; continue; }

      db.prepare(
        `INSERT INTO addons (id, addon_group_id, name, price, is_active, sort_order, created_at, updated_at)
         VALUES (?, ?, ?, ?, 1, 0, ?, ?)`
      ).run(uuidv4(), groupId, r.addon_name, price, now(), now());
      addonsCreated++;
    }

    res.json({ groups_created: groupsCreated, addons_created: addonsCreated, skipped, errors, warnings });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

export { router as menuCsvRoutes };
