import React, { useState, useEffect, useRef, useCallback } from 'react';
import { LayoutGrid, ShoppingCart, BarChart3, Plus, Search, Trash2, X, Package, TrendingUp, AlertTriangle, Minus, Check, Users, Tag, Star, Camera, FileText, Truck, Download } from 'lucide-react';
import * as XLSX from 'xlsx';

// ---------- Design tokens: warm luxury (cream / brass / espresso) ----------
const COLORS = {
  paper: '#F6EFE1',
  espresso: '#2E2117',
  espressoSoft: '#FFFFFF',
  brown: '#DACB9F',
  brass: '#A8783D',
  brassBright: '#8F6B34',
  cream: '#2E2117',
  creamDim: '#8A7A63',
  oxblood: '#8B4038',
};

const uid = () => Math.random().toString(36).slice(2, 9);

function loadScriptOnce(src) {
  return new Promise((resolve, reject) => {
    if (window.Quagga) { resolve(); return; }
    const existing = document.querySelector(`script[src="${src}"]`);
    if (existing) { existing.addEventListener('load', () => resolve()); existing.addEventListener('error', reject); return; }
    const script = document.createElement('script');
    script.src = src;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error('Failed to load scanner library'));
    document.head.appendChild(script);
  });
}
const QUAGGA_SRC = 'https://cdnjs.cloudflare.com/ajax/libs/quagga/0.12.1/quagga.min.js';

const CONDITIONS = ['NM', 'LP', 'MP', 'HP', 'DMG'];
const CONDITION_LABELS = { NM: 'Near Mint', LP: 'Lightly Played', MP: 'Moderately Played', HP: 'Heavily Played', DMG: 'Damaged' };
const LANGUAGES = [{ code: '', label: 'English' }, { code: 'JP', label: 'Japanese' }, { code: 'CN', label: 'Chinese' }];
const VARIANTS = [{ code: '', label: 'Normal' }, { code: 'H', label: 'Holo' }, { code: 'RH', label: 'Reverse Holo' }, { code: 'FA', label: 'Full Art' }, { code: 'IR', label: 'Illustration Rare' }, { code: 'SIR', label: 'Special Illustration Rare' }, { code: 'UR', label: 'Ultra Rare' }, { code: 'HR', label: 'Hyper Rare' }, { code: 'SEC', label: 'Secret Rare' }];

function genSku({ type, game, set, number, variant, condition, language }) {
  if (type === 'accessory') {
    return `ACC-${(game || 'GEN').toUpperCase().slice(0, 3)}-${uid().slice(0, 3).toUpperCase()}`;
  }
  const parts = [];
  if (language) parts.push(language);
  parts.push((game || 'PKM').toUpperCase());
  if (set) parts.push(set.replace(/[^A-Za-z0-9]/g, '').slice(0, 3).toUpperCase());
  if (number) parts.push(number.replace(/[^A-Za-z0-9]/g, ''));
  if (variant) parts.push(variant);
  if (condition && condition !== 'NM') parts.push(condition);
  return parts.join('-');
}

const POINTS_PER_DOLLAR = 1; // 1 point earned per $1 spent
const POINT_REDEMPTION_RATE = 0.05; // 1 point = $0.05 off

// ---------- Persistence (localStorage-backed) ----------
const storageKeys = { items: 'tcg-inventory-items', sales: 'tcg-sales-log', customers: 'tcg-customers', coupons: 'tcg-coupons' };
async function loadKey(key) {
  try { const raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : []; } catch { return []; }
}
async function saveKey(key, val) {
  try { localStorage.setItem(key, JSON.stringify(val)); } catch {}
}

// ---------- Export helpers ----------
function toCSV(rows, headers) {
  const escape = (v) => {
    const s = String(v ?? '');
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [headers.map(h => escape(h.label)).join(',')];
  rows.forEach(row => lines.push(headers.map(h => escape(row[h.key])).join(',')));
  return lines.join('\n');
}
function downloadBlob(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function downloadCSV(filename, rows, headers) {
  downloadBlob(filename, toCSV(rows, headers), 'text/csv;charset=utf-8;');
}
function downloadXLSX(filename, rows, headers, sheetName = 'Report') {
  const data = rows.map(row => {
    const obj = {};
    headers.forEach(h => { obj[h.label] = row[h.key]; });
    return obj;
  });
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

// ---------- Small UI atoms ----------
function StatCard({ label, value, icon: Icon, accent, onClick }) {
  return (
    <div onClick={onClick} style={{ background: COLORS.espressoSoft, borderRadius: 12, padding: '18px 20px', border: `1px solid ${COLORS.brown}`, flex: 1, minWidth: 150, cursor: onClick ? 'pointer' : 'default' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, color: COLORS.creamDim, fontSize: 12, fontFamily: 'Jost, sans-serif', textTransform: 'uppercase', letterSpacing: 1 }}>
        <Icon size={14} color={accent || COLORS.brass} /> {label}
      </div>
      <div style={{ fontFamily: 'Fraunces, serif', fontSize: 28, fontWeight: 700, color: COLORS.cream, marginTop: 6 }}>{value}</div>
    </div>
  );
}
function Field({ label, children }) {
  return (
    <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontFamily: 'Jost, sans-serif', fontSize: 12, color: COLORS.creamDim }}>
      {label}{children}
    </label>
  );
}
const inputStyle = {
  background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 8,
  padding: '9px 10px', color: COLORS.cream, fontFamily: 'Jost, sans-serif', fontSize: 14, outline: 'none',
};
function Button({ children, onClick, variant = 'primary', style, ...rest }) {
  const base = { border: 'none', borderRadius: 8, padding: '10px 16px', fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 13, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6 };
  const variants = {
    primary: { background: COLORS.brass, color: COLORS.espresso },
    ghost: { background: 'transparent', color: COLORS.cream, border: `1px solid ${COLORS.brown}` },
    danger: { background: 'transparent', color: COLORS.oxblood, border: `1px solid ${COLORS.oxblood}` },
  };
  return <button onClick={onClick} style={{ ...base, ...variants[variant], ...style }} {...rest}>{children}</button>;
}
function Modal({ title, onClose, children, width = 420 }) {
  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(10,8,6,0.72)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 50, padding: 16 }}>
      <div style={{ background: COLORS.espressoSoft, borderRadius: 14, padding: 24, width: '100%', maxWidth: width, border: `1px solid ${COLORS.brown}`, maxHeight: '85vh', overflowY: 'auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 18 }}>
          <h3 style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, margin: 0, fontSize: 18 }}>{title}</h3>
          <X size={18} color={COLORS.creamDim} style={{ cursor: 'pointer' }} onClick={onClose} />
        </div>
        {children}
      </div>
    </div>
  );
}

// ---------- Add Item Modal ----------
function AddItemModal({ onClose, onSave }) {
  const [type, setType] = useState('card');
  const [language, setLanguage] = useState('');
  const [game, setGame] = useState('PKM');
  const [name, setName] = useState('');
  const [set, setSet] = useState('');
  const [number, setNumber] = useState('');
  const [variant, setVariant] = useState('');
  const [condition, setCondition] = useState('NM');
  const [cost, setCost] = useState('');
  const [price, setPrice] = useState('');
  const [qty, setQty] = useState('1');
  const [hold, setHold] = useState(false);
  const [schedule, setSchedule] = useState([]); // [{days, percent}]
  const [scheduleDays, setScheduleDays] = useState('');
  const [schedulePercent, setSchedulePercent] = useState('');

  const addScheduleRow = () => {
    const d = parseInt(scheduleDays, 10), p = parseFloat(schedulePercent);
    if (!d || !p) return;
    setSchedule(prev => [...prev, { days: d, percent: p }].sort((a, b) => a.days - b.days));
    setScheduleDays(''); setSchedulePercent('');
  };
  const removeScheduleRow = (idx) => setSchedule(prev => prev.filter((_, i) => i !== idx));

  const canSave = name.trim() && cost !== '' && price !== '' && qty !== '';
  const handleSave = () => {
    if (!canSave) return;
    const item = {
      id: uid(),
      sku: genSku({ type, game, set, number, variant, condition, language }),
      type, language, game: type === 'card' ? game : '', name: name.trim(), set: set.trim(), number: number.trim(),
      variant, condition: type === 'card' ? condition : null,
      cost: parseFloat(cost) || 0, price: parseFloat(price) || 0, originalPrice: parseFloat(price) || 0, qty: parseInt(qty, 10) || 0,
      hold, markdownSchedule: hold ? [] : schedule, appliedMarkdownIndex: -1, dateAdded: new Date().toISOString(),
    };
    onSave(item);
  };

  return (
    <Modal title="Add item" onClose={onClose}>
      <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
        {['card', 'accessory'].map(t => (
          <button key={t} onClick={() => setType(t)} style={{
            flex: 1, padding: '8px', borderRadius: 8, border: `1px solid ${type === t ? COLORS.brass : COLORS.brown}`,
            background: type === t ? COLORS.brass : 'transparent', color: type === t ? COLORS.espresso : COLORS.cream,
            fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 13, cursor: 'pointer', textTransform: 'capitalize',
          }}>{t}</button>
        ))}
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <Field label={type === 'card' ? 'Card name' : 'Item name'}>
          <input style={inputStyle} value={name} onChange={e => setName(e.target.value)} placeholder={type === 'card' ? 'Charizard ex' : 'Ultra Pro Sleeves'} autoFocus />
        </Field>
        {type === 'card' && (
          <>
            <div style={{ display: 'flex', gap: 10 }}>
              <Field label="Game"><input style={inputStyle} value={game} onChange={e => setGame(e.target.value)} placeholder="PKM / MTG / YGO" /></Field>
              <Field label="Language">
                <select style={inputStyle} value={language} onChange={e => setLanguage(e.target.value)}>
                  {LANGUAGES.map(l => <option key={l.code} value={l.code}>{l.label}</option>)}
                </select>
              </Field>
            </div>
            <div style={{ display: 'flex', gap: 10 }}>
              <Field label="Set"><input style={inputStyle} value={set} onChange={e => setSet(e.target.value)} placeholder="OBF" /></Field>
              <Field label="Card #"><input style={inputStyle} value={number} onChange={e => setNumber(e.target.value)} placeholder="054" /></Field>
            </div>
            <Field label="Variant / rarity">
              <select style={inputStyle} value={variant} onChange={e => setVariant(e.target.value)}>
                {VARIANTS.map(v => <option key={v.code} value={v.code}>{v.label}</option>)}
              </select>
            </Field>
            <Field label="Condition (defaults to NM)">
              <select style={inputStyle} value={condition} onChange={e => setCondition(e.target.value)}>
                {CONDITIONS.map(c => <option key={c} value={c}>{c} — {CONDITION_LABELS[c]}</option>)}
              </select>
            </Field>
          </>
        )}
        <div style={{ display: 'flex', gap: 10 }}>
          <Field label="Your cost ($)"><input style={inputStyle} type="number" value={cost} onChange={e => setCost(e.target.value)} placeholder="0.00" /></Field>
          <Field label="Sell price ($)"><input style={inputStyle} type="number" value={price} onChange={e => setPrice(e.target.value)} placeholder="0.00" /></Field>
        </div>
        <Field label="Quantity"><input style={inputStyle} type="number" value={qty} onChange={e => setQty(e.target.value)} placeholder="1" /></Field>

        <div style={{ borderTop: `1px solid ${COLORS.brown}`, paddingTop: 12, marginTop: 4 }}>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: 'Jost, sans-serif', fontSize: 13, color: COLORS.cream, cursor: 'pointer' }}>
            <input type="checkbox" checked={hold} onChange={e => setHold(e.target.checked)} />
            Hold — collectible, never mark down automatically
          </label>
        </div>

        {!hold && (
          <div>
            <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 11, color: COLORS.creamDim, marginBottom: 6, textTransform: 'uppercase', letterSpacing: 1 }}>Markdown schedule (optional)</div>
            {schedule.map((s, idx) => (
              <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 10px', background: COLORS.paper, borderRadius: 6, marginBottom: 6, fontFamily: 'Jost, sans-serif', fontSize: 12, color: COLORS.cream }}>
                <span>After {s.days} days → {s.percent}% off original price</span>
                <X size={12} color={COLORS.creamDim} style={{ cursor: 'pointer' }} onClick={() => removeScheduleRow(idx)} />
              </div>
            ))}
            <div style={{ display: 'flex', gap: 8 }}>
              <input style={{ ...inputStyle, flex: 1 }} type="number" placeholder="Days (e.g. 30)" value={scheduleDays} onChange={e => setScheduleDays(e.target.value)} />
              <input style={{ ...inputStyle, flex: 1 }} type="number" placeholder="% off (e.g. 10)" value={schedulePercent} onChange={e => setSchedulePercent(e.target.value)} />
              <Button variant="ghost" onClick={addScheduleRow}><Plus size={13} /></Button>
            </div>
          </div>
        )}
      </div>
      <div style={{ display: 'flex', gap: 10, marginTop: 20 }}>
        <Button variant="ghost" onClick={onClose} style={{ flex: 1, justifyContent: 'center' }}>Cancel</Button>
        <Button onClick={handleSave} style={{ flex: 1, justifyContent: 'center', opacity: canSave ? 1 : 0.5 }}>Add to inventory</Button>
      </div>
    </Modal>
  );
}

// ---------- Inventory View ----------
function InventoryView({ items, onAdd, onDelete }) {
  const [query, setQuery] = useState('');
  const [showAdd, setShowAdd] = useState(false);
  const filtered = items.filter(i =>
    i.name.toLowerCase().includes(query.toLowerCase()) ||
    i.sku.toLowerCase().includes(query.toLowerCase()) ||
    (i.set || '').toLowerCase().includes(query.toLowerCase())
  );
  return (
    <div>
      <div style={{ display: 'flex', gap: 10, marginBottom: 18, flexWrap: 'wrap' }}>
        <div style={{ position: 'relative', flex: 1, minWidth: 200 }}>
          <Search size={15} color={COLORS.creamDim} style={{ position: 'absolute', left: 10, top: 11 }} />
          <input style={{ ...inputStyle, width: '100%', paddingLeft: 32, boxSizing: 'border-box' }} placeholder="Search or scan SKU..." value={query} onChange={e => setQuery(e.target.value)} />
        </div>
        <Button onClick={() => setShowAdd(true)}><Plus size={15} /> Add item</Button>
      </div>
      {filtered.length === 0 ? (
        <div style={{ textAlign: 'center', padding: '60px 20px', color: COLORS.creamDim, fontFamily: 'Jost, sans-serif' }}>
          {items.length === 0 ? "No inventory yet. Add your first card or accessory to get started." : "No matches for that search."}
        </div>
      ) : (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))', gap: 14 }}>
          {filtered.map(item => (
            <div key={item.id} style={{ position: 'relative', background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 12, padding: '16px 14px', overflow: 'hidden' }}>
              <div style={{ position: 'absolute', top: 0, right: 0, background: COLORS.brass, color: COLORS.espresso, fontFamily: 'JetBrains Mono, monospace', fontSize: 10, fontWeight: 700, padding: '3px 8px', borderRadius: '0 8px 0 8px', letterSpacing: 0.5 }}>
                {item.condition || item.type.toUpperCase()}
              </div>
              <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontSize: 15, fontWeight: 700, marginBottom: 2, paddingRight: 50 }}>{item.name}</div>
              {item.set && <div style={{ fontFamily: 'Jost, sans-serif', color: COLORS.creamDim, fontSize: 12, marginBottom: 8 }}>{item.set} {item.number ? `#${item.number}` : ''}</div>}
              <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 10, color: COLORS.brassBright, marginBottom: 10 }}>{item.sku}</div>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end' }}>
                <div>
                  <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.brass, fontSize: 18, fontWeight: 700 }}>${item.price.toFixed(2)}</div>
                  <div style={{ fontFamily: 'Jost, sans-serif', color: COLORS.creamDim, fontSize: 11 }}>cost ${item.cost.toFixed(2)} · qty {item.qty}</div>
                </div>
                <Trash2 size={15} color={COLORS.creamDim} style={{ cursor: 'pointer' }} onClick={() => onDelete(item.id)} />
              </div>
              {item.qty <= 1 && (
                <div style={{ marginTop: 8, display: 'flex', alignItems: 'center', gap: 5, color: COLORS.oxblood, fontSize: 11, fontFamily: 'Jost, sans-serif' }}>
                  <AlertTriangle size={12} /> Low stock
                </div>
              )}
              {item.dateAdded && (
                <div style={{ marginTop: 6, fontSize: 10, fontFamily: 'Jost, sans-serif', color: COLORS.creamDim }}>
                  {item.hold ? 'Hold — no markdown' : `${Math.floor((Date.now() - new Date(item.dateAdded).getTime()) / MS_DAY)}d in stock`}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
      {showAdd && <AddItemModal onClose={() => setShowAdd(false)} onSave={(item) => { onAdd(item); setShowAdd(false); }} />}
    </div>
  );
}

// ---------- Customers View ----------
function AddCustomerModal({ onClose, onSave }) {
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [email, setEmail] = useState('');
  const canSave = name.trim() && (phone.trim() || email.trim());
  return (
    <Modal title="Add customer" onClose={onClose}>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <Field label="Name"><input style={inputStyle} value={name} onChange={e => setName(e.target.value)} autoFocus placeholder="Jordan Lee" /></Field>
        <Field label="Phone (optional if email given)"><input style={inputStyle} value={phone} onChange={e => setPhone(e.target.value)} placeholder="(555) 123-4567" /></Field>
        <Field label="Email (optional if phone given)"><input style={inputStyle} value={email} onChange={e => setEmail(e.target.value)} placeholder="jordan@email.com" /></Field>
        {!canSave && (name.trim() || phone.trim() || email.trim()) && (
          <div style={{ color: COLORS.oxblood, fontSize: 11, fontFamily: 'Jost, sans-serif' }}>Name plus at least one of phone or email is required.</div>
        )}
      </div>
      <div style={{ display: 'flex', gap: 10, marginTop: 20 }}>
        <Button variant="ghost" onClick={onClose} style={{ flex: 1, justifyContent: 'center' }}>Cancel</Button>
        <Button onClick={() => canSave && onSave({ id: uid(), name: name.trim(), phone: phone.trim(), email: email.trim(), points: 0, wishlist: [] })} style={{ flex: 1, justifyContent: 'center', opacity: canSave ? 1 : 0.5 }}>Add customer</Button>
      </div>
    </Modal>
  );
}
function CustomersView({ customers, sales, onAdd, onUpdate }) {
  const [query, setQuery] = useState('');
  const [showAdd, setShowAdd] = useState(false);
  const [selected, setSelected] = useState(null);
  const [wishInput, setWishInput] = useState('');
  const filtered = customers.filter(c => c.name.toLowerCase().includes(query.toLowerCase()) || (c.phone || '').includes(query) || (c.email || '').toLowerCase().includes(query.toLowerCase()));
  const customerSales = selected ? sales.filter(s => s.customerId === selected.id) : [];

  const addWishlistItem = (customer) => {
    if (!wishInput.trim()) return;
    const updated = { ...customer, wishlist: [...(customer.wishlist || []), wishInput.trim()] };
    onUpdate(updated);
    setSelected(updated);
    setWishInput('');
  };
  const removeWishlistItem = (customer, idx) => {
    const updated = { ...customer, wishlist: customer.wishlist.filter((_, i) => i !== idx) };
    onUpdate(updated);
    setSelected(updated);
  };

  return (
    <div>
      <div style={{ display: 'flex', gap: 10, marginBottom: 18, flexWrap: 'wrap' }}>
        <div style={{ position: 'relative', flex: 1, minWidth: 200 }}>
          <Search size={15} color={COLORS.creamDim} style={{ position: 'absolute', left: 10, top: 11 }} />
          <input style={{ ...inputStyle, width: '100%', paddingLeft: 32, boxSizing: 'border-box' }} placeholder="Search customers..." value={query} onChange={e => setQuery(e.target.value)} />
        </div>
        <Button onClick={() => setShowAdd(true)}><Plus size={15} /> Add customer</Button>
      </div>
      {filtered.length === 0 ? (
        <div style={{ textAlign: 'center', padding: '60px 20px', color: COLORS.creamDim, fontFamily: 'Jost, sans-serif' }}>No customers yet — add one, or they'll be added automatically at checkout.</div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {filtered.map(c => (
            <div key={c.id} onClick={() => setSelected(selected?.id === c.id ? null : c)} style={{ background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 10, padding: '12px 16px', cursor: 'pointer' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div>
                  <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, fontSize: 15 }}>{c.name}</div>
                  <div style={{ fontFamily: 'Jost, sans-serif', color: COLORS.creamDim, fontSize: 12 }}>{[c.phone, c.email].filter(Boolean).join(' · ')}</div>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6, color: COLORS.brass, fontFamily: 'JetBrains Mono, monospace', fontSize: 13 }}>
                  <Star size={13} /> {c.points} pts
                </div>
              </div>
              {selected?.id === c.id && (
                <div style={{ marginTop: 12, paddingTop: 12, borderTop: `1px solid ${COLORS.brown}` }} onClick={e => e.stopPropagation()}>
                  <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 11, color: COLORS.creamDim, marginBottom: 6, textTransform: 'uppercase', letterSpacing: 1 }}>Purchase history</div>
                  {customerSales.length === 0 ? (
                    <div style={{ color: COLORS.creamDim, fontSize: 12, fontFamily: 'Jost, sans-serif', marginBottom: 14 }}>No purchases yet.</div>
                  ) : customerSales.slice().reverse().map(s => (
                    <div key={s.id} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, fontFamily: 'Jost, sans-serif', color: COLORS.creamDim, padding: '4px 0' }}>
                      <span>{s.lines.map(l => `${l.name} ×${l.qty}`).join(', ')}</span>
                      <span style={{ color: COLORS.brass }}>${s.total.toFixed(2)}</span>
                    </div>
                  ))}
                  <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 11, color: COLORS.creamDim, margin: '14px 0 6px', textTransform: 'uppercase', letterSpacing: 1 }}>Wishlist</div>
                  {(c.wishlist || []).length === 0 && <div style={{ color: COLORS.creamDim, fontSize: 12, fontFamily: 'Jost, sans-serif', marginBottom: 8 }}>Nothing on their wishlist yet.</div>}
                  {(c.wishlist || []).map((w, idx) => (
                    <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: 12, fontFamily: 'Jost, sans-serif', color: COLORS.cream, padding: '4px 0' }}>
                      <span>{w}</span>
                      <X size={12} color={COLORS.creamDim} style={{ cursor: 'pointer' }} onClick={() => removeWishlistItem(c, idx)} />
                    </div>
                  ))}
                  <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
                    <input style={{ ...inputStyle, flex: 1 }} placeholder="Add item they want..." value={wishInput} onChange={e => setWishInput(e.target.value)} onKeyDown={e => e.key === 'Enter' && addWishlistItem(c)} />
                    <Button variant="ghost" onClick={() => addWishlistItem(c)}><Plus size={13} /></Button>
                  </div>
                </div>
              )}
            </div>
          ))}
        </div>
      )}
      {showAdd && <AddCustomerModal onClose={() => setShowAdd(false)} onSave={(c) => { onAdd(c); setShowAdd(false); }} />}
    </div>
  );
}

// ---------- Coupons View ----------
function AddCouponModal({ onClose, onSave }) {
  const [code, setCode] = useState('');
  const [type, setType] = useState('percent');
  const [value, setValue] = useState('');
  const canSave = code.trim() && value !== '';
  return (
    <Modal title="Add coupon" onClose={onClose}>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <Field label="Code"><input style={{ ...inputStyle, textTransform: 'uppercase' }} value={code} onChange={e => setCode(e.target.value.toUpperCase())} placeholder="WELCOME10" autoFocus /></Field>
        <div style={{ display: 'flex', gap: 8 }}>
          {['percent', 'fixed'].map(t => (
            <button key={t} onClick={() => setType(t)} style={{ flex: 1, padding: '8px', borderRadius: 8, border: `1px solid ${type === t ? COLORS.brass : COLORS.brown}`, background: type === t ? COLORS.brass : 'transparent', color: type === t ? COLORS.espresso : COLORS.cream, fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 13, cursor: 'pointer' }}>
              {t === 'percent' ? '% off' : '$ off'}
            </button>
          ))}
        </div>
        <Field label={type === 'percent' ? 'Percent off' : 'Dollar amount off'}>
          <input style={inputStyle} type="number" value={value} onChange={e => setValue(e.target.value)} placeholder={type === 'percent' ? '10' : '5.00'} />
        </Field>
      </div>
      <div style={{ display: 'flex', gap: 10, marginTop: 20 }}>
        <Button variant="ghost" onClick={onClose} style={{ flex: 1, justifyContent: 'center' }}>Cancel</Button>
        <Button onClick={() => canSave && onSave({ id: uid(), code: code.trim(), type, value: parseFloat(value) || 0, active: true })} style={{ flex: 1, justifyContent: 'center', opacity: canSave ? 1 : 0.5 }}>Add coupon</Button>
      </div>
    </Modal>
  );
}
function CouponsView({ coupons, onAdd, onDelete }) {
  const [showAdd, setShowAdd] = useState(false);
  return (
    <div>
      <div style={{ marginBottom: 18 }}><Button onClick={() => setShowAdd(true)}><Plus size={15} /> Add coupon</Button></div>
      {coupons.length === 0 ? (
        <div style={{ textAlign: 'center', padding: '60px 20px', color: COLORS.creamDim, fontFamily: 'Jost, sans-serif' }}>No coupons yet. Add one to offer discounts at checkout.</div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {coupons.map(c => (
            <div key={c.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 10, padding: '12px 16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <Tag size={15} color={COLORS.brass} />
                <span style={{ fontFamily: 'JetBrains Mono, monospace', color: COLORS.cream, fontWeight: 700 }}>{c.code}</span>
                <span style={{ fontFamily: 'Jost, sans-serif', color: COLORS.creamDim, fontSize: 13 }}>{c.type === 'percent' ? `${c.value}% off` : `$${c.value.toFixed(2)} off`}</span>
              </div>
              <Trash2 size={15} color={COLORS.creamDim} style={{ cursor: 'pointer' }} onClick={() => onDelete(c.id)} />
            </div>
          ))}
        </div>
      )}
      {showAdd && <AddCouponModal onClose={() => setShowAdd(false)} onSave={(c) => { onAdd(c); setShowAdd(false); }} />}
    </div>
  );
}

// ---------- Quick Add Customer (used inline from POS) ----------
function QuickAddCustomerModal({ onClose, onSave }) {
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [email, setEmail] = useState('');
  const canSave = name.trim() && (phone.trim() || email.trim());
  return (
    <Modal title="New customer" onClose={onClose} width={360}>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <Field label="Name"><input style={inputStyle} value={name} onChange={e => setName(e.target.value)} autoFocus placeholder="Jordan Lee" /></Field>
        <Field label="Phone"><input style={inputStyle} value={phone} onChange={e => setPhone(e.target.value)} placeholder="(555) 123-4567" /></Field>
        <Field label="Email"><input style={inputStyle} value={email} onChange={e => setEmail(e.target.value)} placeholder="jordan@email.com" /></Field>
        {!canSave && (name.trim() || phone.trim() || email.trim()) && (
          <div style={{ color: COLORS.oxblood, fontSize: 11, fontFamily: 'Jost, sans-serif' }}>Name plus phone or email is required.</div>
        )}
      </div>
      <div style={{ display: 'flex', gap: 10, marginTop: 20 }}>
        <Button variant="ghost" onClick={onClose} style={{ flex: 1, justifyContent: 'center' }}>Cancel</Button>
        <Button onClick={() => canSave && onSave({ id: uid(), name: name.trim(), phone: phone.trim(), email: email.trim(), points: 0, wishlist: [] })} style={{ flex: 1, justifyContent: 'center', opacity: canSave ? 1 : 0.5 }}>Add & select</Button>
      </div>
    </Modal>
  );
}

// ---------- Camera Scan Modal ----------
function CameraScanModal({ onDetect, onClose }) {
  const videoTargetRef = useRef(null);
  const [status, setStatus] = useState('loading'); // loading | scanning | error
  const [errorMsg, setErrorMsg] = useState('');
  const [lastScanned, setLastScanned] = useState('');
  const lastCodeRef = useRef({ code: '', time: 0 });

  useEffect(() => {
    let cancelled = false;

    const handleDetected = (result) => {
      const code = result?.codeResult?.code;
      if (!code) return;
      const now = Date.now();
      if (code === lastCodeRef.current.code && now - lastCodeRef.current.time < 1500) return;
      lastCodeRef.current = { code, time: now };
      setLastScanned(code);
      onDetect(code);
    };

    loadScriptOnce(QUAGGA_SRC)
      .then(() => {
        if (cancelled || !videoTargetRef.current) return;
        if (!window.Quagga) { setStatus('error'); setErrorMsg('Scanner library failed to load.'); return; }
        window.Quagga.init({
          inputStream: { name: 'Live', type: 'LiveStream', target: videoTargetRef.current, constraints: { facingMode: 'environment' } },
          decoder: { readers: ['code_128_reader', 'ean_reader', 'ean_8_reader', 'upc_reader', 'upc_e_reader', 'code_39_reader'] },
          locate: true,
        }, (err) => {
          if (cancelled) return;
          if (err) {
            setStatus('error');
            setErrorMsg(err.name === 'NotAllowedError' ? 'Camera access was denied. Check your browser/site permissions and try again.' : 'Could not start the camera on this device.');
            return;
          }
          window.Quagga.start();
          window.Quagga.onDetected(handleDetected);
          setStatus('scanning');
        });
      })
      .catch(() => { if (!cancelled) { setStatus('error'); setErrorMsg('Scanner library failed to load — check your connection.'); } });

    return () => {
      cancelled = true;
      try { window.Quagga && window.Quagga.offDetected(handleDetected); window.Quagga && window.Quagga.stop(); } catch (e) {}
    };
  }, []);

  return (
    <Modal title="Scan with camera" onClose={onClose} width={420}>
      <div style={{ borderRadius: 10, overflow: 'hidden', background: '#000', position: 'relative', minHeight: 220 }}>
        <div ref={videoTargetRef} style={{ width: '100%' }} />
        {status === 'loading' && (
          <div style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>
            Starting camera...
          </div>
        )}
        {status === 'error' && (
          <div style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', padding: 20, color: COLORS.oxblood, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>
            {errorMsg}
          </div>
        )}
      </div>
      <div style={{ marginTop: 12, fontFamily: 'Jost, sans-serif', fontSize: 12, color: COLORS.creamDim }}>
        {status === 'scanning' ? 'Point the camera at a barcode — it will add automatically.' : ' '}
      </div>
      {lastScanned && (
        <div style={{ marginTop: 8, padding: '8px 12px', background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 8, fontFamily: 'JetBrains Mono, monospace', fontSize: 12, color: COLORS.brass }}>
          Last scanned: {lastScanned}
        </div>
      )}
      <div style={{ display: 'flex', marginTop: 16 }}>
        <Button variant="ghost" onClick={onClose} style={{ flex: 1, justifyContent: 'center' }}>Done</Button>
      </div>
    </Modal>
  );
}

// ---------- POS View ----------
function POSView({ items, customers, coupons, onCheckout, onAddCustomer }) {
  const [query, setQuery] = useState('');
  const [cart, setCart] = useState([]);
  const [customerQuery, setCustomerQuery] = useState('');
  const [selectedCustomer, setSelectedCustomer] = useState(null);
  const [showQuickAdd, setShowQuickAdd] = useState(false);
  const [showCameraScan, setShowCameraScan] = useState(false);
  const [couponCode, setCouponCode] = useState('');
  const [couponError, setCouponError] = useState('');
  const [redeemPoints, setRedeemPoints] = useState('');
  const [isOnlineOrder, setIsOnlineOrder] = useState(false);
  const [shippingAddress, setShippingAddress] = useState('');
  const inputRef = useRef(null);

  useEffect(() => { inputRef.current?.focus(); }, []);

  const addToCartBySku = useCallback((sku) => {
    const item = items.find(i => i.sku.toLowerCase() === sku.toLowerCase());
    if (!item) return;
    setCart(prev => {
      const existing = prev.find(c => c.itemId === item.id);
      if (existing) return prev.map(c => c.itemId === item.id ? { ...c, qty: Math.min(c.qty + 1, item.qty) } : c);
      return [...prev, { itemId: item.id, qty: 1 }];
    });
    setQuery('');
  }, [items]);

  const handleKeyDown = (e) => { if (e.key === 'Enter' && query) addToCartBySku(query); };
  const addByClick = (item) => setCart(prev => {
    const existing = prev.find(c => c.itemId === item.id);
    if (existing) return prev.map(c => c.itemId === item.id ? { ...c, qty: Math.min(c.qty + 1, item.qty) } : c);
    return [...prev, { itemId: item.id, qty: 1 }];
  });
  const searchResults = query && !items.some(i => i.sku.toLowerCase() === query.toLowerCase())
    ? items.filter(i => i.name.toLowerCase().includes(query.toLowerCase())).slice(0, 6) : [];

  const updateQty = (itemId, delta) => setCart(prev => prev.map(c => {
    if (c.itemId !== itemId) return c;
    const item = items.find(i => i.id === itemId);
    return { ...c, qty: Math.max(1, Math.min(c.qty + delta, item?.qty ?? c.qty)) };
  }));
  const removeFromCart = (itemId) => setCart(prev => prev.filter(c => c.itemId !== itemId));

  const cartLines = cart.map(c => ({ ...c, item: items.find(i => i.id === c.itemId) })).filter(l => l.item);
  const subtotal = cartLines.reduce((sum, l) => sum + l.item.price * l.qty, 0);

  const appliedCoupon = coupons.find(c => c.code.toLowerCase() === couponCode.toLowerCase() && c.active);
  const couponDiscount = appliedCoupon ? (appliedCoupon.type === 'percent' ? subtotal * (appliedCoupon.value / 100) : Math.min(appliedCoupon.value, subtotal)) : 0;

  const pointsToRedeem = Math.max(0, Math.min(parseInt(redeemPoints || '0', 10), selectedCustomer?.points || 0));
  const pointsDiscount = pointsToRedeem * POINT_REDEMPTION_RATE;

  const total = Math.max(0, subtotal - couponDiscount - pointsDiscount);
  const pointsEarned = Math.floor(total * POINTS_PER_DOLLAR);

  const customerMatches = customerQuery ? customers.filter(c => c.name.toLowerCase().includes(customerQuery.toLowerCase())) : [];

  const handleCheckout = () => {
    if (cartLines.length === 0) return;
    onCheckout({ cartLines, customer: selectedCustomer, couponCode: appliedCoupon?.code, couponDiscount, pointsRedeemed: pointsToRedeem, pointsDiscount, pointsEarned, total, isOnlineOrder, shippingAddress: isOnlineOrder ? shippingAddress : '' });
    setCart([]); setCouponCode(''); setRedeemPoints(''); setSelectedCustomer(null); setCustomerQuery(''); setIsOnlineOrder(false); setShippingAddress('');
  };

  return (
    <div style={{ display: 'flex', gap: 20, flexWrap: 'wrap' }}>
      <div style={{ flex: 2, minWidth: 260 }}>
        <div style={{ display: 'flex', gap: 8, marginBottom: 14 }}>
          <div style={{ position: 'relative', flex: 1 }}>
            <Search size={15} color={COLORS.creamDim} style={{ position: 'absolute', left: 10, top: 13 }} />
            <input ref={inputRef} style={{ ...inputStyle, width: '100%', paddingLeft: 32, boxSizing: 'border-box', fontSize: 16, padding: '11px 10px 11px 32px' }}
              placeholder="Scan barcode or type item name..." value={query} onChange={e => setQuery(e.target.value)} onKeyDown={handleKeyDown} />
          </div>
          <Button variant="ghost" onClick={() => setShowCameraScan(true)} style={{ flexShrink: 0 }}><Camera size={15} /></Button>
        </div>
        {showCameraScan && (
          <CameraScanModal
            onDetect={(code) => addToCartBySku(code)}
            onClose={() => setShowCameraScan(false)}
          />
        )}
        {searchResults.length > 0 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 16 }}>
            {searchResults.map(item => (
              <div key={item.id} onClick={() => { addByClick(item); setQuery(''); }} style={{ display: 'flex', justifyContent: 'space-between', padding: '10px 12px', background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 8, cursor: 'pointer', fontFamily: 'Jost, sans-serif' }}>
                <span style={{ color: COLORS.cream, fontSize: 13 }}>{item.name} {item.condition ? `(${item.condition})` : ''}</span>
                <span style={{ color: COLORS.brass, fontSize: 13, fontWeight: 600 }}>${item.price.toFixed(2)}</span>
              </div>
            ))}
          </div>
        )}

        {/* Customer selector */}
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 11, color: COLORS.creamDim, marginBottom: 6, textTransform: 'uppercase', letterSpacing: 1 }}>Customer (optional)</div>
          {selectedCustomer ? (
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 8, padding: '10px 14px' }}>
              <span style={{ color: COLORS.cream, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>{selectedCustomer.name} · <span style={{ color: COLORS.brass }}>{selectedCustomer.points} pts</span></span>
              <X size={14} color={COLORS.creamDim} style={{ cursor: 'pointer' }} onClick={() => setSelectedCustomer(null)} />
            </div>
          ) : (
            <>
              <input style={{ ...inputStyle, width: '100%', boxSizing: 'border-box' }} placeholder="Search customer by name..." value={customerQuery} onChange={e => setCustomerQuery(e.target.value)} />
              {customerMatches.length > 0 && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4, marginTop: 6 }}>
                  {customerMatches.slice(0, 5).map(c => (
                    <div key={c.id} onClick={() => { setSelectedCustomer(c); setCustomerQuery(''); }} style={{ padding: '8px 12px', background: COLORS.espressoSoft, borderRadius: 6, cursor: 'pointer', fontFamily: 'Jost, sans-serif', fontSize: 13, color: COLORS.cream }}>
                      {c.name} <span style={{ color: COLORS.creamDim }}>({c.points} pts)</span>
                    </div>
                  ))}
                </div>
              )}
              <div onClick={() => setShowQuickAdd(true)} style={{ marginTop: 6, display: 'flex', alignItems: 'center', gap: 5, color: COLORS.brass, fontFamily: 'Jost, sans-serif', fontSize: 12, cursor: 'pointer', fontWeight: 600 }}>
                <Plus size={12} /> New customer
              </div>
            </>
          )}
        </div>

        {showQuickAdd && (
          <QuickAddCustomerModal
            onClose={() => setShowQuickAdd(false)}
            onSave={(c) => { onAddCustomer(c); setSelectedCustomer(c); setShowQuickAdd(false); }}
          />
        )}

        {selectedCustomer && selectedCustomer.points > 0 && (
          <Field label={`Redeem points (have ${selectedCustomer.points}, ${POINT_REDEMPTION_RATE.toFixed(2)}/pt)`}>
            <input style={inputStyle} type="number" value={redeemPoints} onChange={e => setRedeemPoints(e.target.value)} placeholder="0" />
          </Field>
        )}
      </div>

      <div style={{ flex: 1, minWidth: 260, background: COLORS.espressoSoft, borderRadius: 12, border: `1px solid ${COLORS.brown}`, padding: 18 }}>
        <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6 }}>
          <ShoppingCart size={16} color={COLORS.brass} /> Current sale
        </div>
        {cartLines.length === 0 ? (
          <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>Cart is empty.</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginBottom: 14 }}>
            {cartLines.map(l => (
              <div key={l.itemId} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ flex: 1 }}>
                  <div style={{ color: COLORS.cream, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>{l.item.name}</div>
                  <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 11 }}>${l.item.price.toFixed(2)} each</div>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <Minus size={13} color={COLORS.creamDim} style={{ cursor: 'pointer' }} onClick={() => updateQty(l.itemId, -1)} />
                  <span style={{ color: COLORS.cream, fontFamily: 'JetBrains Mono, monospace', fontSize: 13, minWidth: 16, textAlign: 'center' }}>{l.qty}</span>
                  <Plus size={13} color={COLORS.creamDim} style={{ cursor: 'pointer' }} onClick={() => updateQty(l.itemId, 1)} />
                  <X size={14} color={COLORS.oxblood} style={{ cursor: 'pointer', marginLeft: 4 }} onClick={() => removeFromCart(l.itemId)} />
                </div>
              </div>
            ))}
          </div>
        )}

        <div style={{ marginBottom: 12 }}>
          <input style={{ ...inputStyle, width: '100%', boxSizing: 'border-box', textTransform: 'uppercase' }} placeholder="Coupon code" value={couponCode}
            onChange={e => { setCouponCode(e.target.value.toUpperCase()); setCouponError(''); }} />
          {couponCode && !appliedCoupon && <div style={{ color: COLORS.oxblood, fontSize: 11, marginTop: 4, fontFamily: 'Jost, sans-serif' }}>No matching active coupon.</div>}
        </div>

        <div style={{ borderTop: `1px solid ${COLORS.brown}`, paddingTop: 12, marginBottom: 14 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', marginBottom: 4 }}>
            <span>Subtotal</span><span>${subtotal.toFixed(2)}</span>
          </div>
          {couponDiscount > 0 && <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: COLORS.brass, fontFamily: 'Jost, sans-serif', marginBottom: 4 }}><span>Coupon ({appliedCoupon.code})</span><span>-${couponDiscount.toFixed(2)}</span></div>}
          {pointsDiscount > 0 && <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: COLORS.brass, fontFamily: 'Jost, sans-serif', marginBottom: 4 }}><span>Points redeemed ({pointsToRedeem})</span><span>-${pointsDiscount.toFixed(2)}</span></div>}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 8 }}>
            <span style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>Total</span>
            <span style={{ color: COLORS.brass, fontFamily: 'Fraunces, serif', fontSize: 22, fontWeight: 700 }}>${total.toFixed(2)}</span>
          </div>
          {selectedCustomer && <div style={{ fontSize: 11, color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', marginTop: 4 }}>Earns {pointsEarned} pts for {selectedCustomer.name}</div>}
        </div>

        <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: 'Jost, sans-serif', fontSize: 12, color: COLORS.cream, cursor: 'pointer', marginBottom: 10 }}>
          <input type="checkbox" checked={isOnlineOrder} onChange={e => setIsOnlineOrder(e.target.checked)} />
          <Truck size={13} color={COLORS.brass} /> Online order — ships to customer
        </label>
        {isOnlineOrder && (
          <textarea
            style={{ ...inputStyle, width: '100%', boxSizing: 'border-box', minHeight: 60, marginBottom: 12, fontFamily: 'Jost, sans-serif', resize: 'vertical' }}
            placeholder="Shipping address..."
            value={shippingAddress}
            onChange={e => setShippingAddress(e.target.value)}
          />
        )}

        <Button onClick={handleCheckout} style={{ width: '100%', justifyContent: 'center', opacity: cartLines.length ? 1 : 0.5 }}>
          <Check size={15} /> Complete sale
        </Button>
      </div>
    </div>
  );
}

// ---------- Dashboard ----------
const MS_DAY = 1000 * 60 * 60 * 24;
function getMarkdownSuggestion(item) {
  if (item.hold || !item.markdownSchedule || item.markdownSchedule.length === 0) return null;
  const ageDays = (Date.now() - new Date(item.dateAdded).getTime()) / MS_DAY;
  const appliedIdx = item.appliedMarkdownIndex ?? -1;
  let best = null;
  item.markdownSchedule.forEach((tier, idx) => {
    if (idx > appliedIdx && ageDays >= tier.days) best = { idx, tier };
  });
  if (!best) return null;
  const base = item.originalPrice ?? item.price;
  const newPrice = Math.max(0, base * (1 - best.tier.percent / 100));
  return { tierIndex: best.idx, newPrice, days: Math.floor(ageDays), percent: best.tier.percent };
}

function DetailModal({ title, onClose, children }) {
  return (
    <Modal title={title} onClose={onClose} width={520}>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxHeight: '55vh', overflowY: 'auto' }}>{children}</div>
    </Modal>
  );
}

function Dashboard({ items, sales, customers, onApplyMarkdown }) {
  const [detail, setDetail] = useState(null); // 'value' | 'cost' | 'revenue' | 'customers' | 'lowStock' | null

  const inventoryValue = items.reduce((sum, i) => sum + i.price * i.qty, 0);
  const inventoryCost = items.reduce((sum, i) => sum + i.cost * i.qty, 0);
  const totalRevenue = sales.reduce((sum, s) => sum + s.total, 0);
  const lowStock = items.filter(i => i.qty <= 1);

  const topSellers = (() => {
    const map = {};
    sales.forEach(s => s.lines.forEach(l => {
      if (!map[l.name]) map[l.name] = { name: l.name, qty: 0, revenue: 0 };
      map[l.name].qty += l.qty;
      map[l.name].revenue += l.qty * l.price;
    }));
    return Object.values(map).sort((a, b) => b.qty - a.qty).slice(0, 8);
  })();

  const markdownSuggestions = items
    .map(i => ({ item: i, suggestion: getMarkdownSuggestion(i) }))
    .filter(x => x.suggestion);

  const rowStyle = { display: 'flex', justifyContent: 'space-between', background: COLORS.paper, borderRadius: 8, padding: '10px 12px', fontFamily: 'Jost, sans-serif', fontSize: 13 };

  return (
    <div>
      <div style={{ display: 'flex', gap: 14, flexWrap: 'wrap', marginBottom: 24 }}>
        <StatCard label="Inventory value" value={`$${inventoryValue.toFixed(2)}`} icon={Package} onClick={() => setDetail('value')} />
        <StatCard label="Inventory cost" value={`$${inventoryCost.toFixed(2)}`} icon={TrendingUp} onClick={() => setDetail('cost')} />
        <StatCard label="Total revenue" value={`$${totalRevenue.toFixed(2)}`} icon={BarChart3} onClick={() => setDetail('revenue')} />
        <StatCard label="Customers" value={customers.length} icon={Users} onClick={() => setDetail('customers')} />
        <StatCard label="Low stock" value={lowStock.length} icon={AlertTriangle} accent={COLORS.oxblood} onClick={() => setDetail('lowStock')} />
      </div>

      {markdownSuggestions.length > 0 && (
        <div style={{ marginBottom: 24 }}>
          <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, marginBottom: 10, fontSize: 15 }}>Markdown suggestions</div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {markdownSuggestions.map(({ item, suggestion }) => (
              <div key={item.id} style={{ ...rowStyle, background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, alignItems: 'center' }}>
                <span style={{ color: COLORS.cream }}>{item.name} — {suggestion.days} days old, suggest {suggestion.percent}% off → ${suggestion.newPrice.toFixed(2)}</span>
                <Button onClick={() => onApplyMarkdown(item.id, suggestion.tierIndex, suggestion.newPrice)} style={{ padding: '6px 12px', fontSize: 12 }}>Approve</Button>
              </div>
            ))}
          </div>
        </div>
      )}

      <div style={{ marginBottom: 24 }}>
        <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, marginBottom: 10, fontSize: 15 }}>Top sellers</div>
        {topSellers.length === 0 ? (
          <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>No sales yet — top sellers will show up here once you start ringing up sales.</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {topSellers.map((t, idx) => (
              <div key={t.name} style={rowStyle}>
                <span style={{ color: COLORS.cream }}><span style={{ color: COLORS.brass, fontFamily: 'JetBrains Mono, monospace', marginRight: 8 }}>#{idx + 1}</span>{t.name} · {t.qty} sold</span>
                <span style={{ color: COLORS.brass, fontWeight: 600 }}>${t.revenue.toFixed(2)}</span>
              </div>
            ))}
          </div>
        )}
      </div>

      <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, marginBottom: 10, fontSize: 15 }}>Recent sales</div>
      {sales.length === 0 ? (
        <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>No sales logged yet.</div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {sales.slice().reverse().slice(0, 8).map(s => (
            <div key={s.id} style={{ display: 'flex', justifyContent: 'space-between', background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 8, padding: '10px 14px', fontFamily: 'Jost, sans-serif', fontSize: 13 }}>
              <span style={{ color: COLORS.cream }}>{s.lines.map(l => `${l.name} ×${l.qty}`).join(', ')}{s.customerName ? ` · ${s.customerName}` : ''}</span>
              <span style={{ color: COLORS.brass, fontWeight: 600 }}>${s.total.toFixed(2)}</span>
            </div>
          ))}
        </div>
      )}

      {detail === 'value' && (
        <DetailModal title="Inventory value breakdown" onClose={() => setDetail(null)}>
          {items.slice().sort((a, b) => (b.price * b.qty) - (a.price * a.qty)).map(i => (
            <div key={i.id} style={rowStyle}>
              <span style={{ color: COLORS.cream }}>{i.name} · qty {i.qty} × ${i.price.toFixed(2)}</span>
              <span style={{ color: COLORS.brass, fontWeight: 600 }}>${(i.price * i.qty).toFixed(2)}</span>
            </div>
          ))}
        </DetailModal>
      )}
      {detail === 'cost' && (
        <DetailModal title="Inventory cost breakdown" onClose={() => setDetail(null)}>
          {items.slice().sort((a, b) => (b.cost * b.qty) - (a.cost * a.qty)).map(i => (
            <div key={i.id} style={rowStyle}>
              <span style={{ color: COLORS.cream }}>{i.name} · qty {i.qty} × ${i.cost.toFixed(2)}</span>
              <span style={{ color: COLORS.brass, fontWeight: 600 }}>${(i.cost * i.qty).toFixed(2)}</span>
            </div>
          ))}
        </DetailModal>
      )}
      {detail === 'revenue' && (
        <DetailModal title="All sales" onClose={() => setDetail(null)}>
          {sales.slice().reverse().map(s => (
            <div key={s.id} style={rowStyle}>
              <span style={{ color: COLORS.cream }}>{new Date(s.date).toLocaleDateString()} · {s.lines.map(l => `${l.name} ×${l.qty}`).join(', ')}{s.customerName ? ` · ${s.customerName}` : ''}</span>
              <span style={{ color: COLORS.brass, fontWeight: 600 }}>${s.total.toFixed(2)}</span>
            </div>
          ))}
        </DetailModal>
      )}
      {detail === 'customers' && (
        <DetailModal title="All customers" onClose={() => setDetail(null)}>
          {customers.map(c => (
            <div key={c.id} style={rowStyle}>
              <span style={{ color: COLORS.cream }}>{c.name} · {[c.phone, c.email].filter(Boolean).join(' · ')}</span>
              <span style={{ color: COLORS.brass, fontWeight: 600 }}>{c.points} pts</span>
            </div>
          ))}
        </DetailModal>
      )}
      {detail === 'lowStock' && (
        <DetailModal title="Low stock items" onClose={() => setDetail(null)}>
          {lowStock.length === 0 ? <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>Nothing low right now.</div> : lowStock.map(i => (
            <div key={i.id} style={rowStyle}>
              <span style={{ color: COLORS.cream }}>{i.name}</span>
              <span style={{ color: COLORS.oxblood, fontWeight: 600 }}>qty {i.qty}</span>
            </div>
          ))}
        </DetailModal>
      )}
    </div>
  );
}

// ---------- Reports View ----------
const REPORT_TYPES = [
  { id: 'sales', label: 'Sales' },
  { id: 'inventory', label: 'Inventory' },
  { id: 'customers', label: 'Customers' },
  { id: 'topSellers', label: 'Top Sellers' },
];

function ReportsView({ items, sales, customers }) {
  const [reportType, setReportType] = useState('sales');
  const [fromDate, setFromDate] = useState('');
  const [toDate, setToDate] = useState('');
  const printRef = useRef(null);

  const filteredSales = sales.filter(s => {
    const d = new Date(s.date);
    if (fromDate && d < new Date(fromDate)) return false;
    if (toDate && d > new Date(new Date(toDate).getTime() + MS_DAY)) return false;
    return true;
  });

  const reportConfig = {
    sales: {
      headers: [
        { key: 'date', label: 'Date' }, { key: 'items', label: 'Items' }, { key: 'customer', label: 'Customer' },
        { key: 'coupon', label: 'Coupon' }, { key: 'total', label: 'Total' },
      ],
      rows: filteredSales.map(s => ({
        date: new Date(s.date).toLocaleDateString(),
        items: s.lines.map(l => `${l.name} x${l.qty}`).join('; '),
        customer: s.customerName || '',
        coupon: s.couponCode || '',
        total: s.total.toFixed(2),
      })),
    },
    inventory: {
      headers: [
        { key: 'name', label: 'Name' }, { key: 'sku', label: 'SKU' }, { key: 'qty', label: 'Qty' },
        { key: 'cost', label: 'Cost' }, { key: 'price', label: 'Price' }, { key: 'value', label: 'Value' },
      ],
      rows: items.map(i => ({ name: i.name, sku: i.sku, qty: i.qty, cost: i.cost.toFixed(2), price: i.price.toFixed(2), value: (i.price * i.qty).toFixed(2) })),
    },
    customers: {
      headers: [
        { key: 'name', label: 'Name' }, { key: 'phone', label: 'Phone' }, { key: 'email', label: 'Email' }, { key: 'points', label: 'Points' },
      ],
      rows: customers.map(c => ({ name: c.name, phone: c.phone || '', email: c.email || '', points: c.points })),
    },
    topSellers: {
      headers: [{ key: 'name', label: 'Item' }, { key: 'qty', label: 'Units sold' }, { key: 'revenue', label: 'Revenue' }],
      rows: (() => {
        const map = {};
        filteredSales.forEach(s => s.lines.forEach(l => {
          if (!map[l.name]) map[l.name] = { name: l.name, qty: 0, revenue: 0 };
          map[l.name].qty += l.qty; map[l.name].revenue += l.qty * l.price;
        }));
        return Object.values(map).sort((a, b) => b.qty - a.qty).map(r => ({ ...r, revenue: r.revenue.toFixed(2) }));
      })(),
    },
  };

  const { headers, rows } = reportConfig[reportType];
  const label = REPORT_TYPES.find(r => r.id === reportType).label;
  const filenameBase = `pullbar-${reportType}-${new Date().toISOString().slice(0, 10)}`;

  const handlePrint = () => {
    const win = window.open('', '_blank');
    const tableRows = rows.map(r => `<tr>${headers.map(h => `<td style="padding:6px 10px;border-bottom:1px solid #ddd;">${r[h.key]}</td>`).join('')}</tr>`).join('');
    win.document.write(`
      <html><head><title>${label} Report</title></head>
      <body style="font-family: sans-serif; padding: 24px;">
        <h2>${label} Report — The Pull Bar</h2>
        <p style="color:#666;font-size:12px;">Generated ${new Date().toLocaleString()}</p>
        <table style="border-collapse:collapse;width:100%;font-size:13px;">
          <thead><tr>${headers.map(h => `<th style="text-align:left;padding:6px 10px;border-bottom:2px solid #333;">${h.label}</th>`).join('')}</tr></thead>
          <tbody>${tableRows}</tbody>
        </table>
      </body></html>
    `);
    win.document.close();
    win.focus();
    setTimeout(() => win.print(), 300);
  };

  return (
    <div>
      <div style={{ display: 'flex', gap: 8, marginBottom: 18, flexWrap: 'wrap' }}>
        {REPORT_TYPES.map(r => (
          <button key={r.id} onClick={() => setReportType(r.id)} style={{
            padding: '8px 16px', borderRadius: 8, border: `1px solid ${reportType === r.id ? COLORS.brass : COLORS.brown}`,
            background: reportType === r.id ? COLORS.brass : 'transparent', color: reportType === r.id ? COLORS.espresso : COLORS.cream,
            fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 13, cursor: 'pointer',
          }}>{r.label}</button>
        ))}
      </div>

      {(reportType === 'sales' || reportType === 'topSellers') && (
        <div style={{ display: 'flex', gap: 10, marginBottom: 18, flexWrap: 'wrap' }}>
          <Field label="From"><input style={inputStyle} type="date" value={fromDate} onChange={e => setFromDate(e.target.value)} /></Field>
          <Field label="To"><input style={inputStyle} type="date" value={toDate} onChange={e => setToDate(e.target.value)} /></Field>
        </div>
      )}

      <div style={{ display: 'flex', gap: 10, marginBottom: 18, flexWrap: 'wrap' }}>
        <Button onClick={() => downloadCSV(`${filenameBase}.csv`, rows, headers)}><Download size={14} /> CSV</Button>
        <Button onClick={() => downloadXLSX(`${filenameBase}.xlsx`, rows, headers, label)}><Download size={14} /> Excel</Button>
        <Button variant="ghost" onClick={handlePrint}><FileText size={14} /> Print / Save as PDF</Button>
      </div>

      <div style={{ background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 12, overflow: 'auto' }}>
        {rows.length === 0 ? (
          <div style={{ padding: 30, textAlign: 'center', color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>No data for this report yet.</div>
        ) : (
          <table style={{ width: '100%', borderCollapse: 'collapse', fontFamily: 'Jost, sans-serif', fontSize: 13 }}>
            <thead>
              <tr>{headers.map(h => <th key={h.key} style={{ textAlign: 'left', padding: '10px 14px', borderBottom: `1px solid ${COLORS.brown}`, color: COLORS.creamDim, fontSize: 11, textTransform: 'uppercase', letterSpacing: 0.5 }}>{h.label}</th>)}</tr>
            </thead>
            <tbody>
              {rows.map((r, idx) => (
                <tr key={idx}>{headers.map(h => <td key={h.key} style={{ padding: '9px 14px', borderBottom: `1px solid ${COLORS.brown}`, color: COLORS.cream }}>{r[h.key]}</td>)}</tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ---------- Orders View (online order fulfillment) ----------
const FULFILLMENT_STATUSES = ['unfulfilled', 'packed', 'shipped', 'delivered'];
function OrdersView({ sales, onUpdate }) {
  const onlineOrders = sales.filter(s => s.channel === 'online').slice().reverse();
  return (
    <div>
      {onlineOrders.length === 0 ? (
        <div style={{ textAlign: 'center', padding: '60px 20px', color: COLORS.creamDim, fontFamily: 'Jost, sans-serif' }}>
          No online orders yet — check "Online order" at checkout in the POS to log one.
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {onlineOrders.map(s => (
            <div key={s.id} style={{ background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 10, padding: '14px 16px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 8 }}>
                <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, fontSize: 14 }}>{s.customerName || 'Guest'} · ${s.total.toFixed(2)}</div>
                <div style={{ fontFamily: 'Jost, sans-serif', color: COLORS.creamDim, fontSize: 12 }}>{new Date(s.date).toLocaleDateString()}</div>
              </div>
              <div style={{ fontFamily: 'Jost, sans-serif', color: COLORS.creamDim, fontSize: 12, marginBottom: 8 }}>{s.lines.map(l => `${l.name} ×${l.qty}`).join(', ')}</div>
              {s.shippingAddress && <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 11, color: COLORS.brassBright, whiteSpace: 'pre-wrap', marginBottom: 10 }}>{s.shippingAddress}</div>}
              <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
                <select
                  style={{ ...inputStyle, padding: '6px 10px', fontSize: 12 }}
                  value={s.fulfillmentStatus || 'unfulfilled'}
                  onChange={e => onUpdate({ ...s, fulfillmentStatus: e.target.value })}
                >
                  {FULFILLMENT_STATUSES.map(st => <option key={st} value={st}>{st}</option>)}
                </select>
                <input
                  style={{ ...inputStyle, padding: '6px 10px', fontSize: 12, flex: 1, minWidth: 160 }}
                  placeholder="Tracking number"
                  value={s.trackingNumber || ''}
                  onChange={e => onUpdate({ ...s, trackingNumber: e.target.value })}
                />
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- App ----------
export default function App() {
  const [tab, setTab] = useState('dashboard');
  const [items, setItems] = useState([]);
  const [sales, setSales] = useState([]);
  const [customers, setCustomers] = useState([]);
  const [coupons, setCoupons] = useState([]);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    (async () => {
      const [i, s, c, cp] = await Promise.all([loadKey(storageKeys.items), loadKey(storageKeys.sales), loadKey(storageKeys.customers), loadKey(storageKeys.coupons)]);
      setItems(i); setSales(s); setCustomers(c); setCoupons(cp); setLoaded(true);
    })();
  }, []);

  const handleAddItem = (item) => { const next = [...items, item]; setItems(next); saveKey(storageKeys.items, next); };
  const handleDeleteItem = (id) => { const next = items.filter(i => i.id !== id); setItems(next); saveKey(storageKeys.items, next); };
  const handleApplyMarkdown = (itemId, tierIndex, newPrice) => {
    const next = items.map(i => i.id === itemId ? { ...i, price: newPrice, appliedMarkdownIndex: tierIndex } : i);
    setItems(next); saveKey(storageKeys.items, next);
  };
  const handleAddCustomer = (c) => { const next = [...customers, c]; setCustomers(next); saveKey(storageKeys.customers, next); };
  const handleUpdateCustomer = (updated) => { const next = customers.map(c => c.id === updated.id ? updated : c); setCustomers(next); saveKey(storageKeys.customers, next); };
  const handleAddCoupon = (c) => { const next = [...coupons, c]; setCoupons(next); saveKey(storageKeys.coupons, next); };
  const handleDeleteCoupon = (id) => { const next = coupons.filter(c => c.id !== id); setCoupons(next); saveKey(storageKeys.coupons, next); };
  const handleUpdateSale = (updated) => { const next = sales.map(s => s.id === updated.id ? updated : s); setSales(next); saveKey(storageKeys.sales, next); };

  const handleCheckout = ({ cartLines, customer, couponCode, couponDiscount, pointsRedeemed, pointsDiscount, pointsEarned, total, isOnlineOrder, shippingAddress }) => {
    const sale = {
      id: uid(), date: new Date().toISOString(), total, couponCode: couponCode || null,
      couponDiscount, pointsRedeemed, pointsDiscount, pointsEarned,
      customerId: customer?.id || null, customerName: customer?.name || null,
      lines: cartLines.map(l => ({ name: l.item.name, qty: l.qty, price: l.item.price })),
      channel: isOnlineOrder ? 'online' : 'in-store',
      shippingAddress: shippingAddress || '',
      fulfillmentStatus: isOnlineOrder ? 'unfulfilled' : null,
      trackingNumber: '',
    };
    const nextSales = [...sales, sale];
    const nextItems = items.map(i => {
      const line = cartLines.find(l => l.itemId === i.id);
      return line ? { ...i, qty: Math.max(0, i.qty - line.qty) } : i;
    });
    let nextCustomers = customers;
    if (customer) {
      nextCustomers = customers.map(c => c.id === customer.id ? { ...c, points: Math.max(0, c.points - pointsRedeemed + pointsEarned) } : c);
    }
    setSales(nextSales); saveKey(storageKeys.sales, nextSales);
    setItems(nextItems); saveKey(storageKeys.items, nextItems);
    setCustomers(nextCustomers); saveKey(storageKeys.customers, nextCustomers);
  };

  const tabs = [
    { id: 'dashboard', label: 'Dashboard', icon: BarChart3 },
    { id: 'inventory', label: 'Inventory', icon: LayoutGrid },
    { id: 'pos', label: 'Point of sale', icon: ShoppingCart },
    { id: 'customers', label: 'Customers', icon: Users },
    { id: 'coupons', label: 'Coupons', icon: Tag },
    { id: 'reports', label: 'Reports', icon: FileText },
    { id: 'orders', label: 'Orders', icon: Truck },
  ];

  return (
    <div style={{ minHeight: '100vh', background: COLORS.paper, fontFamily: 'Jost, sans-serif' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,700&family=Jost:wght@400;500;600;700&family=JetBrains+Mono:wght@400;700&display=swap');
        * { box-sizing: border-box; }
        input:focus, select:focus { border-color: ${COLORS.brass} !important; }
      `}</style>
      <div style={{ maxWidth: 1150, margin: '0 auto', padding: '28px 20px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 24, flexWrap: 'wrap', gap: 14 }}>
          <div>
            <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontSize: 22, fontWeight: 700, letterSpacing: 1 }}>THE PULL <span style={{ color: COLORS.brass }}>BAR</span></div>
          </div>
          <div style={{ display: 'flex', gap: 6, background: COLORS.espressoSoft, padding: 5, borderRadius: 10, border: `1px solid ${COLORS.brown}`, flexWrap: 'wrap' }}>
            {tabs.map(t => (
              <button key={t.id} onClick={() => setTab(t.id)} style={{
                display: 'flex', alignItems: 'center', gap: 6, border: 'none', cursor: 'pointer',
                padding: '8px 14px', borderRadius: 7, fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 13,
                background: tab === t.id ? COLORS.brass : 'transparent', color: tab === t.id ? COLORS.espresso : COLORS.creamDim,
              }}><t.icon size={14} /> {t.label}</button>
            ))}
          </div>
        </div>

        {!loaded ? (
          <div style={{ color: COLORS.creamDim, textAlign: 'center', padding: 60 }}>Loading your inventory...</div>
        ) : (
          <>
            {tab === 'dashboard' && <Dashboard items={items} sales={sales} customers={customers} onApplyMarkdown={handleApplyMarkdown} />}
            {tab === 'inventory' && <InventoryView items={items} onAdd={handleAddItem} onDelete={handleDeleteItem} />}
            {tab === 'pos' && <POSView items={items} customers={customers} coupons={coupons} onCheckout={handleCheckout} onAddCustomer={handleAddCustomer} />}
            {tab === 'customers' && <CustomersView customers={customers} sales={sales} onAdd={handleAddCustomer} onUpdate={handleUpdateCustomer} />}
            {tab === 'coupons' && <CouponsView coupons={coupons} onAdd={handleAddCoupon} onDelete={handleDeleteCoupon} />}
            {tab === 'reports' && <ReportsView items={items} sales={sales} customers={customers} />}
            {tab === 'orders' && <OrdersView sales={sales} onUpdate={handleUpdateSale} />}
          </>
        )}
      </div>
    </div>
  );
}

