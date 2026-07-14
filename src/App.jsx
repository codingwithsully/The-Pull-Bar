import React, { useState, useEffect, useRef, useCallback, createContext, useContext } from 'react';
import { LayoutGrid, ShoppingCart, BarChart3, Plus, Search, Trash2, X, Package, TrendingUp, AlertTriangle, Minus, Check, Users, Tag, Star, Camera, FileText, Truck, Download, Globe } from 'lucide-react';
import * as XLSX from 'xlsx';

// ---------- Theme system: Light / Dark / High Contrast ----------
const THEMES = {
  light: {
    paper: '#FFFFFF',
    espresso: '#2E2117',
    espressoSoft: '#F8F8F8',
    brown: '#E0D5C7',
    brass: '#B58A4A',
    brassBright: '#D4A574',
    cream: '#2E2117',
    creamDim: '#8A7A63',
    oxblood: '#8B4038',
  },
  dark: {
    paper: '#1A1A1A',
    espresso: '#F1E9D8',
    espressoSoft: '#2E2E2E',
    brown: '#4A4440',
    brass: '#D4A574',
    brassBright: '#E8C68E',
    cream: '#F1E9D8',
    creamDim: '#B3A89F',
    oxblood: '#C9685F',
  },
  highContrast: {
    paper: '#000000',
    espresso: '#FFFFFF',
    espressoSoft: '#1A1A1A',
    brown: '#FFFFFF',
    brass: '#FFD700',
    brassBright: '#FFFF00',
    cream: '#FFFFFF',
    creamDim: '#CCCCCC',
    oxblood: '#FF6B6B',
  },
};

function getColors(theme = 'light') {
  return THEMES[theme] || THEMES.light;
}

const ColorContext = createContext();
const useColors = () => {
  const context = useContext(ColorContext);
  if (!context) throw new Error('useColors must be used within ColorProvider');
  return context;
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
const VARIANTS = [
  { code: 'CC', label: 'Common/Normal' },
  { code: 'H', label: 'Holo' },
  { code: 'RH', label: 'Reverse Holo' },
  { code: 'FA', label: 'Full Art' },
  { code: 'IR', label: 'Illustration Rare' },
  { code: 'SIR', label: 'Special Illustration Rare' },
  { code: 'UR', label: 'Ultra Rare' },
  { code: 'HR', label: 'Hyper Rare' },
  { code: 'SEC', label: 'Secret Rare' },
  { code: 'EX', label: 'EX' },
  { code: 'GX', label: 'GX' },
  { code: 'V', label: 'V' },
  { code: 'VMAX', label: 'V-MAX' },
  { code: 'VSTAR', label: 'V-Star' },
  { code: 'AS', label: 'Special Art' },
  { code: 'AR', label: 'Alternate Rare' },
  { code: 'RAD', label: 'Radiant Rare' },
  { code: 'TG', label: 'Tag Team GX' },
];

const DEFAULT_PLATFORMS = [
  { id: 'instore', name: 'In-Store', icon: 'store', platformFee: 0, promotionFee: 0 },
  { id: 'ebay', name: 'eBay', icon: 'ebay', platformFee: 2.15, promotionFee: 0 },
  { id: 'whatnot', name: 'WhatNot', icon: 'whatnot', platformFee: 0, promotionFee: 0 },
  { id: 'tcgplayer', name: 'TCG Player', icon: 'tcg', platformFee: 1.50, promotionFee: 0 },
  { id: 'tiktok', name: 'TikTok Shop', icon: 'tiktok', platformFee: 5, promotionFee: 0 },
  { id: 'mercari', name: 'Mercari', icon: 'mercari', platformFee: 1, promotionFee: 0 },
];

function genSku({ type, game, set, number, variant, condition, language }) {
  if (type === 'accessory') {
    return `ACC-${(game || 'GEN').toUpperCase().slice(0, 3)}-${uid().slice(0, 3).toUpperCase()}`;
  }
  const parts = [];
  if (language) parts.push(language);
  parts.push((game || 'PKM').toUpperCase());
  if (set) parts.push(set.replace(/[^A-Za-z0-9]/g, '').slice(0, 3).toUpperCase());
  if (number) {
    const numStr = number.replace(/[^0-9]/g, '');
    const padded = numStr.padStart(3, '0');
    parts.push(padded);
  }
  const variantToUse = variant || 'CC';
  parts.push(variantToUse);
  if (condition && condition !== 'NM') parts.push(condition);
  return parts.join('-');
}

const POINTS_PER_DOLLAR = 1;
const POINT_REDEMPTION_RATE = 0.05;

const storageKeys = { items: 'tcg-inventory-items', sales: 'tcg-sales-log', customers: 'tcg-customers', coupons: 'tcg-coupons', platforms: 'tcg-platforms' };
async function loadKey(key) {
  try { const raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : []; } catch { return []; }
}
async function saveKey(key, val) {
  try { localStorage.setItem(key, JSON.stringify(val)); } catch {}
}

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

function printReceipt(sale, storeInfo = {}) {
  const win = window.open('', '_blank');
  const itemsHtml = sale.lines.map(l => `<tr><td style="padding:4px;border-bottom:1px solid #ddd;text-align:left;">${l.sku || 'N/A'}</td><td style="padding:4px;border-bottom:1px solid #ddd;">${l.name}</td><td style="padding:4px;border-bottom:1px solid #ddd;text-align:center;">×${l.qty}</td><td style="padding:4px;border-bottom:1px solid #ddd;text-align:right;">$${(l.price * l.qty).toFixed(2)}</td></tr>`).join('');
  win.document.write(`<html><head><title>Receipt ${sale.receiptNumber}</title></head><body style="font-family:'Courier New',monospace;max-width:400px;margin:0 auto;padding:20px;"><div style="text-align:center;margin-bottom:20px;"><div style="font-size:18px;font-weight:bold;">THE PULL BAR</div><div style="font-size:11px;color:#666;">Receipt ${sale.receiptNumber}</div><div style="font-size:11px;color:#666;">${new Date(sale.date).toLocaleString()}</div></div><table style="width:100%;font-size:12px;"><thead><tr style="border-bottom:2px solid #000;"><th style="text-align:left;padding:4px;">SKU</th><th style="text-align:left;padding:4px;">Item</th><th style="text-align:center;padding:4px;">Qty</th><th style="text-align:right;padding:4px;">Total</th></tr></thead><tbody>${itemsHtml}</tbody></table><div style="margin-top:16px;font-size:13px;font-weight:bold;text-align:right;border-top:2px solid #000;padding-top:8px;">Total: $${sale.total.toFixed(2)}</div>${sale.couponCode ? `<div style="text-align:center;font-size:11px;margin-top:8px;color:#666;">Coupon: ${sale.couponCode} (-$${sale.couponDiscount.toFixed(2)})</div>` : ''}${sale.pointsRedeemed > 0 ? `<div style="text-align:center;font-size:11px;color:#666;">Points redeemed: ${sale.pointsRedeemed} (-$${sale.pointsDiscount.toFixed(2)})</div>` : ''}<div style="text-align:center;font-size:11px;margin-top:16px;color:#666;border-top:1px solid #ddd;padding-top:8px;">Thank you for your business!</div></body></html>`);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 300);
}

function printPackingSlip(sale, storeInfo = {}) {
  const win = window.open('', '_blank');
  const itemsHtml = sale.lines.map(l => `<tr><td style="padding:8px;border-bottom:1px solid #ddd;font-weight:bold;">${l.sku}</td><td style="padding:8px;border-bottom:1px solid #ddd;">${l.name}</td><td style="padding:8px;border-bottom:1px solid #ddd;text-align:center;">×${l.qty}</td></tr>`).join('');
  win.document.write(`<html><head><title>Packing Slip ${sale.receiptNumber}</title></head><body style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;padding:40px;line-height:1.6;"><div style="display:flex;justify-content:space-between;margin-bottom:30px;"><div><div style="font-size:20px;font-weight:bold;">THE PULL BAR</div><div style="font-size:12px;color:#666;">${storeInfo.address || 'Store Address'}</div><div style="font-size:12px;color:#666;">${storeInfo.phone || 'Phone'}</div></div><div style="text-align:right;"><div style="font-size:14px;font-weight:bold;">PACKING SLIP</div><div style="font-size:12px;">Receipt #: ${sale.receiptNumber}</div><div style="font-size:12px;">Date: ${new Date(sale.date).toLocaleDateString()}</div></div></div><div style="margin-bottom:30px;padding:15px;background:#f5f5f5;border-radius:5px;"><div style="font-weight:bold;margin-bottom:8px;">SHIP TO:</div><div style="font-size:14px;white-space:pre-wrap;">${sale.shippingAddress}</div></div><table style="width:100%;margin-bottom:30px;font-size:13px;"><thead><tr style="border-bottom:2px solid #000;"><th style="text-align:left;padding:8px;">SKU</th><th style="text-align:left;padding:8px;">Item Description</th><th style="text-align:center;padding:8px;">Quantity</th></tr></thead><tbody>${itemsHtml}</tbody></table><div style="border-top:1px solid #ddd;padding-top:20px;font-size:11px;color:#666;line-height:1.8;"><strong>Return Policy:</strong> Items may be returned within 7 days of receipt in original condition for full refund. Clearance and final-sale items are non-returnable.</div></body></html>`);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 300);
}

// ---------- UI Atoms ----------
function StatCard({ label, value, icon: Icon, accent, onClick }) {
  const COLORS = useColors();
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
  const COLORS = useColors();
  return <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontFamily: 'Jost, sans-serif', fontSize: 12, color: COLORS.creamDim }}>{label}{children}</label>;
}

function useInputStyle() {
  const COLORS = useColors();
  return { background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 8, padding: '9px 10px', color: COLORS.cream, fontFamily: 'Jost, sans-serif', fontSize: 14, outline: 'none' };
}

function Button({ children, onClick, variant = 'primary', style, ...rest }) {
  const COLORS = useColors();
  const base = { border: 'none', borderRadius: 8, padding: '10px 16px', fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 13, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6 };
  const variants = {
    primary: { background: COLORS.brass, color: COLORS.espresso },
    ghost: { background: 'transparent', color: COLORS.cream, border: `1px solid ${COLORS.brown}` },
    danger: { background: 'transparent', color: COLORS.oxblood, border: `1px solid ${COLORS.oxblood}` },
  };
  return <button onClick={onClick} style={{ ...base, ...variants[variant], ...style }} {...rest}>{children}</button>;
}

function Modal({ title, onClose, children, width = 420 }) {
  const COLORS = useColors();
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

// ---------- Modals ----------
function AddItemModal({ onClose, onSave, existingItems = [] }) {
  const COLORS = useColors();
  const inputStyle = useInputStyle();
  const [type, setType] = useState('card');
  const [language, setLanguage] = useState('');
  const [game, setGame] = useState('PKM');
  const [name, setName] = useState('');
  const [nameQuery, setNameQuery] = useState('');
  const [showNameSuggestions, setShowNameSuggestions] = useState(false);
  const [set, setSet] = useState('');
  const [number, setNumber] = useState('');
  const [variant, setVariant] = useState('');
  const [variantQuery, setVariantQuery] = useState('');
  const [showVariantDropdown, setShowVariantDropdown] = useState(false);
  const filteredVariants = VARIANTS.filter(v => v.code.toLowerCase().includes(variantQuery.toLowerCase()) || v.label.toLowerCase().includes(variantQuery.toLowerCase()));
  const [condition, setCondition] = useState('NM');
  const [cost, setCost] = useState('');
  const [price, setPrice] = useState('');
  const [qty, setQty] = useState('1');
  const [hold, setHold] = useState(false);
  const [schedule, setSchedule] = useState([]);
  const [scheduleDays, setScheduleDays] = useState('');
  const [schedulePercent, setSchedulePercent] = useState('');
  const [barcodeInput, setBarcodeInput] = useState('');
  const [barcodeError, setBarcodeError] = useState('');

  // Name suggestions from existing items
  const uniqueNames = [...new Set(existingItems.map(i => i.name))].filter(n => n && n.toLowerCase().includes(nameQuery.toLowerCase()));
  
  const handleSelectName = (selectedName) => {
    setName(selectedName);
    setNameQuery('');
    setShowNameSuggestions(false);
  };

  const handleBarcodeScanned = (barcode) => {
    const foundItem = existingItems.find(i => i.sku.toLowerCase() === barcode.toLowerCase());
    if (foundItem) {
      // Pre-fill form with existing item data
      setType(foundItem.type);
      setLanguage(foundItem.language);
      setGame(foundItem.game);
      setName(foundItem.name);
      setNameQuery('');
      setSet(foundItem.set);
      setNumber(foundItem.number);
      setVariant(foundItem.variant);
      setVariantQuery(foundItem.variant);
      setCondition(foundItem.condition || 'NM');
      setCost(foundItem.cost.toString());
      setPrice(foundItem.price.toString());
      setQty('1'); // Default to adding 1 unit
      setBarcodeInput('');
      setBarcodeError('');
      alert(`Found: ${foundItem.name}\nAdjust quantity and save to add more units.`);
    } else {
      setBarcodeError(`Barcode not found: ${barcode}. Add as new item.`);
      setTimeout(() => setBarcodeError(''), 3000);
    }
  };

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
    <Modal title="Add item to inventory" onClose={onClose} width={480}>
      <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
        {['card', 'accessory'].map(t => (
          <button key={t} onClick={() => setType(t)} style={{ flex: 1, padding: '8px', borderRadius: 8, border: `1px solid ${type === t ? COLORS.brass : COLORS.brown}`, background: type === t ? COLORS.brass : 'transparent', color: type === t ? COLORS.espresso : COLORS.cream, fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 13, cursor: 'pointer', textTransform: 'capitalize' }}>{t}</button>
        ))}
      </div>

      {/* Barcode lookup section */}
      <Field label="Scan existing item (optional)">
        <div style={{ display: 'flex', gap: 8 }}>
          <input style={inputStyle} type="text" placeholder="Barcode..." value={barcodeInput} onChange={e => setBarcodeInput(e.target.value)} onKeyDown={e => e.key === 'Enter' && barcodeInput && handleBarcodeScanned(barcodeInput)} />
          <Button onClick={() => barcodeInput && handleBarcodeScanned(barcodeInput)} style={{ padding: '8px 12px', fontSize: 12 }}>Lookup</Button>
        </div>
        {barcodeError && <div style={{ color: COLORS.oxblood, fontFamily: 'Jost, sans-serif', fontSize: 11, marginTop: 6 }}>{barcodeError}</div>}
      </Field>

      {/* Character/Name field with autocomplete */}
      <Field label="Name">
        <div style={{ position: 'relative' }}>
          <input 
            style={inputStyle} 
            value={name} 
            onChange={e => { setName(e.target.value); setNameQuery(e.target.value); setShowNameSuggestions(!!e.target.value); }}
            onFocus={() => setShowNameSuggestions(!!nameQuery)}
            onBlur={() => setTimeout(() => setShowNameSuggestions(false), 200)}
            placeholder={type === 'card' ? 'Charizard ex' : 'Ultra Pro Sleeves'} 
            autoFocus 
          />
          {showNameSuggestions && uniqueNames.length > 0 && (
            <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderTop: 'none', borderRadius: '0 0 8px 8px', zIndex: 1000, maxHeight: 150, overflowY: 'auto' }}>
              {uniqueNames.slice(0, 5).map(n => (
                <div key={n} onClick={() => handleSelectName(n)} style={{ padding: '8px 12px', cursor: 'pointer', borderBottom: `1px solid ${COLORS.brown}`, color: COLORS.cream, fontSize: 12, fontFamily: 'Jost, sans-serif' }}>
                  {n}
                </div>
              ))}
            </div>
          )}
        </div>
      </Field>

      {type === 'card' && (
        <>
          <Field label="Game"><input style={inputStyle} value={game} onChange={e => setGame(e.target.value)} placeholder="PKM / MTG / YGO" /></Field>
          <Field label="Language"><select style={inputStyle} value={language} onChange={e => setLanguage(e.target.value)}>{LANGUAGES.map(l => <option key={l.code} value={l.code}>{l.label}</option>)}</select></Field>
          <Field label="Set"><input style={inputStyle} value={set} onChange={e => setSet(e.target.value)} placeholder="OBF" /></Field>
          <Field label="Card #"><input style={inputStyle} value={number} onChange={e => setNumber(e.target.value)} placeholder="054" /></Field>
          
          {/* Variant field with both search and dropdown */}
          <Field label="Variant / rarity">
            <div style={{ position: 'relative' }}>
              <input 
                style={inputStyle} 
                type="text" 
                placeholder="Search or select variant..." 
                value={variantQuery} 
                onChange={e => { setVariantQuery(e.target.value); setShowVariantDropdown(true); }}
                onFocus={() => setShowVariantDropdown(true)}
                onBlur={() => setTimeout(() => setShowVariantDropdown(false), 200)}
              />
              {showVariantDropdown && (
                <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderTop: 'none', borderRadius: '0 0 8px 8px', zIndex: 1000, maxHeight: 250, overflowY: 'auto' }}>
                  {filteredVariants.length > 0 ? (
                    <>
                      {variantQuery && <div style={{ padding: '8px 12px', fontFamily: 'Jost, sans-serif', fontSize: 10, color: COLORS.creamDim, fontWeight: 600, textTransform: 'uppercase' }}>Matching:</div>}
                      {filteredVariants.map(v => (
                        <div key={v.code} onClick={() => { setVariant(v.code); setVariantQuery(v.label); setShowVariantDropdown(false); }} style={{ padding: '8px 12px', cursor: 'pointer', borderBottom: `1px solid ${COLORS.brown}`, color: COLORS.cream, fontSize: 13, fontFamily: 'Jost, sans-serif', background: variant === v.code ? COLORS.brown : 'transparent' }}>
                          <span style={{ fontWeight: 600, color: COLORS.brass }}>{v.code}</span> — {v.label}
                        </div>
                      ))}
                    </>
                  ) : variantQuery ? (
                    <div style={{ padding: '8px 12px', color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 12 }}>No matches. Type custom variant or select from all below.</div>
                  ) : null}
                  
                  {/* Show all variants if no search term or when dropdown is open */}
                  {!variantQuery && (
                    <>
                      <div style={{ padding: '8px 12px', fontFamily: 'Jost, sans-serif', fontSize: 10, color: COLORS.creamDim, fontWeight: 600, textTransform: 'uppercase', background: COLORS.paper }}>All variants:</div>
                      {VARIANTS.map(v => (
                        <div key={v.code} onClick={() => { setVariant(v.code); setVariantQuery(v.label); setShowVariantDropdown(false); }} style={{ padding: '8px 12px', cursor: 'pointer', borderBottom: `1px solid ${COLORS.brown}`, color: COLORS.cream, fontSize: 12, fontFamily: 'Jost, sans-serif', background: variant === v.code ? COLORS.brown : 'transparent' }}>
                          <span style={{ fontWeight: 600, color: COLORS.brass }}>{v.code}</span> — {v.label}
                        </div>
                      ))}
                    </>
                  )}
                </div>
              )}
            </div>
          </Field>

          <Field label="Condition"><select style={inputStyle} value={condition} onChange={e => setCondition(e.target.value)}>{CONDITIONS.map(c => <option key={c} value={c}>{c} — {CONDITION_LABELS[c]}</option>)}</select></Field>
        </>
      )}
      <Field label="Cost $"><input style={inputStyle} type="number" value={cost} onChange={e => setCost(e.target.value)} placeholder="0.00" /></Field>
      <Field label="Price $"><input style={inputStyle} type="number" value={price} onChange={e => setPrice(e.target.value)} placeholder="0.00" /></Field>
      <Field label="Quantity"><input style={inputStyle} type="number" value={qty} onChange={e => setQty(e.target.value)} placeholder="1" /></Field>
      <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: 'Jost, sans-serif', fontSize: 13, color: COLORS.cream, marginBottom: 12 }}>
        <input type="checkbox" checked={hold} onChange={e => setHold(e.target.checked)} />
        Hold — collectible, never mark down automatically
      </label>
      {!hold && (
        <div style={{ marginBottom: 16 }}>
          <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, marginBottom: 8 }}>Markdown schedule (optional)</div>
          {schedule.map((s, idx) => (
            <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6, padding: '8px', background: COLORS.espressoSoft, borderRadius: 6 }}>
              <span style={{ fontFamily: 'Jost, sans-serif', fontSize: 12, color: COLORS.cream }}>After {s.days} days → {s.percent}% off</span>
              <button onClick={() => removeScheduleRow(idx)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.oxblood }}><Trash2 size={14} /></button>
            </div>
          ))}
          <div style={{ display: 'flex', gap: 8, marginTop: 10 }}>
            <input style={inputStyle} type="number" value={scheduleDays} onChange={e => setScheduleDays(e.target.value)} placeholder="Days" />
            <input style={inputStyle} type="number" value={schedulePercent} onChange={e => setSchedulePercent(e.target.value)} placeholder="%" />
            <Button onClick={addScheduleRow} style={{ padding: '8px 12px', fontSize: 12 }}><Plus size={12} /> Add</Button>
          </div>
        </div>
      )}
      <div style={{ display: 'flex', gap: 10, marginTop: 18 }}>
        <Button variant="ghost" onClick={onClose} style={{ flex: 1, justifyContent: 'center' }}>Cancel</Button>
        <Button onClick={handleSave} style={{ flex: 1, justifyContent: 'center', opacity: canSave ? 1 : 0.5 }}>Add to inventory</Button>
      </div>
    </Modal>
  );
}

function AddCustomerModal({ onClose, onSave }) {
  const inputStyle = useInputStyle();
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [email, setEmail] = useState('');
  const canSave = name.trim() && (phone.trim() || email.trim());
  return (
    <Modal title="Add customer" onClose={onClose}>
      <Field label="Name"><input style={inputStyle} value={name} onChange={e => setName(e.target.value)} placeholder="Jordan Lee" autoFocus /></Field>
      <Field label="Phone"><input style={inputStyle} value={phone} onChange={e => setPhone(e.target.value)} placeholder="(555) 123-4567" /></Field>
      <Field label="Email"><input style={inputStyle} value={email} onChange={e => setEmail(e.target.value)} placeholder="jordan@email.com" /></Field>
      <div style={{ display: 'flex', gap: 10, marginTop: 18 }}>
        <Button variant="ghost" onClick={onClose} style={{ flex: 1, justifyContent: 'center' }}>Cancel</Button>
        <Button onClick={() => canSave && onSave({ id: uid(), name: name.trim(), phone: phone.trim(), email: email.trim(), points: 0, wishlist: [] })} style={{ flex: 1, justifyContent: 'center', opacity: canSave ? 1 : 0.5 }}>Add customer</Button>
      </div>
    </Modal>
  );
}

function AddCouponModal({ onClose, onSave }) {
  const COLORS = useColors();
  const inputStyle = useInputStyle();
  const [code, setCode] = useState('');
  const [type, setType] = useState('percent');
  const [value, setValue] = useState('');
  const canSave = code.trim() && value !== '';
  return (
    <Modal title="Add coupon" onClose={onClose}>
      <Field label="Code"><input style={inputStyle} value={code} onChange={e => setCode(e.target.value.toUpperCase())} placeholder="WELCOME10" autoFocus /></Field>
      <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
        {['percent', 'fixed'].map(t => (
          <button key={t} onClick={() => setType(t)} style={{ flex: 1, padding: '8px', borderRadius: 8, border: `1px solid ${type === t ? COLORS.brass : COLORS.brown}`, background: type === t ? COLORS.brass : 'transparent', color: type === t ? COLORS.espresso : COLORS.cream, fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 13, cursor: 'pointer' }}>
            {type === 'percent' ? '% off' : '$ off'}
          </button>
        ))}
      </div>
      <Field label="Amount"><input style={inputStyle} type="number" value={value} onChange={e => setValue(e.target.value)} placeholder={type === 'percent' ? '10' : '5.00'} /></Field>
      <div style={{ display: 'flex', gap: 10, marginTop: 18 }}>
        <Button variant="ghost" onClick={onClose} style={{ flex: 1, justifyContent: 'center' }}>Cancel</Button>
        <Button onClick={() => canSave && onSave({ id: uid(), code: code.trim(), type, value: parseFloat(value) || 0, active: true })} style={{ flex: 1, justifyContent: 'center', opacity: canSave ? 1 : 0.5 }}>Add coupon</Button>
      </div>
    </Modal>
  );
}

function QuickAddCustomerModal({ onClose, onSave }) {
  const inputStyle = useInputStyle();
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [email, setEmail] = useState('');
  const canSave = name.trim() && (phone.trim() || email.trim());
  return (
    <Modal title="Add customer" onClose={onClose}>
      <Field label="Name"><input style={inputStyle} value={name} onChange={e => setName(e.target.value)} placeholder="Jordan Lee" autoFocus /></Field>
      <Field label="Phone"><input style={inputStyle} value={phone} onChange={e => setPhone(e.target.value)} placeholder="(555) 123-4567" /></Field>
      <Field label="Email"><input style={inputStyle} value={email} onChange={e => setEmail(e.target.value)} placeholder="jordan@email.com" /></Field>
      <div style={{ display: 'flex', gap: 10, marginTop: 18 }}>
        <Button variant="ghost" onClick={onClose} style={{ flex: 1, justifyContent: 'center' }}>Cancel</Button>
        <Button onClick={() => canSave && onSave({ id: uid(), name: name.trim(), phone: phone.trim(), email: email.trim(), points: 0, wishlist: [] })} style={{ flex: 1, justifyContent: 'center', opacity: canSave ? 1 : 0.5 }}>Add & select</Button>
      </div>
    </Modal>
  );
}

function CameraScanModal({ onDetect, onClose }) {
  const COLORS = useColors();
  const videoTargetRef = useRef(null);
  const [status, setStatus] = useState('loading');
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
            setErrorMsg(err.name === 'NotAllowedError' ? 'Camera access denied. Check permissions.' : 'Could not start camera.');
            return;
          }
          window.Quagga.start();
          window.Quagga.onDetected(handleDetected);
          setStatus('scanning');
        });
      })
      .catch(() => { if (!cancelled) { setStatus('error'); setErrorMsg('Scanner failed to load.'); } });

    return () => {
      cancelled = true;
      try { window.Quagga && window.Quagga.offDetected && window.Quagga.offDetected(handleDetected); window.Quagga && window.Quagga.stop(); } catch (e) {}
    };
  }, []);

  return (
    <Modal title="Scan barcode" onClose={onClose} width={500}>
      <div ref={videoTargetRef} style={{ width: '100%', height: 300, background: COLORS.paper, borderRadius: 8, marginBottom: 16 }} />
      {status === 'loading' && <div style={{ textAlign: 'center', color: COLORS.creamDim, fontFamily: 'Jost, sans-serif' }}>Starting camera...</div>}
      {status === 'error' && <div style={{ textAlign: 'center', color: COLORS.oxblood, fontFamily: 'Jost, sans-serif' }}>{errorMsg}</div>}
      {status === 'scanning' && <div style={{ textAlign: 'center', color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 12 }}>Point camera at barcode — it will add automatically.</div>}
      {lastScanned && <div style={{ textAlign: 'center', color: COLORS.brass, fontFamily: 'JetBrains Mono, monospace', fontSize: 12, marginTop: 10 }}>Last: {lastScanned}</div>}
      <div style={{ display: 'flex', gap: 10, marginTop: 18 }}>
        <Button onClick={onClose} style={{ flex: 1, justifyContent: 'center' }}>Done</Button>
      </div>
    </Modal>
  );
}

// ---------- Views ----------
function InventoryView({ items, onAdd, onDelete }) {
  const COLORS = useColors();
  const [query, setQuery] = useState('');
  const [showAdd, setShowAdd] = useState(false);
  const inputStyle = useInputStyle();
  const filtered = items.filter(i => i.name.toLowerCase().includes(query.toLowerCase()) || i.sku.toLowerCase().includes(query.toLowerCase()) || (i.set || '').toLowerCase().includes(query.toLowerCase()));
  return (
    <div>
      <div style={{ display: 'flex', gap: 10, marginBottom: 18 }}>
        <input style={{ ...inputStyle, flex: 1 }} type="text" placeholder="Search inventory..." value={query} onChange={e => setQuery(e.target.value)} />
        <Button onClick={() => setShowAdd(true)}><Plus size={14} /> Add item</Button>
      </div>
      {filtered.length === 0 ? (
        <div style={{ textAlign: 'center', padding: 40, color: COLORS.creamDim, fontFamily: 'Jost, sans-serif' }}>{items.length === 0 ? 'No inventory yet.' : 'No matches.'}</div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {filtered.map(item => (
            <div key={item.id} style={{ background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 10, padding: '12px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <div style={{ color: COLORS.brass, fontFamily: 'JetBrains Mono, monospace', fontSize: 11, fontWeight: 600 }}>{item.sku}</div>
                <div style={{ color: COLORS.cream, fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 14, marginBottom: 4 }}>{item.name}</div>
                {item.set && <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 11 }}>{item.set} #{item.number}</div>}
              </div>
              <div style={{ textAlign: 'right' }}>
                <div style={{ color: COLORS.brass, fontFamily: 'Fraunces, serif', fontSize: 16, fontWeight: 700 }}>${item.price.toFixed(2)}</div>
                <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 11 }}>cost ${item.cost.toFixed(2)} · qty {item.qty}</div>
              </div>
              <button onClick={() => onDelete(item.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.oxblood }}><Trash2 size={16} /></button>
            </div>
          ))}
        </div>
      )}
      {showAdd && <AddItemModal existingItems={items} onClose={() => setShowAdd(false)} onSave={(item) => { onAdd(item); setShowAdd(false); }} />}
    </div>
  );
}

const MS_DAY = 1000 * 60 * 60 * 24;

function Dashboard({ items, sales, customers, onApplyMarkdown }) {
  const COLORS = useColors();
  const [detail, setDetail] = useState(null);

  const inventoryValue = items.reduce((sum, i) => sum + i.price * i.qty, 0);
  const inventoryCost = items.reduce((sum, i) => sum + i.cost * i.qty, 0);
  const totalRevenue = sales.reduce((sum, s) => sum + s.total, 0);
  const totalCOGS = sales.reduce((sum, s) => sum + (s.costOfGoods || 0), 0);
  const totalShipping = sales.reduce((sum, s) => sum + (s.shippingPaidBy === 'me' ? s.shippingCost : 0), 0);
  const totalPlatformFees = sales.reduce((sum, s) => sum + (s.platformFees || 0), 0);
  const totalPromoCosts = sales.reduce((sum, s) => sum + (s.promotionCosts || 0), 0);
  const netProfit = totalRevenue - totalCOGS - totalShipping - totalPlatformFees - totalPromoCosts;
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

  if (detail) {
    return (
      <div>
        <button onClick={() => setDetail(null)} style={{ marginBottom: 18, padding: '8px 14px', background: COLORS.brass, color: COLORS.espresso, border: 'none', borderRadius: 8, cursor: 'pointer', fontFamily: 'Jost, sans-serif', fontWeight: 600 }}>← Back to Dashboard</button>
        <h2 style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontSize: 20, marginBottom: 16 }}>
          {detail === 'value' && 'Inventory Value Breakdown'}
          {detail === 'cost' && 'Inventory Cost Breakdown'}
          {detail === 'revenue' && 'All Sales'}
          {detail === 'customers' && 'All Customers'}
          {detail === 'lowStock' && 'Low Stock Items'}
        </h2>
        <div style={{ background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 10, overflow: 'hidden' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontFamily: 'Jost, sans-serif', fontSize: 13 }}>
            <tbody>
              {detail === 'value' && items.slice().sort((a, b) => (b.price * b.qty) - (a.price * a.qty)).map(i => (
                <tr key={i.id} style={{ borderBottom: `1px solid ${COLORS.brown}` }}>
                  <td style={{ padding: '10px 14px', color: COLORS.cream }}>{i.name} · qty {i.qty}</td>
                  <td style={{ padding: '10px 14px', textAlign: 'right', color: COLORS.brass, fontWeight: 600 }}>${(i.price * i.qty).toFixed(2)}</td>
                </tr>
              ))}
              {detail === 'cost' && items.slice().sort((a, b) => (b.cost * b.qty) - (a.cost * a.qty)).map(i => (
                <tr key={i.id} style={{ borderBottom: `1px solid ${COLORS.brown}` }}>
                  <td style={{ padding: '10px 14px', color: COLORS.cream }}>{i.name} · qty {i.qty}</td>
                  <td style={{ padding: '10px 14px', textAlign: 'right', color: COLORS.brass, fontWeight: 600 }}>${(i.cost * i.qty).toFixed(2)}</td>
                </tr>
              ))}
              {detail === 'revenue' && sales.slice().reverse().map(s => (
                <tr key={s.id} style={{ borderBottom: `1px solid ${COLORS.brown}` }}>
                  <td style={{ padding: '10px 14px', color: COLORS.cream }}>{new Date(s.date).toLocaleDateString()} · {s.lines.map(l => `${l.name} ×${l.qty}`).join(', ')}{s.customerName ? ` · ${s.customerName}` : ''}</td>
                  <td style={{ padding: '10px 14px', textAlign: 'right', color: COLORS.brass, fontWeight: 600 }}>${s.total.toFixed(2)}</td>
                </tr>
              ))}
              {detail === 'customers' && customers.map(c => (
                <tr key={c.id} style={{ borderBottom: `1px solid ${COLORS.brown}` }}>
                  <td style={{ padding: '10px 14px', color: COLORS.cream }}>{c.name} · {[c.phone, c.email].filter(Boolean).join(' · ')}</td>
                  <td style={{ padding: '10px 14px', textAlign: 'right', color: COLORS.brass, fontWeight: 600 }}>{c.points} pts</td>
                </tr>
              ))}
              {detail === 'lowStock' && lowStock.map(i => (
                <tr key={i.id} style={{ borderBottom: `1px solid ${COLORS.brown}` }}>
                  <td style={{ padding: '10px 14px', color: COLORS.cream }}>{i.name}</td>
                  <td style={{ padding: '10px 14px', textAlign: 'right', color: COLORS.oxblood, fontWeight: 600 }}>qty {i.qty}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div style={{ display: 'flex', gap: 14, flexWrap: 'wrap', marginBottom: 24 }}>
        <StatCard label="Inventory value" value={`$${inventoryValue.toFixed(2)}`} icon={Package} onClick={() => setDetail('value')} />
        <StatCard label="Inventory cost" value={`$${inventoryCost.toFixed(2)}`} icon={TrendingUp} onClick={() => setDetail('cost')} />
        <StatCard label="Total revenue" value={`$${totalRevenue.toFixed(2)}`} icon={BarChart3} onClick={() => setDetail('revenue')} />
        <StatCard label="Net profit" value={`$${netProfit.toFixed(2)}`} icon={TrendingUp} accent={netProfit >= 0 ? COLORS.brass : COLORS.oxblood} />
        <StatCard label="Customers" value={customers.length} icon={Users} onClick={() => setDetail('customers')} />
        <StatCard label="Low stock" value={lowStock.length} icon={AlertTriangle} accent={COLORS.oxblood} onClick={() => setDetail('lowStock')} />
      </div>

      <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, marginBottom: 10, fontSize: 15 }}>Top sellers</div>
      {topSellers.length === 0 ? (
        <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>No sales yet.</div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {topSellers.map((t, idx) => (
            <div key={t.name} style={{ display: 'flex', justifyContent: 'space-between', background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 8, padding: '10px 14px', fontFamily: 'Jost, sans-serif', fontSize: 13 }}>
              <span style={{ color: COLORS.cream }}><span style={{ color: COLORS.brass, fontFamily: 'JetBrains Mono, monospace', marginRight: 8 }}>#{idx + 1}</span>{t.name} · {t.qty} sold</span>
              <span style={{ color: COLORS.brass, fontWeight: 600 }}>${t.revenue.toFixed(2)}</span>
            </div>
          ))}
        </div>
      )}

      <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, marginBottom: 10, fontSize: 15, marginTop: 20 }}>Recent sales</div>
      {sales.length === 0 ? (
        <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>No sales logged yet.</div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {sales.slice().reverse().slice(0, 8).map(s => (
            <div key={s.id} style={{ display: 'flex', justifyContent: 'space-between', background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 8, padding: '10px 14px', fontFamily: 'Jost, sans-serif', fontSize: 13 }}>
              <span style={{ color: COLORS.cream }}>{s.receiptNumber} · {s.lines.map(l => `${l.name} ×${l.qty}`).join(', ')}{s.customerName ? ` · ${s.customerName}` : ''}</span>
              <span style={{ color: COLORS.brass, fontWeight: 600 }}>${s.total.toFixed(2)}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function CustomersView({ customers, sales, onAdd, onUpdate }) {
  const COLORS = useColors();
  const inputStyle = useInputStyle();
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
      <div style={{ display: 'flex', gap: 10, marginBottom: 18 }}>
        <input style={{ ...inputStyle, flex: 1 }} type="text" placeholder="Search customers..." value={query} onChange={e => setQuery(e.target.value)} />
        <Button onClick={() => setShowAdd(true)}><Plus size={14} /> Add customer</Button>
      </div>
      {filtered.length === 0 ? (
        <div style={{ textAlign: 'center', padding: 40, color: COLORS.creamDim, fontFamily: 'Jost, sans-serif' }}>No customers yet.</div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {filtered.map(c => (
            <div key={c.id}>
              <div onClick={() => setSelected(selected?.id === c.id ? null : c)} style={{ background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 10, padding: '12px 16px', cursor: 'pointer' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <div style={{ color: COLORS.cream, fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 14 }}>{c.name}</div>
                    <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 12 }}>{[c.phone, c.email].filter(Boolean).join(' · ')}</div>
                  </div>
                  <div style={{ color: COLORS.brass, fontFamily: 'Fraunces, serif', fontWeight: 700, fontSize: 16 }}>{c.points} pts</div>
                </div>
              </div>
              {selected?.id === c.id && (
                <div style={{ background: COLORS.paper, border: `1px solid ${COLORS.brown}`, borderTop: 'none', borderRadius: '0 0 10px 10px', padding: '12px 16px' }}>
                  <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, fontSize: 13, marginBottom: 8 }}>Purchase history</div>
                  {customerSales.length === 0 ? (
                    <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 12, marginBottom: 12 }}>No purchases yet.</div>
                  ) : (
                    <div style={{ marginBottom: 12 }}>
                      {customerSales.slice().reverse().map(s => (
                        <div key={s.id} style={{ padding: '6px 0', borderBottom: `1px solid ${COLORS.brown}`, color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 12 }}>
                          {s.lines.map(l => `${l.name} ×${l.qty}`).join(', ')} — ${s.total.toFixed(2)}
                        </div>
                      ))}
                    </div>
                  )}
                  <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, fontSize: 13, marginBottom: 8 }}>Wishlist</div>
                  {(c.wishlist || []).length === 0 && <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 12, marginBottom: 8 }}>Nothing on their list.</div>}
                  {(c.wishlist || []).map((w, idx) => (
                    <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', padding: '4px 0', marginBottom: 4 }}>
                      <span style={{ color: COLORS.cream, fontFamily: 'Jost, sans-serif', fontSize: 12 }}>{w}</span>
                      <button onClick={() => removeWishlistItem(c, idx)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.oxblood }}><X size={14} /></button>
                    </div>
                  ))}
                  <div style={{ display: 'flex', gap: 6, marginTop: 8 }}>
                    <input style={{ ...inputStyle, flex: 1, fontSize: 12, padding: '6px 8px' }} type="text" placeholder="Add to wishlist..." value={wishInput} onChange={e => setWishInput(e.target.value)} onKeyDown={e => e.key === 'Enter' && addWishlistItem(c)} />
                    <Button onClick={() => addWishlistItem(c)} style={{ padding: '6px 10px', fontSize: 12 }}>+</Button>
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

function CouponsView({ coupons, onAdd, onDelete }) {
  const COLORS = useColors();
  const [showAdd, setShowAdd] = useState(false);
  return (
    <div>
      <Button onClick={() => setShowAdd(true)} style={{ marginBottom: 18 }}><Plus size={14} /> Add coupon</Button>
      {coupons.length === 0 ? (
        <div style={{ textAlign: 'center', padding: 40, color: COLORS.creamDim, fontFamily: 'Jost, sans-serif' }}>No coupons yet.</div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {coupons.map(c => (
            <div key={c.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 10, padding: '12px 16px' }}>
              <div style={{ color: COLORS.cream, fontFamily: 'Jost, sans-serif', fontWeight: 600 }}>{c.code} · {c.type === 'percent' ? `${c.value}% off` : `$${c.value.toFixed(2)} off`}</div>
              <button onClick={() => onDelete(c.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.oxblood }}><Trash2 size={16} /></button>
            </div>
          ))}
        </div>
      )}
      {showAdd && <AddCouponModal onClose={() => setShowAdd(false)} onSave={(c) => { onAdd(c); setShowAdd(false); }} />}
    </div>
  );
}

function POSView({ items, customers, coupons, platforms, onCheckout, onAddCustomer }) {
  const COLORS = useColors();
  const inputStyle = useInputStyle();
  const [query, setQuery] = useState('');
  const [cart, setCart] = useState([]);
  const [customerQuery, setCustomerQuery] = useState('');
  const [selectedCustomer, setSelectedCustomer] = useState(null);
  const [showQuickAdd, setShowQuickAdd] = useState(false);
  const [showCameraScan, setShowCameraScan] = useState(false);
  const [couponCode, setCouponCode] = useState('');
  const [redeemPoints, setRedeemPoints] = useState('');
  const [isOnlineOrder, setIsOnlineOrder] = useState(false);
  const [shippingAddress, setShippingAddress] = useState('');
  const [channel, setChannel] = useState('instore');
  const [shippingCost, setShippingCost] = useState('');
  const [shippingPaidBy, setShippingPaidBy] = useState('user');
  const [platformFees, setPlatformFees] = useState('');
  const [promotionCosts, setPromotionCosts] = useState('');
  const inputRef = useRef(null);

  useEffect(() => { 
    inputRef.current?.focus(); 
    // Initialize fees for the current channel (In-Store by default)
    const defaultPlatform = platforms.find(p => p.id === 'instore');
    if (defaultPlatform) {
      setPlatformFees(String(defaultPlatform.platformFee || 0));
      setPromotionCosts(String(defaultPlatform.promotionFee || 0));
    }
  }, []);

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
  const searchResults = query && !items.some(i => i.sku.toLowerCase() === query.toLowerCase()) ? items.filter(i => i.name.toLowerCase().includes(query.toLowerCase())).slice(0, 6) : [];

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
    onCheckout({ cartLines, customer: selectedCustomer, couponCode: appliedCoupon?.code, couponDiscount, pointsRedeemed: pointsToRedeem, pointsDiscount, pointsEarned, total, isOnlineOrder, shippingAddress: isOnlineOrder ? shippingAddress : '', channel, shippingCost: parseFloat(shippingCost) || 0, shippingPaidBy, platformFees: parseFloat(platformFees) || 0, promotionCosts: parseFloat(promotionCosts) || 0 });
    setCart([]); setCouponCode(''); setRedeemPoints(''); setSelectedCustomer(null); setCustomerQuery(''); setIsOnlineOrder(false); setShippingAddress(''); setChannel('instore'); setShippingCost(''); setShippingPaidBy('user'); setPlatformFees(''); setPromotionCosts('');
  };

  return (
    <div style={{ display: 'flex', gap: 20 }}>
      <div style={{ flex: 1 }}>
        <div style={{ marginBottom: 18 }}>
          <input ref={inputRef} style={{ ...inputStyle, width: '100%', marginBottom: 10 }} type="text" placeholder="SKU or product name..." value={query} onChange={e => setQuery(e.target.value)} onKeyDown={handleKeyDown} />
          <Button onClick={() => setShowCameraScan(true)} style={{ width: '100%', justifyContent: 'center' }}><Camera size={14} /> Scan barcode</Button>
        </div>
        {showCameraScan && <CameraScanModal onDetect={addToCartBySku} onClose={() => setShowCameraScan(false)} />}
        {searchResults.length > 0 && (
          <div style={{ background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 10, marginBottom: 18 }}>
            {searchResults.map(item => (
              <div key={item.id} onClick={() => { addByClick(item); setQuery(''); }} style={{ padding: '10px 12px', background: COLORS.espressoSoft, borderRadius: 8, cursor: 'pointer', fontFamily: 'Jost, sans-serif', display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
                <span style={{ color: COLORS.cream }}>{item.name} {item.condition ? `(${item.condition})` : ''}</span>
                <span style={{ color: COLORS.brass, fontWeight: 600 }}>${item.price.toFixed(2)}</span>
              </div>
            ))}
          </div>
        )}

        <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, marginBottom: 10, fontSize: 15 }}>Customer (optional)</div>
        {selectedCustomer ? (
          <div style={{ background: COLORS.espressoSoft, border: `1px solid ${COLORS.brass}`, borderRadius: 10, padding: '10px 12px', marginBottom: 16, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ color: COLORS.cream, fontFamily: 'Jost, sans-serif', fontWeight: 600 }}>{selectedCustomer.name} · {selectedCustomer.points} pts</span>
            <button onClick={() => setSelectedCustomer(null)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.creamDim }}><X size={16} /></button>
          </div>
        ) : (
          <>
            <input style={{ ...inputStyle, width: '100%', marginBottom: 10 }} type="text" placeholder="Search customers..." value={customerQuery} onChange={e => setCustomerQuery(e.target.value)} />
            {customerMatches.length > 0 && (
              <div style={{ background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 10, marginBottom: 16 }}>
                {customerMatches.slice(0, 5).map(c => (
                  <div key={c.id} onClick={() => { setSelectedCustomer(c); setCustomerQuery(''); }} style={{ padding: '8px 12px', cursor: 'pointer', fontFamily: 'Jost, sans-serif', fontSize: 13, color: COLORS.cream, borderBottom: `1px solid ${COLORS.brown}` }}>
                    {c.name} ({c.points} pts)
                  </div>
                ))}
              </div>
            )}
            <Button onClick={() => setShowQuickAdd(true)} style={{ marginBottom: 16, fontSize: 12, color: COLORS.brass, background: 'transparent', border: `1px solid ${COLORS.brass}` }}><Plus size={12} /> New customer</Button>
          </>
        )}
        {showQuickAdd && <QuickAddCustomerModal onClose={() => setShowQuickAdd(false)} onSave={(c) => { onAddCustomer(c); setSelectedCustomer(c); setShowQuickAdd(false); }} />}

        {selectedCustomer && selectedCustomer.points > 0 && (
          <Field label="Redeem points"><input style={inputStyle} type="number" value={redeemPoints} onChange={e => setRedeemPoints(e.target.value)} placeholder="0" /></Field>
        )}
      </div>

      <div style={{ width: 360, background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 12, padding: 16 }}>
        <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, marginBottom: 10, fontSize: 15 }}>Current sale</div>
        {cartLines.length === 0 ? (
          <div style={{ color: COLORS.creamDim, textAlign: 'center', padding: '20px 0', fontFamily: 'Jost, sans-serif' }}>Cart is empty</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 16, maxHeight: 300, overflowY: 'auto' }}>
            {cartLines.map(l => (
              <div key={l.itemId} style={{ background: COLORS.paper, borderRadius: 8, padding: '10px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ flex: 1 }}>
                  <div style={{ color: COLORS.cream, fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 12 }}>{l.item.name}</div>
                  <div style={{ color: COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 11 }}>${l.item.price.toFixed(2)} each</div>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <button onClick={() => updateQty(l.itemId, -1)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.brass }}><Minus size={14} /></button>
                  <span style={{ color: COLORS.cream, fontFamily: 'JetBrains Mono, monospace', minWidth: 20, textAlign: 'center' }}>{l.qty}</span>
                  <button onClick={() => updateQty(l.itemId, 1)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.brass }}><Plus size={14} /></button>
                  <button onClick={() => removeFromCart(l.itemId)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.oxblood, marginLeft: 8 }}><Trash2 size={14} /></button>
                </div>
              </div>
            ))}
          </div>
        )}

        <Field label="Coupon code"><input style={inputStyle} type="text" value={couponCode} onChange={e => { setCouponCode(e.target.value.toUpperCase()); }} placeholder="WELCOME10" /></Field>

        <div style={{ background: COLORS.paper, borderRadius: 8, padding: 12, marginBottom: 16, fontFamily: 'Jost, sans-serif', fontSize: 13 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6, color: COLORS.creamDim }}>
            <span>Subtotal</span>
            <span>${subtotal.toFixed(2)}</span>
          </div>
          {couponDiscount > 0 && (
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6, color: COLORS.creamDim, fontSize: 12 }}>
              <span>Coupon ({appliedCoupon.code})</span>
              <span>-${couponDiscount.toFixed(2)}</span>
            </div>
          )}
          {pointsDiscount > 0 && (
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6, color: COLORS.creamDim, fontSize: 12 }}>
              <span>Points ({pointsToRedeem})</span>
              <span>-${pointsDiscount.toFixed(2)}</span>
            </div>
          )}
          <div style={{ display: 'flex', justifyContent: 'space-between', paddingTop: 8, borderTop: `1px solid ${COLORS.brown}`, fontFamily: 'Fraunces, serif', fontSize: 16, fontWeight: 700, color: COLORS.brass }}>
            <span>Total</span>
            <span>${total.toFixed(2)}</span>
          </div>
        </div>

        {selectedCustomer && <div style={{ color: COLORS.brass, fontFamily: 'Jost, sans-serif', fontSize: 11, marginBottom: 12, textAlign: 'center' }}>Earns {pointsEarned} pts for {selectedCustomer.name}</div>}

        <div style={{ marginBottom: 16 }}>
          <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontWeight: 700, marginBottom: 8, fontSize: 13 }}>Sale Channel & Costs</div>
          
          <Field label="Channel">
            <select 
              style={useInputStyle()} 
              value={channel} 
              onChange={e => {
                const selectedChannelId = e.target.value;
                setChannel(selectedChannelId);
                // Auto-populate fees from platform config
                const platform = platforms.find(p => p.id === selectedChannelId);
                if (platform) {
                  setPlatformFees(String(platform.platformFee || 0));
                  setPromotionCosts(String(platform.promotionFee || 0));
                }
              }}
            >
              {platforms.map(p => (
                <option key={p.id} value={p.id}>{p.name}</option>
              ))}
            </select>
          </Field>

          <Field label="Shipping Cost $">
            <input style={useInputStyle()} type="number" value={shippingCost} onChange={e => setShippingCost(e.target.value)} placeholder="0.00" step="0.01" />
          </Field>

          {shippingCost && (
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12, fontFamily: 'Jost, sans-serif', fontSize: 12, color: COLORS.cream, cursor: 'pointer' }}>
              <input type="radio" checked={shippingPaidBy === 'user'} onChange={() => setShippingPaidBy('user')} />
              Customer paid
            </label>
          )}
          {shippingCost && (
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12, fontFamily: 'Jost, sans-serif', fontSize: 12, color: COLORS.cream, cursor: 'pointer' }}>
              <input type="radio" checked={shippingPaidBy === 'me'} onChange={() => setShippingPaidBy('me')} />
              I paid (deduct from profit)
            </label>
          )}

          <Field label="Platform Fees $">
            <input style={useInputStyle()} type="number" value={platformFees} onChange={e => setPlatformFees(e.target.value)} placeholder="0.00" step="0.01" />
          </Field>

          <Field label="Promotion Cost $">
            <input style={useInputStyle()} type="number" value={promotionCosts} onChange={e => setPromotionCosts(e.target.value)} placeholder="0.00" step="0.01" />
          </Field>
        </div>

        <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12, fontFamily: 'Jost, sans-serif', fontSize: 12, color: COLORS.cream, cursor: 'pointer' }}>
          <input type="checkbox" checked={isOnlineOrder} onChange={e => setIsOnlineOrder(e.target.checked)} />
          Online order
        </label>
        {isOnlineOrder && (
          <Field label="Shipping address"><textarea style={{ ...inputStyle, minHeight: 80, fontFamily: 'Jost, sans-serif', resize: 'none' }} value={shippingAddress} onChange={e => setShippingAddress(e.target.value)} /></Field>
        )}

        <Button onClick={handleCheckout} style={{ width: '100%', justifyContent: 'center', opacity: cartLines.length ? 1 : 0.5 }}>
          <Check size={14} /> Complete sale
        </Button>
      </div>
    </div>
  );
}

const REPORT_TYPES = [
  { id: 'sales', label: 'Sales' },
  { id: 'inventory', label: 'Inventory' },
  { id: 'customers', label: 'Customers' },
  { id: 'topSellers', label: 'Top Sellers' },
  { id: 'profitAndLoss', label: 'P&L (Profit & Loss)' },
  { id: 'channelBreakdown', label: 'Channel Breakdown' },
];

function PlatformsView({ platforms, onAddPlatform, onDeletePlatform, onUpdatePlatform }) {
  const COLORS = useColors();
  const inputStyle = useInputStyle();
  const [newName, setNewName] = useState('');
  const [newPlatformFee, setNewPlatformFee] = useState('');
  const [newPromotionFee, setNewPromotionFee] = useState('');

  const handleAdd = () => {
    if (!newName.trim()) return;
    const id = newName.toLowerCase().replace(/\s+/g, '-');
    onAddPlatform({ 
      id, 
      name: newName, 
      icon: 'custom',
      platformFee: parseFloat(newPlatformFee) || 0,
      promotionFee: parseFloat(newPromotionFee) || 0,
    });
    setNewName(''); 
    setNewPlatformFee('');
    setNewPromotionFee('');
  };

  return (
    <div style={{ padding: 20 }}>
      <h2 style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontSize: 20, marginBottom: 20 }}>Sales Platforms & Default Fees</h2>
      <p style={{ fontFamily: 'Jost, sans-serif', color: COLORS.creamDim, fontSize: 13, marginBottom: 20 }}>Set platform fees and default promotion costs. These will auto-fill in POS when you select a channel.</p>
      
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: 16, marginBottom: 30 }}>
        {platforms.map(p => (
          <div key={p.id} style={{ background: COLORS.espressoSoft, border: `1px solid ${COLORS.brown}`, borderRadius: 8, padding: 16 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 12 }}>
              <div>
                <div style={{ fontFamily: 'Fraunces, serif', fontWeight: 700, color: COLORS.brass, fontSize: 14 }}>{p.name}</div>
                <div style={{ fontFamily: 'Jost, sans-serif', color: COLORS.creamDim, fontSize: 11, marginTop: 2 }}>ID: {p.id}</div>
              </div>
              {p.id !== 'instore' && (
                <button onClick={() => onDeletePlatform(p.id)} style={{ padding: '4px 8px', background: COLORS.oxblood, color: COLORS.paper, border: 'none', borderRadius: 4, cursor: 'pointer', fontFamily: 'Jost, sans-serif', fontSize: 10, fontWeight: 600 }}>
                  Remove
                </button>
              )}
            </div>
            
            <div style={{ marginBottom: 10 }}>
              <label style={{ display: 'block', fontFamily: 'Jost, sans-serif', fontSize: 11, color: COLORS.creamDim, marginBottom: 4, fontWeight: 600 }}>Platform Fee $</label>
              <input
                type="number"
                value={p.platformFee || ''}
                onChange={e => onUpdatePlatform(p.id, { platformFee: parseFloat(e.target.value) || 0 })}
                placeholder="0.00"
                step="0.01"
                style={{ ...inputStyle, width: '100%', padding: '8px 12px', fontFamily: 'Jost, sans-serif', fontSize: 13 }}
              />
              <div style={{ fontFamily: 'Jost, sans-serif', fontSize: 10, color: COLORS.creamDim, marginTop: 2 }}>e.g., eBay fees, TCG Player commission</div>
            </div>

            <div>
              <label style={{ display: 'block', fontFamily: 'Jost, sans-serif', fontSize: 11, color: COLORS.creamDim, marginBottom: 4, fontWeight: 600 }}>Default Promo Cost $</label>
              <input
                type="number"
                value={p.promotionFee || ''}
                onChange={e => onUpdatePlatform(p.id, { promotionFee: parseFloat(e.target.value) || 0 })}
                placeholder="0.00"
                step="0.01"
                style={{ ...inputStyle, width: '100%', padding: '8px 12px', fontFamily: 'Jost, sans-serif', fontSize: 13 }}
              />
              <div style={{ fontFamily: 'Jost, sans-serif', fontSize: 10, color: COLORS.creamDim, marginTop: 2 }}>Optional: ads, promoted listings, boosts</div>
            </div>
          </div>
        ))}
      </div>

      <div style={{ background: COLORS.paper, border: `1px solid ${COLORS.brown}`, borderRadius: 8, padding: 16 }}>
        <div style={{ fontFamily: 'Fraunces, serif', fontWeight: 700, color: COLORS.brass, marginBottom: 16, fontSize: 14 }}>Add New Platform</div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr auto', gap: 10, alignItems: 'flex-end' }}>
          <div>
            <label style={{ display: 'block', fontFamily: 'Jost, sans-serif', fontSize: 11, color: COLORS.creamDim, marginBottom: 4, fontWeight: 600 }}>Platform Name</label>
            <input
              type="text"
              value={newName}
              onChange={e => setNewName(e.target.value)}
              placeholder="e.g., Depop, Poshmark"
              style={{ ...inputStyle, width: '100%', padding: '8px 12px', fontFamily: 'Jost, sans-serif', fontSize: 13 }}
            />
          </div>
          <div>
            <label style={{ display: 'block', fontFamily: 'Jost, sans-serif', fontSize: 11, color: COLORS.creamDim, marginBottom: 4, fontWeight: 600 }}>Fee $</label>
            <input
              type="number"
              value={newPlatformFee}
              onChange={e => setNewPlatformFee(e.target.value)}
              placeholder="0.00"
              step="0.01"
              style={{ ...inputStyle, width: '100%', padding: '8px 12px', fontFamily: 'Jost, sans-serif', fontSize: 13 }}
            />
          </div>
          <div>
            <label style={{ display: 'block', fontFamily: 'Jost, sans-serif', fontSize: 11, color: COLORS.creamDim, marginBottom: 4, fontWeight: 600 }}>Promo $</label>
            <input
              type="number"
              value={newPromotionFee}
              onChange={e => setNewPromotionFee(e.target.value)}
              placeholder="0.00"
              step="0.01"
              style={{ ...inputStyle, width: '100%', padding: '8px 12px', fontFamily: 'Jost, sans-serif', fontSize: 13 }}
            />
          </div>
          <Button onClick={handleAdd} style={{ minWidth: 100 }}>Add</Button>
        </div>
      </div>
    </div>
  );
}

function ReportsView({ items, sales, customers }) {
  const COLORS = useColors();
  const inputStyle = useInputStyle();
  const [reportType, setReportType] = useState('sales');
  const [fromDate, setFromDate] = useState('');
  const [toDate, setToDate] = useState('');

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
    profitAndLoss: {
      headers: [
        { key: 'metric', label: 'Metric' },
        { key: 'amount', label: 'Amount' },
      ],
      rows: (() => {
        const totalRevenue = filteredSales.reduce((sum, s) => sum + s.total, 0);
        const totalCOGS = filteredSales.reduce((sum, s) => sum + (s.costOfGoods || 0), 0);
        const totalShipping = filteredSales.reduce((sum, s) => sum + (s.shippingPaidBy === 'me' ? s.shippingCost : 0), 0);
        const totalFees = filteredSales.reduce((sum, s) => sum + (s.platformFees || 0), 0);
        const totalPromo = filteredSales.reduce((sum, s) => sum + (s.promotionCosts || 0), 0);
        const netProfit = totalRevenue - totalCOGS - totalShipping - totalFees - totalPromo;
        return [
          { metric: 'Total Revenue', amount: totalRevenue.toFixed(2) },
          { metric: 'Cost of Goods', amount: totalCOGS.toFixed(2) },
          { metric: 'Gross Profit', amount: (totalRevenue - totalCOGS).toFixed(2) },
          { metric: 'Shipping (you paid)', amount: totalShipping.toFixed(2) },
          { metric: 'Platform Fees', amount: totalFees.toFixed(2) },
          { metric: 'Promotion Costs', amount: totalPromo.toFixed(2) },
          { metric: 'NET PROFIT', amount: netProfit.toFixed(2) },
        ];
      })(),
    },
    channelBreakdown: {
      headers: [
        { key: 'channel', label: 'Channel' },
        { key: 'sales', label: 'Sales' },
        { key: 'revenue', label: 'Revenue' },
        { key: 'cogs', label: 'COGS' },
        { key: 'profit', label: 'Net Profit' },
      ],
      rows: (() => {
        const channels = {};
        filteredSales.forEach(s => {
          const ch = s.channel || 'instore';
          if (!channels[ch]) channels[ch] = { channel: ch, sales: 0, revenue: 0, cogs: 0, totalFees: 0, totalPromo: 0, totalShipping: 0 };
          channels[ch].sales += 1;
          channels[ch].revenue += s.total;
          channels[ch].cogs += s.costOfGoods || 0;
          channels[ch].totalFees += s.platformFees || 0;
          channels[ch].totalPromo += s.promotionCosts || 0;
          if (s.shippingPaidBy === 'me') channels[ch].totalShipping += s.shippingCost;
        });
        return Object.values(channels).map(c => ({
          channel: c.channel.charAt(0).toUpperCase() + c.channel.slice(1),
          sales: c.sales,
          revenue: c.revenue.toFixed(2),
          cogs: c.cogs.toFixed(2),
          profit: (c.revenue - c.cogs - c.totalShipping - c.totalFees - c.totalPromo).toFixed(2),
        }));
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
      <body style="font-family:sans-serif;padding:24px;">
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
          <button key={r.id} onClick={() => setReportType(r.id)} style={{ padding: '8px 16px', borderRadius: 8, border: `1px solid ${reportType === r.id ? COLORS.brass : COLORS.brown}`, background: reportType === r.id ? COLORS.brass : 'transparent', color: reportType === r.id ? COLORS.espresso : COLORS.cream, fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 13, cursor: 'pointer' }}>
            {r.label}
          </button>
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
        <Button variant="ghost" onClick={handlePrint}><FileText size={14} /> Print / PDF</Button>
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

const FULFILLMENT_STATUSES = ['unfulfilled', 'packed', 'shipped', 'delivered'];
function OrdersView({ sales, onUpdate }) {
  const COLORS = useColors();
  const inputStyle = useInputStyle();
  const onlineOrders = sales.filter(s => s.channel === 'online').slice().reverse();
  return (
    <div>
      {onlineOrders.length === 0 ? (
        <div style={{ textAlign: 'center', padding: '60px 20px', color: COLORS.creamDim, fontFamily: 'Jost, sans-serif' }}>
          No online orders yet.
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
                <select style={{ ...inputStyle, padding: '6px 10px', fontSize: 12 }} value={s.fulfillmentStatus || 'unfulfilled'} onChange={e => onUpdate({ ...s, fulfillmentStatus: e.target.value })}>
                  {FULFILLMENT_STATUSES.map(st => <option key={st} value={st}>{st}</option>)}
                </select>
                <input style={{ ...inputStyle, padding: '6px 10px', fontSize: 12, flex: 1, minWidth: 160 }} placeholder="Tracking number" value={s.trackingNumber || ''} onChange={e => onUpdate({ ...s, trackingNumber: e.target.value })} />
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [tab, setTab] = useState('dashboard');
  const [theme, setTheme] = useState(() => {
    try { return localStorage.getItem('pullbar-theme') || 'light'; } catch { return 'light'; }
  });
  const [items, setItems] = useState([]);
  const [sales, setSales] = useState([]);
  const [customers, setCustomers] = useState([]);
  const [coupons, setCoupons] = useState([]);
  const [platforms, setPlatforms] = useState(() => {
    try { const p = localStorage.getItem(storageKeys.platforms); return p ? JSON.parse(p) : DEFAULT_PLATFORMS; } catch { return DEFAULT_PLATFORMS; }
  });
  const [loaded, setLoaded] = useState(false);
  
  const COLORS = getColors(theme);
  const handleThemeChange = (newTheme) => {
    setTheme(newTheme);
    try { localStorage.setItem('pullbar-theme', newTheme); } catch {}
  };

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
  const handleAddPlatform = (platform) => { const next = [...platforms, platform]; setPlatforms(next); saveKey(storageKeys.platforms, next); };
  const handleDeletePlatform = (id) => { const next = platforms.filter(p => p.id !== id); setPlatforms(next); saveKey(storageKeys.platforms, next); };
  const handleUpdatePlatform = (id, updates) => { const next = platforms.map(p => p.id === id ? { ...p, ...updates } : p); setPlatforms(next); saveKey(storageKeys.platforms, next); };

  const handleCheckout = ({ cartLines, customer, couponCode, couponDiscount, pointsRedeemed, pointsDiscount, pointsEarned, total, isOnlineOrder, shippingAddress, channel = 'instore', shippingCost = 0, shippingPaidBy = 'user', platformFees = 0, promotionCosts = 0 }) => {
    const receiptNumber = `REC-${String(sales.length + 1).padStart(4, '0')}`;
    const costOfGoods = cartLines.reduce((sum, l) => sum + (l.item.cost * l.qty), 0);
    const netProfit = total - costOfGoods - (shippingPaidBy === 'user' ? 0 : shippingCost) - platformFees - promotionCosts;
    
    const sale = {
      id: uid(), receiptNumber, date: new Date().toISOString(), total, couponCode: couponCode || null,
      couponDiscount, pointsRedeemed, pointsDiscount, pointsEarned,
      customerId: customer?.id || null, customerName: customer?.name || null,
      lines: cartLines.map(l => ({ name: l.item.name, qty: l.qty, price: l.item.price, sku: l.item.sku, cost: l.item.cost })),
      channel, // ebay, whatnot, tcgplayer, tiktok, mercari, instore, etc.
      shippingCost, // Amount paid for shipping
      shippingPaidBy, // 'user' or 'me'
      platformFees, // eBay fees, TCG Player fees, etc.
      promotionCosts, // Ad spend, promotion costs
      costOfGoods, // Total cost of items sold
      netProfit, // Revenue - COGS - (user-paid shipping) - fees - promo
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
    alert(`Sale complete! ${receiptNumber}\nTotal: $${total.toFixed(2)}\nNet Profit: $${netProfit.toFixed(2)}`);
  };

  const tabs = [
    { id: 'dashboard', label: 'Dashboard', icon: BarChart3 },
    { id: 'inventory', label: 'Inventory', icon: LayoutGrid },
    { id: 'pos', label: 'POS', icon: ShoppingCart },
    { id: 'customers', label: 'Customers', icon: Users },
    { id: 'coupons', label: 'Coupons', icon: Tag },
    { id: 'platforms', label: 'Platforms', icon: Globe },
    { id: 'reports', label: 'Reports', icon: FileText },
    { id: 'orders', label: 'Orders', icon: Truck },
  ];

  return (
    <ColorContext.Provider value={COLORS}>
      <div style={{ minHeight: '100vh', background: COLORS.paper, fontFamily: 'Jost, sans-serif' }}>
        <style>{`
          @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,700&family=Jost:wght@400;500;600;700&family=JetBrains+Mono:wght@400;700&display=swap');
          * { box-sizing: border-box; }
          input:focus, select:focus, textarea:focus { border-color: ${COLORS.brass} !important; }
          body { margin: 0; }
        `}</style>
        <div style={{ maxWidth: 1400, margin: '0 auto', padding: '20px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 20, gap: 14, flexWrap: 'wrap' }}>
            <div style={{ fontFamily: 'Fraunces, serif', color: COLORS.cream, fontSize: 20, fontWeight: 700, letterSpacing: 1 }}>THE PULL <span style={{ color: COLORS.brass }}>BAR</span></div>
            <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
              <div style={{ display: 'flex', gap: 4, background: COLORS.espressoSoft, padding: 4, borderRadius: 8, border: `1px solid ${COLORS.brown}` }}>
                {['light', 'dark', 'highContrast'].map(t => (
                  <button key={t} onClick={() => handleThemeChange(t)} style={{ padding: '6px 10px', borderRadius: 6, border: 'none', cursor: 'pointer', background: theme === t ? COLORS.brass : 'transparent', color: theme === t ? COLORS.espresso : COLORS.creamDim, fontFamily: 'Jost, sans-serif', fontSize: 11, fontWeight: 600 }}>
                    {t === 'highContrast' ? 'HC' : t.slice(0, 1).toUpperCase() + t.slice(1)}
                  </button>
                ))}
              </div>
              <div style={{ display: 'flex', gap: 6, background: COLORS.espressoSoft, padding: 4, borderRadius: 8, border: `1px solid ${COLORS.brown}`, flexWrap: 'wrap' }}>
                {tabs.map(t => (
                  <button key={t.id} onClick={() => setTab(t.id)} style={{ display: 'flex', alignItems: 'center', gap: 6, border: 'none', cursor: 'pointer', padding: '6px 12px', borderRadius: 6, fontFamily: 'Jost, sans-serif', fontWeight: 600, fontSize: 12, background: tab === t.id ? COLORS.brass : 'transparent', color: tab === t.id ? COLORS.espresso : COLORS.creamDim }}>
                    <t.icon size={14} /> {t.label}
                  </button>
                ))}
              </div>
            </div>
          </div>

          {!loaded ? (
            <div style={{ color: COLORS.creamDim, textAlign: 'center', padding: 60 }}>Loading...</div>
          ) : (
            <>
              {tab === 'dashboard' && <Dashboard items={items} sales={sales} customers={customers} onApplyMarkdown={handleApplyMarkdown} />}
              {tab === 'inventory' && <InventoryView items={items} onAdd={handleAddItem} onDelete={handleDeleteItem} />}
              {tab === 'pos' && <POSView items={items} customers={customers} coupons={coupons} platforms={platforms} onCheckout={handleCheckout} onAddCustomer={handleAddCustomer} />}
              {tab === 'customers' && <CustomersView customers={customers} sales={sales} onAdd={handleAddCustomer} onUpdate={handleUpdateCustomer} />}
              {tab === 'coupons' && <CouponsView coupons={coupons} onAdd={handleAddCoupon} onDelete={handleDeleteCoupon} />}
              {tab === 'platforms' && <PlatformsView platforms={platforms} onAddPlatform={handleAddPlatform} onDeletePlatform={handleDeletePlatform} onUpdatePlatform={handleUpdatePlatform} />}
              {tab === 'reports' && <ReportsView items={items} sales={sales} customers={customers} />}
              {tab === 'orders' && <OrdersView sales={sales} onUpdate={handleUpdateSale} />}
            </>
          )}
        </div>
      </div>
    </ColorContext.Provider>
  );
}

