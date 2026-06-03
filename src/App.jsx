import { useState, useCallback, useEffect } from "react";
import { supabase } from "./supabase.js";
import { generateInternalExcel, generateClientExcelMisura, generateClientPdfCatalogo } from "./excel.js";

const LOGO = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCAFsApoDASIAAhEBAxEB/8QAHQABAAICAwEBAAAAAAAAAAAAAAYHBQgBAgQDCf/EAEsQAAEDBAAEAwQGBgYIBQUBAAEAAgMEBQYRBxIhMQgTQRQiMlEVYXGBkbEjM0JScqEWFyRissE0NjdDc9Hh8CUoOHSCGCZTVJKi/8QAGwEBAAIDAQEAAAAAAAAAAAAAAAMEAQIFBgf/xAA5EQACAgIABQIEBAQFAwUAAAAAAQIDBBEFEiExQRNRBiJhcRQygZEjM7HBFUKh0eEkNDUWUmLw8f/aAAwDAQACEQMRAD8A3LREQBERAEREAREQBEXBOkByi42ulRPHBC6WRwaxo2SVhvQPoiwON5Rbr9PURUMgf5B04r31d3t1I7lnqmMP1laRthKPMn0N3XJPla6nvRfKmnjqIhLC4OYexX1C3T2aBERZARfKpqIqaIyzODGN7kry0d3t9Y7lp6qN5+orVzinpszptbPeiwWSZRbbDUU0VdKGeedArMQTxzwtlicHMcNghYjZGUnFPqjLi0k2j6ouASuVuahERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBEQoAuCnVee41TKOjlqZPhjbtYbS6sd3pHwv9yhtFqmr5yAyNu1XxysZliVU2ld5DwCHHfoFxNktDnVlq6Rh8vyth433UCxSvtlttl2o5Ji14IEYAPXr1XleJcUnKz0qWtOL6nYxsRRi5TXzJo9XDa60+N09xm5nODj8WidryU1rvOd5IJaWSaKha/bnHY2vXassxyyWplBNajPzO94nfXqrjw2ppJ8dZW0dCKVjmkhmuqqcOwo5SjB3bUeukT5F86G58nV9NmUstEy222Gja/YjbrZPdezmb+8PxVNXSv4g3rJKmO1c1PRQnTTodVjJbnxApr/T2htQ6aR7vfIA6Bd//ABOEOig9duxQ/Ayl3kt9y+guj3sB057R9pWEut2bj2NGtr5NvjZ136lV7bzlWX2+a8QVD6Vg2YGa3vSt3ZirkoJbk+uivXjuS5m9L3LQvlA252qaic7lEjdbC17mtt6wHJi+sfNLQuftrxs6U+4XZ5XVd1mxzIW+XXQnla4/tKaZ1NR0tilq6yhFXGwbLddVzcyiriVKtjJxcfPsXKZ2YdjrnHakU9xMulJkkdula57Wns7RGipezLBhuJ0bap3tDjob36bUMuuW45fLO6igtRgdE73D12OvVebJ6+2XGktFEyYu3sSAg9PkvPLKnj2zuhZzPXXx/odD0OeEapx0tmwNhucF2tkNdAQWSN2veqtpskocGsFDSPPmOm6NG+3VWTQ1LKukjqYztsjQQvZYeZDIgtP5tLZw76HW966eD0ouFyFdK4REQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREA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j41qPcKartddJQ3SB8E8Z0eYd1xr6JUy0+x9N4TxiniVXNHpJd0fNFxzNI2HDS6Omib67PyAUB2D6Ivl5z9b9nl5fnyldmSsf0B0fkeiGDuia+sLqZIwdc42hnR2RPTfdEMBERAEREAREQBERAEREAHqro8HNSY8xqqcu0HEdPuVLjurT8Kc3k8S2x70HH/JT4z1bE5XHIKfD7U/Y3aHdcrhvYLlehPjIREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAFwVyiA6HqO3RQDiRwoxrNGOfV0rIqk/7xo/5KwkWsoKS0yajItx589ctM1bqPDB/bCIbkfJ3091S/FPDjjNtkbNcH+1uHcEEK9U0oFiVJ70dS34h4hZHldhC28McLFN5H0RFy60qy4g+HKzXN76mySeySHqGAbWwKLeVFclporY/FszHnzwsZpu3w1ZT7UGmqPlb6nQ7KdUHhhtP0aBU1x9pI6nlPdbHIVFHCqj4L93xPxG3Xz6+xo/xI4IZLiXmVlA11ZRt69B6KrDMGOLKiN8T29CCCv0unhiqIzHNG17CNEOG1WmW8E8Ov9Wal9EyKRx24j1Ve3h/ms7fDfjBxXLlrf1Ro82eB5012j9i7ra3N/Dxj39G5nWiPkrImksI31WqtVS1FBXTW+sYY54XFpBVC6iVP5j1nDeL4/EU3V3XhnQouPtXKhOoEXD3NYOvf0A9V9zQXZtL7YbbMKf8Ae5Si6mHJLu9HxRdWPbINjv6g+i7BDIREQBWH4a5RFxTptnWyfyVeBTDgXOKfipQdT1d/kpKXqyL+pR4nHnw7Y/Rm/jDtjfsC7L50/WCM/wB0fkvovSHxJhERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAcEIQuUQHUtBBB6g+ioHj/wTbkUz79YQIq7W3saPjWwC4cAe42o7K42R1It4WbbhWq2p6Z+b96sl/sc7oLrbJWOadbAJXjoorhXTCCjt8z5D0Hulfotdscst037bQwyE/NoXkt+GY3QSiSmtkLHfPlC5z4d16Poezh8ark+av5vua0cEOBtxuVdFd8phMcDTzNhd6rZw4nYXWwW82+EwBvLrlCzbGMY3lY0NA7ADS7K9VRCuOkeV4hxfJzrfUm9eyXg1O4+8EH2rzMgxeEmIbdJA0KgY38xLHNMcrTpzHdNFfpVPFHPE6KVgexw0QRtUTxY8P1ryCeW42Nwo6p/Uho7lUsnC2+as9NwP4oVcVTlv7P/AHNTOvyXD3CNvM86Cs2t4FZ1TVJhjYZGg6DtBTvhr4c6h1dHXZRNzxg78ohU4YtsnrR6jI49g018/Pv6I1zE8hIDKSVxPQe4eqv7ww8L6+pvceVXeB0UcfWJrh3WwtHw3xGmEfl2qL3ANdFKaSmgpIWwU8TY429AGjSv0YPJLmk9nkOK/FjyqXTTHW+7PowBoDR2A0F3C6hdl0TxgREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREATSIgCIiAdEREB15Wnu0H7lz0HQBcogCIiAaREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREB//9k=";


// ─── UTILS ───
const uid = () => Math.random().toString(36).slice(2, 9);
const fmt = (v) => v == null || v === "" ? "—" : new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(v);
const round5 = (v) => Math.ceil(v / 5) * 5;
const fmtDate = (d) => { if (!d) return ""; try { return new Date(d + "T00:00:00").toLocaleDateString("it-IT", { weekday: "short", day: "numeric", month: "short" }); } catch { return d; } };
const nightsFrom = (a, b) => { if (!a || !b) return 0; return Math.max(0, Math.round((new Date(b) - new Date(a)) / 86400000)); };

// ─── DATA ───
const mkRoom = (name = "", rate = 0, tax = 0, occ = 2) => ({ id: uid(), name, rate, taxRate: tax, occupancy: occ, breakfast: 0 });
const mkLeg = () => ({ id: uid(), destination: "", dateFrom: "", dateTo: "", hotelName: "", roomTypes: [mkRoom("", 0, 0, 2)], baseRoomIndex: 0, days: [] });
const mkDay = () => ({ id: uid(), date: "", title: "", activities: [] });
const mkAct = () => ({ id: uid(), description: "", costType: "individual", amount: 0, staffMirror: [], isBase: false, clientPrice: 0 });
const mkStaff = (role = "") => ({ id: uid(), role, name: "", totalFee: 0, totalHotel: 0, flight: 0 });
const mkMisc = () => ({ id: uid(), description: "", amount: 0, costType: "group", isBase: false, clientPrice: 0 });
const mkPart = () => ({ id: uid(), firstName: "", lastName: "", roomChoices: {} });
const mkCommission = (name = "", pct = 0, enabled = true) => ({ id: uid(), name, pct, enabled });
const mkTrip = (type = "catalogo") => ({
  id: uid(), name: "", type, status: "planning", dateFrom: "", dateTo: "", totalParticipants: 14,
  markupPerPerson: 0, sellPrice: 0,
  commissions: [mkCommission("Commissione", 0, false)],
  legs: [mkLeg()], staff: [mkStaff("Guida"), mkStaff("")],
  miscCosts: [], createdBy: "", snapshot: null,
});

// ─── CALC ───
function calcTrip(trip) {
  const pax = trip.totalParticipants || 1;
  let totalGroupAll = 0, totalIndAll = 0, totalHotelBase = 0, totalNights = 0;
  const legBreakdowns = [];
  trip.legs.forEach(leg => {
    const nights = nightsFrom(leg.dateFrom, leg.dateTo);
    totalNights += nights;
    const base = leg.roomTypes[leg.baseRoomIndex]; if (!base) return;
    const hotelBase = ((base.rate + base.taxRate) / base.occupancy + (base.breakfast || 0)) * nights;
    totalHotelBase += hotelBase;
    let grp = 0, ind = 0;
    (leg.days || []).forEach(day => { (day.activities || []).forEach(act => {
      if (act.costType === "group") { grp += act.amount; grp += (act.staffMirror?.length || 0) * act.amount; }
      else { ind += act.amount; const guideCost = (act.staffMirror?.length || 0) * act.amount; if (guideCost > 0) grp += guideCost; }
    }); });
    totalGroupAll += grp; totalIndAll += ind;
    legBreakdowns.push({ leg, nights, hotelBase, grp, ind, baseRoom: base });
  });
  let staffTotal = 0;
  trip.staff.forEach(s => { staffTotal += (s.totalFee || 0) + (s.totalHotel || 0) + (s.flight || 0); });
  totalGroupAll += staffTotal;
  const miscTotal = trip.miscCosts.reduce((s, m) => s + (m.amount || 0), 0);
  totalGroupAll += miscTotal;
  const groupPP = totalGroupAll / pax;
  const subtotalPP = totalHotelBase + groupPP + totalIndAll;
  const withMarkup = subtotalPP + (trip.markupPerPerson || 0);
  let commTotal = 0;
  const commDetails = trip.commissions.filter(c => c.enabled).map(c => {
    const val = withMarkup * (c.pct / 100);
    commTotal += val;
    return { ...c, value: val };
  });
  const total = withMarkup + commTotal;
  const rounded = round5(total);
  const supplements = [];
  trip.legs.forEach(leg => {
    const base = leg.roomTypes[leg.baseRoomIndex]; if (!base) return;
    const nights = nightsFrom(leg.dateFrom, leg.dateTo);
    leg.roomTypes.forEach((r, ri) => {
      if (ri === leg.baseRoomIndex) return;
      const diff = ((r.rate + r.taxRate) / r.occupancy + (r.breakfast || 0)) - ((base.rate + base.taxRate) / base.occupancy + (base.breakfast || 0));
      supplements.push({ room: r, leg, diffPerNight: diff, total: diff * nights, rounded: round5(diff * nights) });
    });
  });
  return { pax, totalGroupAll, totalIndAll, totalHotelBase, groupPP, subtotalPP, withMarkup, commDetails, commTotal, total, rounded, supplements, legBreakdowns, staffTotal, miscTotal, totalNights };
}

// ─── STYLE ───
const P = { bg: "#FAF7F2", card: "#FFFFFF", accent: "#8B1A1A", accent2: "#6B1414", gold: "#C49B2B", goldBg: "#FDF8EC", red: "#8B1A1A", redBg: "#FDF0F0", green: "#2E7D32", greenBg: "#E8F5E9", blue: "#1565C0", blueBg: "#E3F2FD", text: "#2D2218", muted: "#8A7E70", border: "#E6DFD4", input: "#FDFCFA", cream: "#F3EDE2" };
const ff = "'Cormorant Garamond', 'Palatino', Georgia, serif";
const fs = "'Nunito Sans', 'Segoe UI', system-ui, sans-serif";

// ─── IVA HELPER ───
const IVA_RATES = [{ v: 0, l: "No IVA" }, { v: 4, l: "4%" }, { v: 10, l: "10%" }, { v: 22, l: "22%" }];
const addIva = (amount, rate) => rate > 0 ? Math.round((amount * (1 + rate / 100)) * 100) / 100 : amount;

function IvaToggle({ active, onChange }) {
  return (<button onClick={() => onChange(!active)} style={{ padding: "4px 10px", borderRadius: 5, fontSize: 10, fontWeight: 600, fontFamily: fs, cursor: "pointer", border: active ? "1.5px solid #1565C0" : "1px solid " + P.border, background: active ? "#E3F2FD" : "transparent", color: active ? "#1565C0" : P.muted, transition: "all .15s" }}>
    {active ? "IVA attiva" : "IVA"}
  </button>);
}

function IvaInput({ value, onChange, small }) {
  const [rate, setRate] = useState(22);
  const apply = () => { onChange(addIva(value || 0, rate)); };
  return (<div style={{ display: "inline-flex", gap: 3, alignItems: "center" }}>
    <select value={rate} onChange={e => setRate(Number(e.target.value))} style={{ padding: "2px 4px", border: "1px solid #B3D4FC", borderRadius: 3, fontSize: 10, fontFamily: fs, background: "#E3F2FD", color: "#1565C0", cursor: "pointer" }}>
      {IVA_RATES.filter(r => r.v > 0).map(r => <option key={r.v} value={r.v}>{r.l}</option>)}
    </select>
    <button onClick={apply} style={{ padding: "2px 6px", borderRadius: 3, fontSize: 9, fontWeight: 700, fontFamily: fs, background: "#1565C0", color: "#fff", border: "none", cursor: "pointer" }}>+IVA</button>
  </div>);
}

// ─── ATOMS ───
function I({ label, value, onChange, type = "text", placeholder, w, suffix, small, step, style: sx }) {
  return (<div style={{ display: "flex", flexDirection: "column", gap: 3, width: w || "100%", ...sx }}>
    {label && <label style={{ fontSize: 10, fontWeight: 700, color: P.muted, textTransform: "uppercase", letterSpacing: "0.06em", fontFamily: fs }}>{label}</label>}
    <div style={{ position: "relative" }}>
      <input type={type} value={value ?? ""} step={step} onChange={e => onChange(type === "number" ? (e.target.value === "" ? "" : Number(e.target.value)) : e.target.value)} placeholder={placeholder}
        style={{ width: "100%", padding: small ? "5px 8px" : "8px 11px", paddingRight: suffix ? 30 : 11, border: "1px solid " + P.border, borderRadius: 6, fontSize: small ? 12 : 13, fontFamily: fs, background: P.input, color: P.text, outline: "none", boxSizing: "border-box" }}
        onFocus={e => e.target.style.borderColor = P.accent} onBlur={e => e.target.style.borderColor = P.border} />
      {suffix && <span style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", fontSize: 11, color: P.muted, fontFamily: fs }}>{suffix}</span>}
    </div>
  </div>);
}
function Sel({ label, value, onChange, options, small, w }) {
  return (<div style={{ display: "flex", flexDirection: "column", gap: 3, width: w || "100%" }}>
    {label && <label style={{ fontSize: 10, fontWeight: 700, color: P.muted, textTransform: "uppercase", letterSpacing: "0.06em", fontFamily: fs }}>{label}</label>}
    <select value={value} onChange={e => onChange(e.target.value)} style={{ padding: small ? "5px 8px" : "8px 11px", border: "1px solid " + P.border, borderRadius: 6, fontSize: small ? 12 : 13, fontFamily: fs, background: P.input, color: P.text, cursor: "pointer" }}>
      {options.map(o => <option key={o.v} value={o.v}>{o.l}</option>)}
    </select>
  </div>);
}
function B({ children, onClick, v = "primary", small, sx }) {
  const s = { primary: { background: P.accent, color: "#fff" }, secondary: { background: "transparent", color: P.accent, border: "1px solid " + P.border }, danger: { background: P.redBg, color: P.red }, gold: { background: P.gold, color: "#fff", fontWeight: 700 }, ghost: { background: "transparent", color: P.muted }, green: { background: P.green, color: "#fff" } };
  return <button onClick={onClick} style={{ padding: small ? "5px 12px" : "8px 18px", borderRadius: 6, fontSize: small ? 11 : 12, fontWeight: 600, fontFamily: fs, cursor: "pointer", border: "none", display: "inline-flex", alignItems: "center", gap: 5, whiteSpace: "nowrap", ...s[v], ...sx }}>{children}</button>;
}
function Card({ children, title, actions, sx }) {
  return (<div style={{ background: P.card, borderRadius: 10, border: "1px solid " + P.border, padding: 20, marginBottom: 14, ...sx }}>
    {(title || actions) && <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
      {title && <h3 style={{ margin: 0, fontSize: 15, fontFamily: ff, fontWeight: 700, color: P.accent }}>{title}</h3>}
      <div style={{ display: "flex", gap: 6 }}>{actions}</div>
    </div>}
    {children}
  </div>);
}
function Tabs({ items, active, onChange }) {
  return (<div style={{ display: "flex", gap: 0, borderBottom: "2px solid " + P.border, marginBottom: 18, overflowX: "auto" }}>
    {items.map(t => <button key={t.k} onClick={() => onChange(t.k)} style={{ padding: "9px 14px", fontSize: 12, fontWeight: active === t.k ? 700 : 500, fontFamily: fs, color: active === t.k ? P.accent : P.muted, background: "transparent", border: "none", borderBottom: active === t.k ? "2px solid " + P.accent : "2px solid transparent", marginBottom: -2, cursor: "pointer", whiteSpace: "nowrap" }}>{t.l}</button>)}
  </div>);
}
function Tag({ children, color = P.accent, bg = P.goldBg }) {
  return <span style={{ display: "inline-block", padding: "3px 10px", borderRadius: 14, fontSize: 10, fontWeight: 700, fontFamily: fs, color, background: bg }}>{children}</span>;
}
function SRow({ label, value, sub, bold, big, color }) {
  return (<><tr><td style={{ padding: "5px 0", fontWeight: bold ? 700 : 400, fontSize: big ? 16 : 12, color: color || P.text, fontFamily: fs }}>{label}</td>
    <td style={{ padding: "5px 0", textAlign: "right", fontWeight: bold ? 700 : 600, fontSize: big ? 20 : 13, color: color || P.text, fontFamily: fs }}>{fmt(value)}</td></tr>
    {sub && <tr><td colSpan={2} style={{ padding: "0 0 3px", fontSize: 10, color: P.muted, fontFamily: fs }}>{sub}</td></tr>}</>);
}

// ─── DECORATIONS ───
function Curtain() {
  return (<div style={{ position: "relative", width: "100%", overflow: "hidden", height: 110, background: "linear-gradient(180deg, " + P.accent + " 0%, #6B1414 100%)" }}>
    <div style={{ position: "absolute", left: -20, top: -30, width: "55%", height: 150, background: "radial-gradient(ellipse at 30% 0%, #A02020 0%, " + P.accent + " 50%, #5A1010 100%)", borderRadius: "0 0 60% 0", transform: "rotate(-2deg)", opacity: 0.9 }} />
    <div style={{ position: "absolute", right: -20, top: -30, width: "55%", height: 150, background: "radial-gradient(ellipse at 70% 0%, #A02020 0%, " + P.accent + " 50%, #5A1010 100%)", borderRadius: "0 0 0 60%", transform: "rotate(2deg)", opacity: 0.9 }} />
    <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: 5, background: "linear-gradient(180deg, #E8C84A 0%, " + P.gold + " 40%, #A07820 100%)" }} />
    <svg style={{ position: "absolute", bottom: -1, left: 0, width: "100%" }} viewBox="0 0 400 25" preserveAspectRatio="none" height="25">
      <path d="M0,0 Q25,24 50,0 Q75,24 100,0 Q125,24 150,0 Q175,24 200,0 Q225,24 250,0 Q275,24 300,0 Q325,24 350,0 Q375,24 400,0 L400,25 L0,25 Z" fill={P.bg} />
    </svg>
  </div>);
}
function Flourish({ width = 180, color = P.gold }) {
  return (<svg width={width} height={20} viewBox="0 0 200 20" style={{ display: "block", margin: "0 auto" }}>
    <path d="M20,10 Q40,2 60,8 Q80,14 100,10 Q120,6 140,12 Q160,18 180,10" fill="none" stroke={color} strokeWidth="1.2" />
    <circle cx="100" cy="10" r="2" fill={color} /><circle cx="16" cy="10" r="1.2" fill={color} /><circle cx="184" cy="10" r="1.2" fill={color} />
  </svg>);
}

// ─── LOGIN ───
function LoginScreen({ onLogin, error: extError }) {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(extError || "");
  const handleLogin = async () => {
    setLoading(true); setError("");
    const { data, error: err } = await supabase.auth.signInWithPassword({ email, password: pass });
    setLoading(false);
    if (err) { setError("Email o password non corretti"); return; }
    onLogin(data.user);
  };
  return (
    <div style={{ minHeight: "100vh", background: P.bg, display: "flex", flexDirection: "column" }}>
      <Curtain />
      <div style={{ flex: 1, display: "flex", justifyContent: "center", alignItems: "flex-start", padding: "0 20px" }}>
        <div style={{ width: 340, marginTop: 20 }}>
          <div style={{ textAlign: "center", marginBottom: 20 }}>
            <img src={LOGO} alt="Il Sipario Musicale" style={{ height: 80, marginBottom: 10 }} />
            <h1 style={{ fontFamily: ff, fontSize: 22, color: P.accent, margin: "0 0 4px" }}>Gestione Preventivi</h1>
            <Flourish width={140} />
          </div>
          <Card sx={{ padding: 24 }}>
            <I label="Email" value={email} onChange={setEmail} placeholder="nome@sipariomusicale.com" />
            <I label="Password" type="password" value={pass} onChange={setPass} placeholder="Password" style={{ marginTop: 12 }} />
            {error && <p style={{ fontFamily: fs, fontSize: 11, color: P.red, margin: "8px 0 0" }}>{error}</p>}
            <div style={{ marginTop: 18 }}>
              <B v="gold" onClick={handleLogin} sx={{ width: "100%", justifyContent: "center", padding: "10px", fontSize: 13, opacity: loading ? 0.6 : 1 }}>{loading ? "Accesso..." : "Accedi"}</B>
            </div>
          </Card>
          <p style={{ textAlign: "center", fontFamily: fs, fontSize: 10, color: P.muted, marginTop: 16 }}>
            Il Sipario Musicale S.r.l. · Via Molino delle Armi 11, Milano
          </p>
        </div>
      </div>
    </div>
  );
}

// ─── ADMIN PANEL ───
function AdminPanel({ currentUser, onClose }) {
  const [users, setUsers] = useState([]);
  const [newEmail, setNewEmail] = useState("");
  const [newPass, setNewPass] = useState("");
  const [newName, setNewName] = useState("");
  const [msg, setMsg] = useState("");

  useEffect(() => {
    supabase.from("profiles").select("*").order("created_at").then(({ data }) => { if (data) setUsers(data); });
  }, []);

  const createUser = async () => {
    if (!newEmail || !newPass) { setMsg("Inserisci email e password"); return; }
    setMsg("Creazione in corso...");
    // Note: creating users requires admin API or edge function
    // For now, we use signUp which sends a confirmation email
    const { data, error } = await supabase.auth.signUp({ email: newEmail, password: newPass, options: { data: { full_name: newName } } });
    if (error) { setMsg("Errore: " + error.message); return; }
    setMsg("Utente creato! " + newEmail);
    setNewEmail(""); setNewPass(""); setNewName("");
    // Refresh users
    const { data: updated } = await supabase.from("profiles").select("*").order("created_at");
    if (updated) setUsers(updated);
  };

  const toggleRole = async (userId, currentRole) => {
    const newRole = currentRole === "admin" ? "operator" : "admin";
    await supabase.from("profiles").update({ role: newRole }).eq("id", userId);
    setUsers(users.map(u => u.id === userId ? { ...u, role: newRole } : u));
  };

  return (<div style={{ minHeight: "100vh", background: P.bg }}>
    <div style={{ background: P.accent, padding: "10px 20px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
      <h2 style={{ margin: 0, fontFamily: ff, fontSize: 16, color: "#fff" }}>Gestione Utenti</h2>
      <B v="ghost" onClick={onClose} sx={{ color: "#fff" }}>Chiudi</B>
    </div>
    <div style={{ maxWidth: 600, margin: "0 auto", padding: "20px 16px" }}>
      <Card title="Nuovo Utente">
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
          <I label="Nome" value={newName} onChange={setNewName} small placeholder="Mario Rossi" />
          <I label="Email" value={newEmail} onChange={setNewEmail} small placeholder="mario@..." />
          <I label="Password" value={newPass} onChange={setNewPass} small placeholder="min 6 caratteri" />
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <B v="primary" small onClick={createUser}>Crea Utente</B>
          {msg && <span style={{ fontFamily: fs, fontSize: 11, color: P.muted }}>{msg}</span>}
        </div>
      </Card>
      <Card title={"Utenti (" + users.length + ")"}>
        {users.map(u => (
          <div key={u.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid " + P.border }}>
            <div>
              <strong style={{ fontFamily: fs, fontSize: 13 }}>{u.full_name || u.email}</strong>
              <p style={{ fontFamily: fs, fontSize: 10, color: P.muted, margin: 0 }}>{u.email}</p>
            </div>
            <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
              <Tag color={u.role === "admin" ? P.red : P.green} bg={u.role === "admin" ? P.redBg : P.greenBg}>{u.role}</Tag>
              <B v="ghost" small onClick={() => toggleRole(u.id, u.role)}>{u.role === "admin" ? "Rendi operatore" : "Rendi admin"}</B>
            </div>
          </div>
        ))}
      </Card>
    </div>
  </div>);
}

// ─── FOLDER TAB ───
function FolderTabs({ active, onChange, counts }) {
  const tabs = [
    { k: "planning", l: "Pianificazione", c: counts.planning, color: P.accent },
    { k: "published", l: "Pubblicati", c: counts.published, color: P.green },
    { k: "completed", l: "Conclusi", c: counts.completed, color: P.muted },
  ];
  return (<div style={{ display: "flex", gap: 6, marginBottom: 20 }}>
    {tabs.map(t => (
      <button key={t.k} onClick={() => onChange(t.k)} style={{
        flex: 1, padding: "12px 8px", borderRadius: 8, fontFamily: fs, fontSize: 12, fontWeight: active === t.k ? 700 : 500,
        background: active === t.k ? (t.k === "planning" ? P.redBg : t.k === "published" ? P.greenBg : "#F0F0F0") : P.card,
        color: active === t.k ? t.color : P.muted, border: active === t.k ? "1.5px solid " + t.color : "1px solid " + P.border,
        cursor: "pointer", textAlign: "center"
      }}>
        {t.l} <span style={{ display: "block", fontSize: 18, fontFamily: ff, fontWeight: 700, marginTop: 2 }}>{t.c}</span>
      </button>
    ))}
  </div>);
}

// ─── SUMMARY TAB ───
function SummaryTab({ trip, onUpdate }) {
  const c = calcTrip(trip);
  const pax = c.pax;
  const isComm = trip.type === "misura";

  // Collect all cost items for misura view
  const allItems = [];
  trip.legs.forEach((leg, li) => {
    const nights = nightsFrom(leg.dateFrom, leg.dateTo);
    const base = leg.roomTypes[leg.baseRoomIndex];
    if (base) allItems.push({ type: "hotel", id: "hotel-" + leg.id, label: "Hotel " + (leg.destination || "Tappa " + (li+1)) + " (" + base.name + ")", costPP: ((base.rate + base.taxRate) / base.occupancy + (base.breakfast || 0)) * nights, costTotal: ((base.rate + base.taxRate) / base.occupancy + (base.breakfast || 0)) * nights * pax, isBase: base.isBase === true, fixed: false });
    (leg.days || []).forEach(day => { (day.activities || []).forEach(act => {
      if (!act.description && !act.amount) return;
      const staffExtra = (act.staffMirror?.length || 0) * act.amount;
      const ppCost = act.costType === "group" ? (act.amount + staffExtra) / pax : act.amount;
      const totalCost = act.costType === "group" ? act.amount + staffExtra : act.amount * pax + staffExtra;
      allItems.push({ type: "activity", id: act.id, label: act.description || "Attivita", costPP: ppCost, costTotal: totalCost, isBase: act.isBase !== false, clientPrice: act.clientPrice || 0, actRef: act, costType: act.costType });
    }); });
  });
  trip.staff.forEach(s => {
    const tot = (s.totalFee || 0) + (s.totalHotel || 0) + (s.flight || 0);
    if (tot > 0) allItems.push({ type: "staff", id: s.id, label: (s.role || "Staff") + (s.name ? " — " + s.name : ""), costPP: tot / pax, costTotal: tot, isBase: s.isBase === true, fixed: false });
  });
  trip.miscCosts.forEach(m => {
    if (!m.description && !m.amount) return;
    allItems.push({ type: "misc", id: m.id, label: m.description || "Costo", costPP: (m.amount || 0) / pax, costTotal: m.amount || 0, isBase: m.isBase !== false, clientPrice: m.clientPrice || 0, miscRef: m });
  });

  // Toggle isBase for an activity
  const toggleBase = (item) => {
    if (item.type === "activity") {
      const legs = trip.legs.map(leg => ({ ...leg, days: (leg.days || []).map(day => ({ ...day, activities: day.activities.map(a => a.id === item.id ? { ...a, isBase: !a.isBase } : a) })) }));
      onUpdate({ ...trip, legs });
    } else if (item.type === "misc") {
      const miscCosts = trip.miscCosts.map(m => m.id === item.id ? { ...m, isBase: !m.isBase } : m);
      onUpdate({ ...trip, miscCosts });
    } else if (item.type === "staff") {
      const staff = trip.staff.map(s => s.id === item.id ? { ...s, isBase: !s.isBase } : s);
      onUpdate({ ...trip, staff });
    } else if (item.type === "hotel") {
      const legs = trip.legs.map(leg => {
        if ("hotel-" + leg.id === item.id) {
          const rt = [...leg.roomTypes];
          const bi = leg.baseRoomIndex;
          if (rt[bi]) rt[bi] = { ...rt[bi], isBase: !rt[bi].isBase };
          return { ...leg, roomTypes: rt };
        }
        return leg;
      });
      onUpdate({ ...trip, legs });
    }
  };
  const setClientPrice = (item, val) => {
    if (item.type === "activity") {
      const legs = trip.legs.map(leg => ({ ...leg, days: (leg.days || []).map(day => ({ ...day, activities: day.activities.map(a => a.id === item.id ? { ...a, clientPrice: val } : a) })) }));
      onUpdate({ ...trip, legs });
    } else if (item.type === "misc") {
      const miscCosts = trip.miscCosts.map(m => m.id === item.id ? { ...m, clientPrice: val } : m);
      onUpdate({ ...trip, miscCosts });
    }
  };

  // Misura calc
  const baseItems = allItems.filter(i => i.isBase);
  const extraItems = allItems.filter(i => !i.isBase);
  const baseCostTotal = baseItems.reduce((s, i) => s + i.costTotal, 0);
  const baseCostPP = baseCostTotal / pax;
  const basePPwithMarkup = baseCostPP + (trip.markupPerPerson || 0);
  let baseCommTotal = 0;
  const baseCommDetails = trip.commissions.filter(cm => cm.enabled).map(cm => { const v = basePPwithMarkup * (cm.pct / 100); baseCommTotal += v; return { ...cm, value: v }; });
  const basePPfinal = basePPwithMarkup + baseCommTotal;
  const extraTotalClient = extraItems.reduce((s, i) => s + (i.clientPrice || 0), 0);

  // General totals
  const totalCostTrip = c.subtotalPP * pax;
  const totalMarkup = (trip.markupPerPerson || 0) * pax;
  const totalComm = c.commTotal * pax;
  const totalRevenueTrip = c.rounded * pax;
  const totalMargin = totalRevenueTrip - totalCostTrip;

  return (<div>
    {/* ── DOWNLOAD ── */}
    <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
      <B v="primary" small onClick={() => generateInternalExcel(trip, c)}>Scarica Excel Interno</B>
      {!isComm && <B v="gold" small onClick={() => generateClientPdfCatalogo(trip, c)}>Genera PDF Cliente</B>}
      {isComm && <B v="green" small onClick={() => generateClientExcelMisura(trip, allItems, basePPfinal, extraItems)}>Scarica Excel Cliente</B>}
    </div>

    {/* ── RESOCONTO TOTALE VIAGGIO ── */}
    <Card title="Resoconto Totale Viaggio" sx={{ background: P.goldBg, border: "1.5px solid " + P.gold }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, textAlign: "center" }}>
        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: P.muted, fontFamily: fs, textTransform: "uppercase", marginBottom: 4 }}>Costo totale viaggio</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: P.accent, fontFamily: ff }}>{fmt(totalCostTrip)}</div>
        </div>
        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: P.muted, fontFamily: fs, textTransform: "uppercase", marginBottom: 4 }}>Ricavo totale ({pax} pax)</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: P.accent, fontFamily: ff }}>{fmt(totalRevenueTrip)}</div>
        </div>
        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: P.muted, fontFamily: fs, textTransform: "uppercase", marginBottom: 4 }}>Margine totale</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: totalMargin >= 0 ? P.green : P.red, fontFamily: fs }}>{fmt(totalMargin)}</div>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, textAlign: "center", marginTop: 12, paddingTop: 12, borderTop: "1px solid " + P.gold + "44" }}>
        <div><span style={{ fontSize: 10, color: P.muted, fontFamily: fs }}>Utile totale</span><br/><strong style={{ fontFamily: fs, fontSize: 14, color: P.green }}>{fmt(totalMarkup)}</strong></div>
        <div><span style={{ fontSize: 10, color: P.muted, fontFamily: fs }}>Commissioni totali</span><br/><strong style={{ fontFamily: fs, fontSize: 14, color: P.text }}>{fmt(totalComm)}</strong></div>
        <div><span style={{ fontSize: 10, color: P.muted, fontFamily: fs }}>Prezzo pacchetto pp</span><br/><strong style={{ fontFamily: fs, fontSize: 18, fontWeight: 700, color: P.accent }}>{fmt(c.total)}</strong></div>
      </div>
    </Card>

    {/* ── BREAKDOWN PER PERSONA ── */}
    <Card title="Breakdown a Persona">
      <table style={{ width: "100%", borderCollapse: "collapse" }}><tbody>
        {c.legBreakdowns.map((lb, i) => <SRow key={lb.leg.id} label={(lb.leg.hotelName || "Hotel") + " - " + (lb.leg.destination || "Tappa " + (i+1))} value={lb.hotelBase} sub={lb.nights + " notti"} />)}
        <SRow label="Complessivi pro-capite" value={c.groupPP} sub={"Totale " + fmt(c.totalGroupAll) + " / " + pax + " pax"} color={P.text} />
        <SRow label="Individuali" value={c.totalIndAll} />
        <tr><td colSpan={2} style={{ borderBottom: "1px solid " + P.border, padding: "5px 0" }} /></tr>
        <SRow label="Subtotale pp" value={c.subtotalPP} bold />
        <SRow label="Utile pp" value={trip.markupPerPerson} color={P.green} />
        {c.commDetails.map(cd => <SRow key={cd.id} label={cd.name + " (" + cd.pct + "%)"} value={cd.value} />)}
        <tr><td colSpan={2} style={{ borderBottom: "2px solid " + P.accent, padding: "3px 0" }} /></tr>
        <>
          <tr><td colSpan={2} style={{ borderBottom: "2px solid " + P.accent, padding: "3px 0" }} /></tr>
          <tr>
            <td style={{ padding: "10px 0", fontSize: 16, fontWeight: 700, color: P.accent, fontFamily: fs }}>PACCHETTO PP</td>
            <td style={{ padding: "10px 0", textAlign: "right", fontSize: 24, fontWeight: 700, color: P.accent, fontFamily: fs }}>{fmt(c.total)}</td>
          </tr>
          </>

      </tbody></table>
    </Card>

    {/* ── COMMISSIONE: BASE vs EXTRA ── */}
    {isComm && <Card title="Composizione Pacchetto Su Misura" sx={{ border: "1.5px solid " + P.green }}>
      <p style={{ fontFamily: fs, fontSize: 11, color: P.muted, margin: "0 0 12px" }}>Seleziona i costi che fanno parte del pacchetto base. I restanti saranno extra con prezzo finale a tua scelta.</p>
      <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: fs, fontSize: 12 }}>
        <thead><tr>
          <th style={{ textAlign: "center", padding: "6px 4px", fontSize: 10, color: P.muted, fontWeight: 700, borderBottom: "1.5px solid " + P.border, width: 40 }}>Base</th>
          <th style={{ textAlign: "left", padding: "6px 4px", fontSize: 10, color: P.muted, fontWeight: 700, borderBottom: "1.5px solid " + P.border }}>Voce</th>
          <th style={{ textAlign: "right", padding: "6px 4px", fontSize: 10, color: P.muted, fontWeight: 700, borderBottom: "1.5px solid " + P.border }}>Costo totale</th>
          <th style={{ textAlign: "right", padding: "6px 4px", fontSize: 10, color: P.muted, fontWeight: 700, borderBottom: "1.5px solid " + P.border }}>Costo pp</th>
          <th style={{ textAlign: "right", padding: "6px 4px", fontSize: 10, color: P.muted, fontWeight: 700, borderBottom: "1.5px solid " + P.border, width: 120 }}>Prezzo cliente</th>
        </tr></thead>
        <tbody>
          {allItems.map(item => (
            <tr key={item.id} style={{ borderBottom: "1px solid " + P.border, background: item.isBase ? P.greenBg : "transparent" }}>
              <td style={{ padding: "6px 4px", textAlign: "center" }}>
                <input type="checkbox" checked={item.isBase} onChange={() => toggleBase(item)} style={{ accentColor: P.green, width: 16, height: 16 }} />
              </td>
              <td style={{ padding: "6px 4px", fontWeight: item.isBase ? 600 : 400 }}>{item.label}</td>
              <td style={{ padding: "6px 4px", textAlign: "right", fontSize: 11, color: P.muted }}>{fmt(item.costTotal)}</td>
              <td style={{ padding: "6px 4px", textAlign: "right" }}>{fmt(item.costPP)}</td>
              <td style={{ padding: "6px 4px", textAlign: "right" }}>
                {!item.isBase ? <input type="number" value={item.clientPrice || ""} onChange={e => setClientPrice(item, e.target.value === "" ? "" : Number(e.target.value))}
                  style={{ width: 90, padding: "3px 6px", border: "1px solid " + P.border, borderRadius: 4, fontSize: 12, fontFamily: fs, textAlign: "right" }} placeholder="€ cliente" /> :
                  <span style={{ color: P.muted, fontSize: 10 }}>incluso</span>}
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {/* Misura summary */}
      <div style={{ marginTop: 16, padding: 14, background: P.card, borderRadius: 8, border: "1px solid " + P.border }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div>
            <div style={{ fontSize: 10, fontWeight: 700, color: P.green, fontFamily: fs, textTransform: "uppercase", marginBottom: 6 }}>Pacchetto Base</div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, fontFamily: fs }}><tbody>
              <SRow label="Costo base pp" value={baseCostPP} />
              <SRow label="Utile pp" value={trip.markupPerPerson} color={P.green} />
              {baseCommDetails.map(cd => <SRow key={cd.id} label={cd.name + " " + cd.pct + "%" } value={cd.value} />)}
              <tr><td colSpan={2} style={{ borderBottom: "2px solid " + P.green, padding: "3px 0" }} /></tr>
              <SRow label="Prezzo base al cliente" value={round5(basePPfinal)} bold big color={P.accent} />
            </tbody></table>
          </div>
          <div>
            <div style={{ fontSize: 10, fontWeight: 700, color: P.gold, fontFamily: fs, textTransform: "uppercase", marginBottom: 6 }}>Extra</div>
            {extraItems.length === 0 ? <p style={{ fontSize: 11, color: P.muted }}>Nessun extra</p> :
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, fontFamily: fs }}><tbody>
                {extraItems.map(item => <tr key={item.id} style={{ borderBottom: "1px solid " + P.border }}>
                  <td style={{ padding: "4px 0", fontSize: 11 }}>{item.label}</td>
                  <td style={{ padding: "4px 0", textAlign: "right", fontWeight: 600, color: P.gold }}>{item.clientPrice ? fmt(item.clientPrice) : "da definire"}</td>
                </tr>)}
                <tr><td colSpan={2} style={{ borderBottom: "2px solid " + P.gold, padding: "3px 0" }} /></tr>
                <SRow label="Totale extra" value={extraTotalClient} bold color={P.gold} />
              </tbody></table>}
          </div>
        </div>
      </div>
    </Card>}

    {/* ── SUPPLEMENTI ── */}
    {c.supplements.length > 0 && <Card title="Supplementi Camera">
      <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: fs, fontSize: 12 }}>
        <thead><tr>{["Camera","Diff/notte","Totale","Venduto a"].map(h => <th key={h} style={{ textAlign: h === "Camera" ? "left" : "right", padding: "5px 4px", fontSize: 10, color: P.muted, fontWeight: 700, textTransform: "uppercase", borderBottom: "1px solid " + P.border }}>{h}</th>)}</tr></thead>
        <tbody>{c.supplements.map((s,i) => <tr key={i} style={{ borderBottom: "1px solid " + P.border }}>
          <td style={{ padding: "6px 4px" }}>{s.room.name}</td>
          <td style={{ padding: "6px 4px", textAlign: "right" }}>{fmt(s.diffPerNight)}</td>
          <td style={{ padding: "6px 4px", textAlign: "right", fontWeight: 600 }}>{fmt(s.total)}</td>
          <td style={{ padding: "6px 4px", textAlign: "right", fontWeight: 700, color: P.gold }}>{fmt(s.rounded)}</td>
        </tr>)}</tbody>
      </table>
    </Card>}
  </div>);
}

// ─── COMPARISON VIEW ───
function collectItems(trip) {
  const items = [];
  const pax = trip.totalParticipants || 1;
  trip.legs.forEach((leg, li) => {
    const nights = nightsFrom(leg.dateFrom, leg.dateTo);
    const base = leg.roomTypes[leg.baseRoomIndex];
    if (base) items.push({ cat: "hotel", id: "h-" + leg.id, label: "Hotel " + (leg.destination || "Tappa " + (li+1)), totalCost: ((base.rate + base.taxRate) / base.occupancy + (base.breakfast || 0)) * nights * pax, editable: false });
    (leg.days || []).forEach(day => { (day.activities || []).forEach(act => {
      if (!act.description && !act.amount) return;
      const staffExtra = (act.staffMirror?.length || 0) * act.amount;
      const total = act.costType === "group" ? act.amount + staffExtra : act.amount * pax + staffExtra;
      items.push({ cat: "activity", id: act.id, label: act.description || "Attivita", totalCost: total, costType: act.costType, editable: true, amount: act.amount });
    }); });
  });
  trip.staff.forEach(s => {
    const tot = (s.totalFee || 0) + (s.totalHotel || 0) + (s.flight || 0);
    if (tot > 0) items.push({ cat: "staff", id: s.id, label: (s.role || "Staff") + (s.name ? " " + s.name : ""), totalCost: tot, editable: true, totalFee: s.totalFee, totalHotel: s.totalHotel, flight: s.flight });
  });
  trip.miscCosts.forEach(m => {
    if (!m.description && !m.amount) return;
    items.push({ cat: "misc", id: m.id, label: m.description || "Costo", totalCost: m.amount || 0, editable: true, amount: m.amount });
  });
  return items;
}

function CompareView({ trip, onUpdate }) {
  if (!trip.snapshot) return <p style={{ fontFamily: fs, color: P.muted, padding: 20 }}>Nessun preventivo originale salvato. Pubblica il viaggio per abilitare il confronto.</p>;

  const snap = trip.snapshot;
  const origCalc = calcTrip(snap);
  const realCalc = calcTrip(trip);
  const origItems = collectItems(snap);
  const realItems = collectItems(trip);

  const sellPrice = snap.sellPrice || origCalc.rounded;
  const origPax = snap.totalParticipants || 1;
  const realPax = trip.totalParticipants || 1;

  const origRevenue = sellPrice * origPax;
  const realRevenue = sellPrice * realPax;
  const origCostTotal = origCalc.subtotalPP * origPax;
  const realCostTotal = realCalc.subtotalPP * realPax;
  const origMarginTotal = origRevenue - origCostTotal;
  const realMarginTotal = realRevenue - realCostTotal;
  const marginDiff = realMarginTotal - origMarginTotal;

  const origMarginPP = sellPrice - origCalc.subtotalPP;
  const realMarginPP = sellPrice - realCalc.subtotalPP;

  // Edit handlers
  const setPax = (val) => onUpdate({ ...trip, totalParticipants: val });
  const editActivity = (actId, newAmount) => {
    const legs = trip.legs.map(leg => ({ ...leg, days: (leg.days || []).map(day => ({ ...day, activities: day.activities.map(a => a.id === actId ? { ...a, amount: newAmount } : a) })) }));
    onUpdate({ ...trip, legs });
  };
  const editStaff = (sId, field, val) => {
    const staff = trip.staff.map(s => s.id === sId ? { ...s, [field]: val } : s);
    onUpdate({ ...trip, staff });
  };
  const editMisc = (mId, val) => {
    const miscCosts = trip.miscCosts.map(m => m.id === mId ? { ...m, amount: val } : m);
    onUpdate({ ...trip, miscCosts });
  };

  const diffColor = (diff, upIsBad = true) => {
    if (Math.abs(diff) < 0.5) return P.muted;
    if (upIsBad) return diff > 0 ? P.red : P.green;
    return diff > 0 ? P.green : P.red;
  };
  const diffFmt = (diff) => Math.abs(diff) < 0.5 ? "—" : (diff > 0 ? "+" : "") + fmt(diff);

  const th = { padding: "6px 4px", fontSize: 10, color: P.muted, fontWeight: 700, textTransform: "uppercase", borderBottom: "2px solid " + P.accent, fontFamily: fs };
  const td = { padding: "7px 4px", borderBottom: "1px solid " + P.border, fontSize: 12, fontFamily: fs };
  const numInput = { width: 80, padding: "4px 6px", border: "1px solid " + P.border, borderRadius: 4, fontSize: 12, fontFamily: fs, textAlign: "right", background: "#FFFFF0" };

  return (<div>
    {/* ── BIG PICTURE ── */}
    <Card title="Quadro Generale" sx={{ background: P.goldBg, border: "1.5px solid " + P.gold }}>
      <div style={{ marginBottom: 14 }}>
        <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 6 }}>
          <span style={{ fontSize: 11, fontFamily: fs, color: P.muted, fontWeight: 600 }}>Prezzo vendita fisso:</span>
          <span style={{ fontSize: 16, fontFamily: ff, fontWeight: 700, color: P.accent }}>{fmt(sellPrice)} / persona</span>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <span style={{ fontSize: 11, fontFamily: fs, color: P.muted, fontWeight: 600 }}>Partecipanti consuntivo:</span>
          <input type="number" value={trip.totalParticipants} onChange={e => setPax(e.target.value === "" ? "" : Number(e.target.value))}
            style={{ ...numInput, width: 60, fontSize: 14, fontWeight: 700, background: "#fff" }} />
          <span style={{ fontSize: 11, fontFamily: fs, color: P.muted }}>(preventivo: {origPax})</span>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 0, borderTop: "1px solid " + P.gold + "44" }}>
        {[
          { label: "Ricavo totale", orig: origRevenue, real: realRevenue, upIsBad: false },
          { label: "Costo totale", orig: origCostTotal, real: realCostTotal, upIsBad: true },
          { label: "MARGINE TOTALE", orig: origMarginTotal, real: realMarginTotal, highlight: true, upIsBad: false },
        ].map(({ label, orig, real, highlight, upIsBad }) => {
          const diff = real - orig;
          return (<div key={label} style={{ padding: "14px 10px", textAlign: "center", borderRight: highlight ? "none" : "1px solid " + P.gold + "33" }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: P.muted, fontFamily: fs, textTransform: "uppercase", marginBottom: 6 }}>{label}</div>
            <div style={{ fontSize: 11, color: P.muted, fontFamily: fs, marginBottom: 2 }}>Preventivo: {fmt(orig)}</div>
            <div style={{ fontSize: highlight ? 22 : 18, fontWeight: 700, color: highlight ? (diff >= 0 ? P.green : P.red) : P.text, fontFamily: fs }}>{fmt(real)}</div>
            {diff !== 0 && <div style={{ fontSize: 11, fontWeight: 600, color: diffColor(diff, upIsBad), fontFamily: fs, marginTop: 2 }}>{diffFmt(diff)}</div>}
          </div>);
        })}
      </div>

      <div style={{ display: "flex", justifyContent: "center", gap: 24, marginTop: 10, paddingTop: 10, borderTop: "1px solid " + P.gold + "44" }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 10, color: P.muted, fontFamily: fs }}>Margine pp preventivo</div>
          <div style={{ fontSize: 14, fontWeight: 700, color: P.green, fontFamily: ff }}>{fmt(origMarginPP)}</div>
        </div>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 10, color: P.muted, fontFamily: fs }}>Margine pp consuntivo</div>
          <div style={{ fontSize: 14, fontWeight: 700, color: realMarginPP >= origMarginPP ? P.green : P.red, fontFamily: ff }}>{fmt(realMarginPP)}</div>
        </div>
      </div>
    </Card>

    {/* ── DETTAGLIO COSTI ── */}
    <Card title="Dettaglio Costi — Modifica qui i valori reali">
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead><tr>
          <th style={{ ...th, textAlign: "left" }}>Voce</th>
          <th style={{ ...th, textAlign: "right" }}>Preventivo</th>
          <th style={{ ...th, textAlign: "right", background: "#FFFFF0" }}>Consuntivo</th>
          <th style={{ ...th, textAlign: "right" }}>Diff</th>
        </tr></thead>
        <tbody>
          {/* Participants row */}
          <tr style={{ background: P.cream }}>
            <td style={{ ...td, fontWeight: 700 }}>Partecipanti</td>
            <td style={{ ...td, textAlign: "right" }}>{origPax}</td>
            <td style={{ ...td, textAlign: "right", background: "#FFFFF0" }}>
              <input type="number" value={trip.totalParticipants} onChange={e => setPax(e.target.value === "" ? "" : Number(e.target.value))} style={{ ...numInput, width: 50 }} />
            </td>
            <td style={{ ...td, textAlign: "right", color: diffColor(realPax - origPax, false), fontWeight: 600 }}>{realPax - origPax !== 0 ? (realPax > origPax ? "+" : "") + (realPax - origPax) : "—"}</td>
          </tr>

          {/* Cost items */}
          {origItems.map((origItem, idx) => {
            const realItem = realItems.find(r => r.id === origItem.id) || origItem;
            const diff = realItem.totalCost - origItem.totalCost;
            const changed = Math.abs(diff) > 0.5;

            return (<tr key={origItem.id} style={{ background: changed ? (diff > 0 ? "#FEF5F5" : "#F0FAF0") : "transparent" }}>
              <td style={{ ...td, fontWeight: changed ? 600 : 400 }}>
                {origItem.label}
                <span style={{ fontSize: 9, color: P.muted, marginLeft: 6 }}>{origItem.cat === "staff" ? "staff" : origItem.costType === "group" ? "comp" : origItem.cat === "hotel" ? "hotel" : "ind"}</span>
              </td>
              <td style={{ ...td, textAlign: "right" }}>{fmt(origItem.totalCost)}</td>
              <td style={{ ...td, textAlign: "right", background: "#FFFFF0" }}>
                {origItem.editable ? (
                  origItem.cat === "activity" ? <input type="number" value={realItem.amount ?? ""} onChange={e => editActivity(origItem.id, e.target.value === "" ? "" : Number(e.target.value))} style={numInput} /> :
                  origItem.cat === "misc" ? <input type="number" value={realItem.amount ?? ""} onChange={e => editMisc(origItem.id, e.target.value === "" ? "" : Number(e.target.value))} style={numInput} /> :
                  origItem.cat === "staff" ? <div style={{ display: "flex", gap: 3, justifyContent: "flex-end" }}>
                    <input type="number" value={realItem.totalFee ?? ""} onChange={e => editStaff(origItem.id, "totalFee", e.target.value === "" ? "" : Number(e.target.value))} style={{ ...numInput, width: 55 }} title="Onorario" />
                    <input type="number" value={realItem.totalHotel ?? ""} onChange={e => editStaff(origItem.id, "totalHotel", e.target.value === "" ? "" : Number(e.target.value))} style={{ ...numInput, width: 55 }} title="Hotel" />
                    <input type="number" value={realItem.flight ?? ""} onChange={e => editStaff(origItem.id, "flight", e.target.value === "" ? "" : Number(e.target.value))} style={{ ...numInput, width: 55 }} title="Volo" />
                  </div> : fmt(realItem.totalCost)
                ) : <span style={{ color: P.muted }}>{fmt(realItem.totalCost)}</span>}
              </td>
              <td style={{ ...td, textAlign: "right", fontWeight: changed ? 700 : 400, color: diffColor(diff, true) }}>{changed ? diffFmt(diff) : "—"}</td>
            </tr>);
          })}

          {/* Totals */}
          <tr style={{ borderTop: "2px solid " + P.accent }}>
            <td style={{ ...td, fontWeight: 700, fontSize: 13 }}>COSTO TOTALE VIAGGIO</td>
            <td style={{ ...td, textAlign: "right", fontWeight: 700, fontFamily: ff, fontSize: 14 }}>{fmt(origCostTotal)}</td>
            <td style={{ ...td, textAlign: "right", fontWeight: 700, fontFamily: ff, fontSize: 14, background: "#FFFFF0" }}>{fmt(realCostTotal)}</td>
            <td style={{ ...td, textAlign: "right", fontWeight: 700, fontSize: 13, color: diffColor(realCostTotal - origCostTotal, true) }}>{diffFmt(realCostTotal - origCostTotal)}</td>
          </tr>
          <tr>
            <td style={{ ...td, fontWeight: 700, fontSize: 13 }}>RICAVO TOTALE</td>
            <td style={{ ...td, textAlign: "right", fontWeight: 700, fontFamily: ff, fontSize: 14 }}>{fmt(origRevenue)}</td>
            <td style={{ ...td, textAlign: "right", fontWeight: 700, fontFamily: ff, fontSize: 14 }}>{fmt(realRevenue)}</td>
            <td style={{ ...td, textAlign: "right", fontWeight: 700, fontSize: 13, color: diffColor(realRevenue - origRevenue, false) }}>{diffFmt(realRevenue - origRevenue)}</td>
          </tr>
          <tr style={{ background: marginDiff >= 0 ? P.greenBg : P.redBg }}>
            <td style={{ padding: "10px 4px", fontWeight: 700, fontSize: 14, fontFamily: fs }}>MARGINE TOTALE</td>
            <td style={{ padding: "10px 4px", textAlign: "right", fontWeight: 700, fontFamily: ff, fontSize: 16, color: P.green }}>{fmt(origMarginTotal)}</td>
            <td style={{ padding: "10px 4px", textAlign: "right", fontWeight: 700, fontFamily: ff, fontSize: 16, color: realMarginTotal >= origMarginTotal ? P.green : P.red }}>{fmt(realMarginTotal)}</td>
            <td style={{ padding: "10px 4px", textAlign: "right", fontWeight: 700, fontSize: 14, color: diffColor(marginDiff, false) }}>{diffFmt(marginDiff)}</td>
          </tr>
        </tbody>
      </table>
    </Card>
  </div>);
}

// ─── MAIN APP ───
export default function App() {
  const [user, setUser] = useState(null);
  const [userProfile, setUserProfile] = useState(null);
  const [trips, setTrips] = useState([]);
  const [active, setActive] = useState(null);
  const [tab, setTab] = useState("settings");
  const [folder, setFolder] = useState("planning");
  const [activeLeg, setActiveLeg] = useState(0);
  const [loaded, setLoaded] = useState(false);
  const [showNewType, setShowNewType] = useState(false);
  const [ivaLegs, setIvaLegs] = useState(false);
  const [ivaProgram, setIvaProgram] = useState(false);
  const [ivaCosts, setIvaCosts] = useState(false);
  const [showAdmin, setShowAdmin] = useState(false);

  // Auth: check session on mount
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      if (session?.user) { setUser(session.user); loadProfile(session.user.id); loadTrips(); }
      setLoaded(true);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      if (session?.user) { setUser(session.user); loadProfile(session.user.id); loadTrips(); }
      else { setUser(null); setUserProfile(null); setTrips([]); }
    });
    return () => subscription.unsubscribe();
  }, []);

  const loadProfile = async (uid) => {
    const { data } = await supabase.from("profiles").select("*").eq("id", uid).single();
    if (data) setUserProfile(data);
  };
  const loadTrips = async () => {
    const { data } = await supabase.from("trips").select("*").order("updated_at", { ascending: false });
    if (data) setTrips(data.map(row => ({ ...row.data, _dbId: row.id, _dbStatus: row.status, _createdBy: row.created_by_name })));
  };

  const saveTrip = useCallback(async (tripData) => {
    const { _dbId, _dbStatus, _createdBy, ...cleanData } = tripData;
    if (_dbId) {
      await supabase.from("trips").update({ data: cleanData, status: cleanData.status }).eq("id", _dbId);
    } else {
      const { data } = await supabase.from("trips").insert({ data: cleanData, status: cleanData.status, created_by: user?.id, created_by_name: userProfile?.full_name || user?.email || "" }).select().single();
      if (data) tripData._dbId = data.id;
    }
    return tripData;
  }, [user, userProfile]);

  const updateTrip = useCallback(async (t) => {
    setActive(t);
    setTrips(prev => prev.map(x => x.id === t.id ? t : x));
    await saveTrip(t);
  }, [saveTrip]);
  const u = (f, v) => updateTrip({ ...active, [f]: v });

  const handleLogout = async () => { await supabase.auth.signOut(); setUser(null); setUserProfile(null); setTrips([]); };

  if (!loaded) return <div style={{ display: "flex", justifyContent: "center", alignItems: "center", height: "100vh", fontFamily: fs, color: P.muted }}>Caricamento...</div>;
  if (!user) return <LoginScreen onLogin={u => { setUser(u); loadProfile(u.id); loadTrips(); }} />;
  if (showAdmin) return <AdminPanel currentUser={userProfile} onClose={() => setShowAdmin(false)} />;

  // ─── HOME ───
  if (!active) {
    const planning = trips.filter(t => t.status === "planning");
    const published = trips.filter(t => t.status === "published");
    const completed = trips.filter(t => t.status === "completed");
    const filtered = folder === "planning" ? planning : folder === "published" ? published : completed;

    const createTrip = async (type) => {
      const t = mkTrip(type);
      t.createdBy = userProfile?.full_name || user?.email || "";
      const saved = await saveTrip(t);
      setTrips(prev => [...prev, saved]);
      setActive(saved); setTab("settings"); setActiveLeg(0); setShowNewType(false);
    };
    const duplicate = async (t) => {
      const { _dbId, _dbStatus, _createdBy, ...cleanData } = t;
      const nt = { ...JSON.parse(JSON.stringify(cleanData)), id: uid(), name: (t.name || "") + " (copia)", status: "planning", snapshot: null, createdBy: userProfile?.full_name || "" };
      const saved = await saveTrip(nt);
      setTrips(prev => [...prev, saved]);
    };
    const publish = async (t) => {
      const c = calcTrip(t);
      const updated = { ...t, status: "published", sellPrice: c.rounded, snapshot: JSON.parse(JSON.stringify({ ...t, sellPrice: c.rounded })) };
      await saveTrip(updated);
      setTrips(prev => prev.map(x => x.id === t.id ? updated : x));
    };
    const complete = async (t) => {
      const updated = { ...t, status: "completed" };
      await saveTrip(updated);
      setTrips(prev => prev.map(x => x.id === t.id ? updated : x));
    };
    const del = async (id) => {
      const trip = trips.find(x => x.id === id);
      if (trip?._dbId) await supabase.from("trips").delete().eq("id", trip._dbId);
      setTrips(prev => prev.filter(x => x.id !== id));
    };

    return (
      <div style={{ minHeight: "100vh", background: P.bg, display: "flex", flexDirection: "column" }}>
        <Curtain />
        <div style={{ flex: 1, maxWidth: 680, margin: "0 auto", padding: "0 20px 32px", width: "100%" }}>
          <div style={{ textAlign: "center", padding: "16px 20px 12px" }}>
            <img src={LOGO} alt="Il Sipario Musicale" style={{ height: 80, marginBottom: 8 }} />
            <h1 style={{ fontFamily: ff, fontSize: 24, color: P.accent, margin: "0 0 2px" }}>Gestione Preventivi</h1>
            <Flourish width={140} />
            <p style={{ fontFamily: fs, fontSize: 11, color: P.muted, margin: "6px 0 0" }}>{userProfile?.full_name || user?.email}</p>
            <div style={{ display: "flex", gap: 8, justifyContent: "center", marginTop: 8 }}>
              {userProfile?.role === "admin" && <B v="secondary" small onClick={() => setShowAdmin(true)}>Gestione Utenti</B>}
              <B v="ghost" small onClick={handleLogout}>Esci</B>
            </div>
          </div>

          <FolderTabs active={folder} onChange={setFolder} counts={{ planning: planning.length, published: published.length, completed: completed.length }} />

          {folder === "planning" && <div style={{ display: "flex", justifyContent: "center", marginBottom: 16 }}>
            {!showNewType ? <B v="gold" onClick={() => setShowNewType(true)}>+ Nuovo Preventivo</B> :
              <Card sx={{ padding: 16, textAlign: "center", width: "100%" }}>
                <p style={{ fontFamily: fs, fontSize: 12, color: P.muted, margin: "0 0 12px" }}>Che tipo di viaggio?</p>
                <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
                  <B v="primary" onClick={() => createTrip("catalogo")}>Catalogo</B>
                  <B v="green" onClick={() => createTrip("misura")}>Su misura</B>
                  <B v="ghost" small onClick={() => setShowNewType(false)}>Annulla</B>
                </div>
              </Card>}
          </div>}

          {filtered.length === 0 && <div style={{ textAlign: "center", padding: 32, background: P.card, borderRadius: 10, border: "1px dashed " + P.border, fontFamily: fs, color: P.muted, fontSize: 12 }}>
            {folder === "planning" ? "Nessun viaggio in pianificazione" : folder === "published" ? "Nessun viaggio pubblicato" : "Nessun viaggio concluso"}
          </div>}

          {filtered.map(t => (
            <div key={t.id} style={{ background: P.card, borderRadius: 8, border: "1px solid " + P.border, padding: "12px 16px", marginBottom: 8, cursor: "pointer" }}
              onClick={() => { setActive(t); setTab("settings"); setActiveLeg(0); }}
              onMouseEnter={e => e.currentTarget.style.borderColor = P.accent} onMouseLeave={e => e.currentTarget.style.borderColor = P.border}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 2 }}>
                    <Tag color={t.type === "catalogo" ? P.accent : P.green} bg={t.type === "catalogo" ? P.redBg : P.greenBg}>{t.type === "catalogo" ? "Catalogo" : "Su misura"}</Tag>
                    <h3 style={{ margin: 0, fontFamily: ff, fontSize: 15, color: P.accent }}>{t.name || "Senza titolo"}</h3>
                  </div>
                  <p style={{ margin: 0, fontSize: 10, fontFamily: fs, color: P.muted }}>{t.totalParticipants} pax · {t.legs?.map(l => l.destination).filter(Boolean).join(" → ") || "—"} · <em>{t.createdBy || t._createdBy || ""}</em></p>
                </div>
                <div style={{ display: "flex", gap: 4 }} onClick={e => e.stopPropagation()}>
                  {t.status === "planning" && <><B v="secondary" small onClick={() => duplicate(t)}>Duplica</B><B v="green" small onClick={() => publish(t)}>Pubblica</B></>}
                  {t.status === "published" && <B v="ghost" small onClick={() => complete(t)}>Concludi</B>}
                  <B v="danger" small onClick={() => del(t.id)}>Elimina</B>
                </div>
              </div>
            </div>
          ))}
        </div>
        <div style={{ borderTop: "1px solid " + P.border, padding: "12px 20px", textAlign: "center", background: P.card }}>
          <Flourish width={100} color={P.border} />
          <p style={{ fontFamily: fs, fontSize: 9, color: P.muted, margin: "4px 0 0" }}>Il Sipario Musicale S.r.l. · Via Molino delle Armi 11, Milano · P.IVA 01044750675</p>
        </div>
      </div>
    );
  }

  // ─── EDITOR ───
  const trip = active;
  const editorTabs = [
    { k: "settings", l: "Impostazioni" }, { k: "legs", l: "Hotel" }, { k: "program", l: "Programma" },
    { k: "costs", l: "Staff & Costi" }, { k: "participants", l: "Partecipanti" }, { k: "summary", l: "Riepilogo" },
    ...(trip.status === "published" ? [{ k: "compare", l: "Confronto" }] : []),
  ];

  return (
    <div style={{ minHeight: "100vh", background: P.bg }}>
      <div style={{ background: P.accent, position: "sticky", top: 0, zIndex: 10 }}>
        <div style={{ padding: "9px 20px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <button onClick={() => setActive(null)} style={{ background: "none", border: "none", color: "#ffffffaa", cursor: "pointer", fontSize: 13, fontFamily: fs }}>← Indietro</button>
            <div style={{ width: 1, height: 16, background: "rgba(255,255,255,0.2)" }} />
            <h2 style={{ margin: 0, fontFamily: ff, fontSize: 15, color: "#fff" }}>{trip.name || "Nuovo Preventivo"}</h2>
          </div>
          <Tag color="#fff" bg={trip.type === "catalogo" ? "rgba(255,255,255,.15)" : "rgba(46,125,50,.6)"}>{trip.type === "catalogo" ? "Catalogo" : "Su misura"}</Tag>
        </div>
      </div>

      <div style={{ maxWidth: 920, margin: "0 auto", padding: "14px 16px" }}>
        <Tabs items={editorTabs} active={tab} onChange={setTab} />

        {tab === "settings" && <Card title="Dettagli Viaggio">
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <I label="Nome viaggio" value={trip.name} onChange={v => u("name", v)} placeholder="es. Vienna" />
            <I label="Partecipanti" type="number" value={trip.totalParticipants} onChange={v => u("totalParticipants", v)} />
            <I label="Data inizio" type="date" value={trip.dateFrom} onChange={v => u("dateFrom", v)} />
            <I label="Data fine" type="date" value={trip.dateTo} onChange={v => u("dateTo", v)} />
            <I label="Utile per persona" type="number" value={trip.markupPerPerson} onChange={v => u("markupPerPerson", v)} suffix="€" />
            <div />
          </div>
          <div style={{ marginTop: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <span style={{ fontSize: 10, fontWeight: 700, color: P.muted, textTransform: "uppercase", fontFamily: fs }}>Commissioni</span>
              <B v="secondary" small onClick={() => u("commissions", [...trip.commissions, mkCommission()])}>+ Commissione</B>
            </div>
            {trip.commissions.map((c, ci) => (
              <div key={c.id} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "center" }}>
                <label style={{ display: "flex", alignItems: "center", gap: 4, cursor: "pointer" }}>
                  <input type="checkbox" checked={c.enabled} onChange={e => { const n = [...trip.commissions]; n[ci] = { ...c, enabled: e.target.checked }; u("commissions", n); }} style={{ accentColor: P.accent }} />
                </label>
                <I value={c.name} onChange={v => { const n = [...trip.commissions]; n[ci] = { ...c, name: v }; u("commissions", n); }} small />
                <I type="number" value={c.pct} onChange={v => { const n = [...trip.commissions]; n[ci] = { ...c, pct: v }; u("commissions", n); }} small suffix="%" w="90px" step="0.1" />
                <B v="ghost" small onClick={() => u("commissions", trip.commissions.filter((_, j) => j !== ci))}>Rimuovi</B>
              </div>
            ))}
          </div>
        </Card>}

        {tab === "legs" && <div>
          <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 10 }}><IvaToggle active={ivaLegs} onChange={setIvaLegs} /></div>
          {trip.legs.map((leg, i) => {
            const uL = (f, v2) => { const ls = [...trip.legs]; ls[i] = { ...leg, [f]: v2 }; u("legs", ls); };
            const uR = (ri, f, v2) => { const n = [...leg.roomTypes]; n[ri] = { ...n[ri], [f]: v2 }; uL("roomTypes", n); };
            const nights = nightsFrom(leg.dateFrom, leg.dateTo);
            return (<Card key={leg.id} title={"Tappa " + (i+1) + ": " + (leg.destination || "Nuova")} actions={<B v="danger" small onClick={() => u("legs", trip.legs.filter((_, j) => j !== i))}>Elimina</B>}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
                <I label="Destinazione" value={leg.destination} onChange={v2 => uL("destination", v2)} small />
                <I label="Hotel" value={leg.hotelName} onChange={v2 => uL("hotelName", v2)} small />
                <I label="Dal" type="date" value={leg.dateFrom} onChange={v2 => uL("dateFrom", v2)} small />
                <I label="Al" type="date" value={leg.dateTo} onChange={v2 => uL("dateTo", v2)} small />
              </div>
              {nights > 0 && <Tag bg={P.greenBg} color={P.green}>{nights} notti</Tag>}
              <div style={{ marginTop: 10, fontSize: 10, fontWeight: 700, color: P.muted, fontFamily: fs, textTransform: "uppercase", marginBottom: 6 }}>Camere <span style={{ fontWeight: 400, textTransform: "none", fontSize: 9 }}>(€/notte · tax · colaz/pp · pax)</span></div>
              {leg.roomTypes.map((r, ri) => (
                <div key={r.id} style={{ display: "grid", gridTemplateColumns: "1fr 80px 65px 65px 50px 60px 30px" + (ivaLegs ? " auto" : ""), gap: 5, marginBottom: 4, alignItems: "center" }}>
                  <I value={r.name} onChange={v2 => uR(ri, "name", v2)} small placeholder="Nome" />
                  <I type="number" value={r.rate} onChange={v2 => uR(ri, "rate", v2)} small suffix="€" />
                  <I type="number" value={r.taxRate} onChange={v2 => uR(ri, "taxRate", v2)} small suffix="tax" />
                  <I type="number" value={r.breakfast} onChange={v2 => uR(ri, "breakfast", v2)} small suffix="col" />
                  <Sel value={r.occupancy} onChange={v2 => uR(ri, "occupancy", Number(v2))} small options={[{v:1,l:"1"},{v:2,l:"2"},{v:3,l:"3"},{v:4,l:"4"}]} />
                  <button onClick={() => uL("baseRoomIndex", ri)} style={{ padding: "4px 6px", borderRadius: 5, fontSize: 9, fontWeight: 700, fontFamily: fs, cursor: "pointer", border: leg.baseRoomIndex === ri ? "none" : "1px solid " + P.border, background: leg.baseRoomIndex === ri ? P.accent : "transparent", color: leg.baseRoomIndex === ri ? "#fff" : P.muted }}>{leg.baseRoomIndex === ri ? "BASE" : "base"}</button>
                  <B v="ghost" small onClick={() => uL("roomTypes", leg.roomTypes.filter((_, j) => j !== ri))}>x</B>
                  {ivaLegs && <IvaInput value={r.rate} onChange={v2 => uR(ri, "rate", v2)} small />}
                </div>
              ))}
              <B v="secondary" small onClick={() => uL("roomTypes", [...leg.roomTypes, mkRoom()])}>+ Camera</B>
            </Card>);
          })}
          <B v="primary" onClick={() => u("legs", [...trip.legs, mkLeg()])}>+ Aggiungi Tappa</B>
        </div>}

        {tab === "program" && (() => {
          const leg = trip.legs[activeLeg]; if (!leg) return <p style={{ fontFamily: fs, color: P.muted }}>Nessuna tappa.</p>;
          const uLeg = (nl) => { const ls = [...trip.legs]; ls[activeLeg] = nl; u("legs", ls); };
          const uDay = (di, d) => { const ds = [...leg.days]; ds[di] = d; uLeg({ ...leg, days: ds }); };
          return (<div>
            {trip.legs.length > 1 && <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
              {trip.legs.map((l, i) => <B key={l.id} v={i === activeLeg ? "primary" : "secondary"} small onClick={() => setActiveLeg(i)}>{l.destination || "Tappa " + (i+1)}</B>)}
            </div>}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
              <div style={{ fontSize: 11, fontFamily: fs, color: P.muted, padding: "6px 10px", background: P.goldBg, borderRadius: 6, flex: 1, marginRight: 8 }}>
                <strong style={{ color: P.accent }}>Complessivo</strong> = diviso tra tutti · <strong style={{ color: P.gold }}>Individuale</strong> = a persona · Flagga staff per includerlo nel costo
              </div>
              <IvaToggle active={ivaProgram} onChange={setIvaProgram} />
            </div>
            {leg.days.map((day, di) => (
              <Card key={day.id} sx={{ marginBottom: 10 }}>
                <div style={{ display: "flex", gap: 8, marginBottom: 8, alignItems: "flex-end" }}>
                  <I label="Data" type="date" value={day.date} onChange={v2 => uDay(di, { ...day, date: v2 })} w="130px" small />
                  <I label="Titolo" value={day.title} onChange={v2 => uDay(di, { ...day, title: v2 })} small placeholder="es. Escursione a Melk" />
                  <B v="danger" small onClick={() => uLeg({ ...leg, days: leg.days.filter((_, j) => j !== di) })}>Elimina</B>
                </div>
                {day.activities.map((act, ai) => {
                  const ua = (f, v2) => { const as2 = [...day.activities]; as2[ai] = { ...act, [f]: v2 }; uDay(di, { ...day, activities: as2 }); };
                  return (<div key={act.id} style={{ display: "grid", gridTemplateColumns: "1fr 95px 100px" + (ivaProgram ? " auto" : "") + " auto 20px 24px", gap: 5, padding: "4px 0", borderBottom: "1px solid " + P.border, alignItems: "center" }}>
                    <I value={act.description} onChange={v2 => ua("description", v2)} small placeholder="Attivita..." />
                    <Sel value={act.costType} onChange={v2 => ua("costType", v2)} small options={[{v:"individual",l:"Individuale"},{v:"group",l:"Complessivo"}]} />
                    <I type="number" value={act.amount} onChange={v2 => ua("amount", v2)} small suffix="€" />
                    {ivaProgram && <IvaInput value={act.amount} onChange={v2 => ua("amount", v2)} />}
                    <div style={{ display: "flex", gap: 3, flexWrap: "wrap" }}>
                      {trip.staff.map(s => { const on = (act.staffMirror || []).includes(s.id);
                        return <label key={s.id} style={{ fontSize: 9, fontFamily: fs, color: on ? P.accent : P.muted, display: "flex", alignItems: "center", gap: 2, cursor: "pointer" }}>
                          <input type="checkbox" checked={on} onChange={() => ua("staffMirror", on ? act.staffMirror.filter(x => x !== s.id) : [...(act.staffMirror || []), s.id])} style={{ accentColor: P.accent, width: 11, height: 11 }} />{s.role}
                        </label>; })}
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 1 }}>
                      {ai > 0 && <button onClick={() => { const a = [...day.activities]; [a[ai-1], a[ai]] = [a[ai], a[ai-1]]; uDay(di, { ...day, activities: a }); }} style={{ border: "none", background: "transparent", cursor: "pointer", fontSize: 9, color: P.muted, padding: "0 3px" }}>▲</button>}
                      {ai < day.activities.length - 1 && <button onClick={() => { const a = [...day.activities]; [a[ai], a[ai+1]] = [a[ai+1], a[ai]]; uDay(di, { ...day, activities: a }); }} style={{ border: "none", background: "transparent", cursor: "pointer", fontSize: 9, color: P.muted, padding: "0 3px" }}>▼</button>}
                    </div>
                    <B v="ghost" small onClick={() => uDay(di, { ...day, activities: day.activities.filter((_, j) => j !== ai) })}>x</B>
                  </div>);
                })}
                <B v="secondary" small onClick={() => uDay(di, { ...day, activities: [...day.activities, mkAct()] })} sx={{ marginTop: 6 }}>+ Attivita</B>
              </Card>
            ))}
            <B v="primary" onClick={() => uLeg({ ...leg, days: [...leg.days, mkDay()] })}>+ Aggiungi Giorno</B>
          </div>);
        })()}

        {tab === "costs" && <div>
          <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 10 }}><IvaToggle active={ivaCosts} onChange={setIvaCosts} /></div>
          <Card title="Staff / Relatori">
            {trip.staff.map((s, i) => {
              const us = (f, v2) => { const n = [...trip.staff]; n[i] = { ...s, [f]: v2 }; u("staff", n); };
              const tot = (s.totalFee || 0) + (s.totalHotel || 0) + (s.flight || 0);
              return (<div key={s.id} style={{ marginBottom: 8, padding: 12, background: P.cream, borderRadius: 6, border: "1px solid " + P.border }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", gap: 8 }}>
                  <I label="Ruolo" value={s.role} onChange={v2 => us("role", v2)} small />
                  <I label="Nome" value={s.name} onChange={v2 => us("name", v2)} small />
                  <div><I label="Onorario totale" type="number" value={s.totalFee} onChange={v2 => us("totalFee", v2)} small suffix="€" />
                    {ivaCosts && <div style={{ marginTop: 3 }}><IvaInput value={s.totalFee} onChange={v2 => us("totalFee", v2)} /></div>}</div>
                  <div><I label="Hotel totale" type="number" value={s.totalHotel} onChange={v2 => us("totalHotel", v2)} small suffix="€" />
                    {ivaCosts && <div style={{ marginTop: 3 }}><IvaInput value={s.totalHotel} onChange={v2 => us("totalHotel", v2)} /></div>}</div>
                  <div><I label="Volo" type="number" value={s.flight} onChange={v2 => us("flight", v2)} small suffix="€" />
                    {ivaCosts && <div style={{ marginTop: 3 }}><IvaInput value={s.flight} onChange={v2 => us("flight", v2)} /></div>}</div>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
                  <span style={{ fontSize: 11, fontFamily: fs, color: P.muted }}>Totale: <strong style={{ color: P.accent }}>{fmt(tot)}</strong></span>
                  <B v="ghost" small onClick={() => u("staff", trip.staff.filter((_, j) => j !== i))}>Rimuovi</B>
                </div>
              </div>);
            })}
            <B v="secondary" small onClick={() => u("staff", [...trip.staff, mkStaff()])}>+ Staff</B>
          </Card>
          <Card title="Costi Vari">
            {trip.miscCosts.map((m, i) => (
              <div key={m.id} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "center" }}>
                <I value={m.description} onChange={v2 => { const n = [...trip.miscCosts]; n[i] = { ...m, description: v2 }; u("miscCosts", n); }} small />
                <Sel value={m.costType || "group"} onChange={v2 => { const n = [...trip.miscCosts]; n[i] = { ...m, costType: v2 }; u("miscCosts", n); }} small options={[{v:"group",l:"Complessivo"},{v:"individual",l:"Individuale"}]} w="110px" />
                <I type="number" value={m.amount} onChange={v2 => { const n = [...trip.miscCosts]; n[i] = { ...m, amount: v2 }; u("miscCosts", n); }} small suffix="€" w="120px" />
                {ivaCosts && <IvaInput value={m.amount} onChange={v2 => { const n = [...trip.miscCosts]; n[i] = { ...m, amount: v2 }; u("miscCosts", n); }} />}
                <B v="ghost" small onClick={() => u("miscCosts", trip.miscCosts.filter((_, j) => j !== i))}>Rimuovi</B>
              </div>
            ))}
            <B v="secondary" small onClick={() => u("miscCosts", [...trip.miscCosts, mkMisc()])}>+ Voce</B>
          </Card>
        </div>}

        {tab === "participants" && (<div>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 14 }}>
            <span style={{ fontSize: 12, fontFamily: fs, color: P.muted }}>{(trip.participants || []).length} partecipanti</span>
            <B v="primary" small onClick={() => u("participants", [...(trip.participants || []), mkPart()])}>+ Partecipante</B>
          </div>
          {(!trip.participants || trip.participants.length === 0) && <div style={{ textAlign: "center", padding: 32, background: P.card, borderRadius: 10, border: "1px dashed " + P.border, fontFamily: fs, color: P.muted, fontSize: 12 }}>Nessun partecipante</div>}
          {(trip.participants || []).map((p, pi) => {
            const up = (f, v2) => { const n = [...trip.participants]; n[pi] = { ...p, [f]: v2 }; u("participants", n); };
            return (
              <Card key={p.id} sx={{ padding: 12, marginBottom: 6 }}>
                <div style={{ display: "flex", gap: 8, alignItems: "flex-end", flexWrap: "wrap" }}>
                  <I label="Nome" value={p.firstName} onChange={v2 => up("firstName", v2)} w="130px" small />
                  <I label="Cognome" value={p.lastName} onChange={v2 => up("lastName", v2)} w="130px" small />
                  {trip.legs.map(leg => (
                    <Sel key={leg.id} label={"Camera " + (leg.destination || "Tappa")}
                      value={p.roomChoices?.[leg.id] || leg.roomTypes[leg.baseRoomIndex]?.id || ""}
                      onChange={v2 => up("roomChoices", { ...p.roomChoices, [leg.id]: v2 })}
                      options={leg.roomTypes.map(r => ({ v: r.id, l: r.name }))} small w="130px" />
                  ))}
                  <B v="danger" small onClick={() => u("participants", trip.participants.filter((_, j) => j !== pi))}>Elimina</B>
                </div>
              </Card>
            );
          })}
        </div>)}

        {tab === "summary" && <SummaryTab trip={trip} onUpdate={updateTrip} />}
        {tab === "compare" && <CompareView trip={trip} onUpdate={updateTrip} />}
      </div>
    </div>
  );
}
