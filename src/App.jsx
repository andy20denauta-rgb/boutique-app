import React, { useState } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend, CartesianGrid } from "recharts";

const todayISO = new Date().toISOString().split("T")[0];

function excelDateToISO(val) {
  if (!val) return todayISO;
  if (typeof val === "number") {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return d.toISOString().split("T")[0];
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const dm = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (dm) {
    const y = dm[3].length === 2 ? "20" + dm[3] : dm[3];
    return `${y}-${dm[2].padStart(2,"0")}-${dm[1].padStart(2,"0")}`;
  }
  return s.slice(0, 10) || todayISO;
}

const G = "#c9a96e";

const inputSt = (w) => ({
  background: "#0b0b0e", color: "#e0e0e0", border: "1px solid #2a2a35",
  padding: "7px 10px", borderRadius: "6px", width: w || "100%",
  boxSizing: "border-box", outline: "none", fontSize: "13px",
});

const btnGold = {
  background: G, color: "#000", border: "none",
  padding: "10px 20px", borderRadius: "8px", fontWeight: "700",
  cursor: "pointer", fontSize: "13px",
};
const btnGray = {
  background: "#2a2a35", color: "#ccc", border: "none",
  padding: "10px 20px", borderRadius: "8px", cursor: "pointer", fontSize: "13px",
};

const StatBox = ({ label, val, color, sub }) => (
  <div style={{ background: "#13131a", padding: "20px 22px", borderRadius: "14px", borderLeft: `4px solid ${color}` }}>
    <div style={{ fontSize: "10px", color: "#555", letterSpacing: "1.5px", marginBottom: "8px" }}>{label}</div>
    <div style={{ fontSize: "22px", fontWeight: "800", color, fontFamily: "monospace" }}>{val}</div>
    {sub && <div style={{ fontSize: "11px", color: "#555", marginTop: "4px" }}>{sub}</div>}
  </div>
);

export default function App() {
  const [rates, setRates] = useState({ usd: 540, eur: 610, transfe: 1.10 });
  const [activeTab, setActiveTab] = useState("Dashboard");

  const [data, setData] = useState({
    inventory: [],
    transactions: [],
    balance: { usd: 0, efectivo: 0, transfe: 0, euro: 0 },
    dailyHistory: [],
    messengers: ["JORGE","MIGUEL","ROYLAN","CHRIS","DIANA","JONATHAN","LEO","FERDINANDO","JUAN CARLOS"],
    activeOrders: [],
  });

  // Sales state
  const [cart, setCart] = useState([]);
  const [selMessenger, setSelMessenger] = useState("");
  const [orderDest, setOrderDest] = useState("");

  // Inventory form
  const emptyProd = { codigo:"", producto:"", descripcion:"", precio:0, stockInicial:0, stockActual:0 };
  const [showAddProd, setShowAddProd] = useState(false);
  const [newProd, setNewProd] = useState(emptyProd);

  // Finance form
  const emptyTx = { fecha:todayISO, tipo:"Ingreso", descripcion:"", monto:0, moneda:"Efectivo" };
  const [showAddTx, setShowAddTx] = useState(false);
  const [newTx, setNewTx] = useState(emptyTx);

  // ── IMPORT ──────────────────────────────────────────────
  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "binary" });

      const transactions = (wb.Sheets["Hoja1"] ? XLSX.utils.sheet_to_json(wb.Sheets["Hoja1"]) : []).map(r => ({
        fecha: excelDateToISO(r["FECHA"] || r["Fecha"] || ""),
        tipo: r["TIPO"] || r["Tipo"] || "",
        descripcion: r["DESCRIPCIÓN"] || r["DESCRIPCION"] || r["Descripción"] || r["descripcion"] || "",
        monto: parseFloat(r["MONTO"] || r["Monto"] || 0),
        moneda: r["MONEDA"] || r["Moneda"] || "",
      }));

      let balance = { usd:0, efectivo:0, transfe:0, euro:0 };
      if (wb.Sheets["Hoja2"]) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets["Hoja2"]);
        const row = rows.find(r => String(r["CONCEPTO"] || r["Concepto"] || "").toUpperCase() === "BALANCE");
        if (row) balance = {
          usd: parseFloat(row["USD"] || 0),
          efectivo: parseFloat(row["EFECTIVO"] || 0),
          transfe: parseFloat(row["TRANSFERENCIA"] || row["TRANSFE"] || 0),
          euro: parseFloat(row["EURO"] || 0),
        };
      }

      const dailyHistory = (wb.Sheets["Hoja3"] ? XLSX.utils.sheet_to_json(wb.Sheets["Hoja3"]) : []).map(r => ({
        fecha: excelDateToISO(r["FECHA"] || r["Fecha"] || ""),
        usd: parseFloat(r["USD"] || 0),
        efectivo: parseFloat(r["EFECTIVO"] || 0),
        transfe: parseFloat(r["TRANSFERENCIA"] || r["TRANSFE"] || 0),
        euro: parseFloat(r["EURO"] || 0),
      })).slice(-10);

      const inventory = (wb.Sheets["Inventario"] ? XLSX.utils.sheet_to_json(wb.Sheets["Inventario"]) : []).map(r => ({
        codigo: String(r["CÓDIGO"] || r["CODIGO"] || r["Código"] || ""),
        producto: r["PRODUCTO"] || r["Producto"] || "",
        descripcion: r["DESCRIPCION"] || r["DESCRIPCIÓN"] || r["DESCRIPCION "] || r["Descripcion"] || "",
        precio: parseFloat(r["PRECIO"] || r["Precio"] || 0),
        stockInicial: parseInt(r["STOCK INICIAL"] || r["Stock Inicial"] || 0),
        stockActual: parseInt(r["STOCK ACTUAL"] || r["Stock Actual"] || 0),
      }));

      let messengers = data.messengers;
      const listsSheet = wb.Sheets["Listas"] || wb.Sheets["listas"] || wb.Sheets["Hoja6"];
      if (listsSheet) {
        const raw = XLSX.utils.sheet_to_json(listsSheet, { header: 1 });
        const extracted = raw.flat().filter(v => v && typeof v === "string" && v.trim().length > 1);
        if (extracted.length > 0) messengers = extracted;
      }

      setData(prev => ({ ...prev, inventory, transactions, balance, dailyHistory, messengers }));
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  // ── INVENTORY ────────────────────────────────────────────
  const updateInv = (codigo, field, val) =>
    setData(prev => ({ ...prev, inventory: prev.inventory.map(i => i.codigo===codigo ? {...i,[field]:val} : i) }));
  const deleteProduct = (codigo) =>
    setData(prev => ({ ...prev, inventory: prev.inventory.filter(i => i.codigo!==codigo) }));
  const saveNewProduct = () => {
    setData(prev => ({ ...prev, inventory: [...prev.inventory, {...newProd}] }));
    setNewProd(emptyProd); setShowAddProd(false);
  };

  // ── CART ─────────────────────────────────────────────────
  const addToCart = (item) => setCart(prev => {
    const idx = prev.findIndex(c => c.codigo===item.codigo);
    if (idx>=0) { const n=[...prev]; n[idx]={...n[idx],qty:n[idx].qty+1}; return n; }
    return [...prev, {...item, qty:1}];
  });
  const removeFromCart = (codigo) => setCart(prev => prev.filter(c => c.codigo!==codigo));
  const setCartQty = (codigo, qty) => {
    if (qty<=0) { removeFromCart(codigo); return; }
    setCart(prev => prev.map(c => c.codigo===codigo ? {...c,qty} : c));
  };

  const cartCUP  = cart.reduce((a,i) => a + i.precio*rates.usd*i.qty, 0);
  const cartUSD  = cart.reduce((a,i) => a + i.precio*i.qty, 0);
  const cartEUR  = cartCUP / rates.eur;
  const cartTX   = cartCUP * rates.transfe;

  // ── FINALIZE ORDER ───────────────────────────────────────
  const handleFinalizeOrder = () => {
    if (!selMessenger || cart.length===0) return;
    const order = {
      id: Date.now(), fecha: todayISO, mensajero: selMessenger, destino: orderDest,
      items: cart.map(i => ({...i, vendido:false})), estado: "En camino", totalCUP: cartCUP,
    };
    const nuevoIngreso = {
      fecha: todayISO, tipo: "Ingreso", moneda: "Efectivo", monto: cartCUP,
      descripcion: `Venta: ${cart.map(i=>`${i.producto}${i.qty>1?" x"+i.qty:""}`).join(", ")} | ${selMessenger}${orderDest?" | "+orderDest:""}`,
    };
    setData(prev => ({
      ...prev,
      inventory: prev.inventory.map(inv => { const c=cart.find(ci=>ci.codigo===inv.codigo); return c?{...inv,stockActual:Math.max(0,inv.stockActual-c.qty)}:inv; }),
      transactions: [nuevoIngreso, ...prev.transactions],
      activeOrders: [order, ...prev.activeOrders],
    }));
    setCart([]); setOrderDest(""); setSelMessenger("");
    alert("✅ Orden procesada correctamente.");
  };

  // ── MESSENGERS ───────────────────────────────────────────
  const toggleItem = (orderId, idx) => setData(prev => ({
    ...prev, activeOrders: prev.activeOrders.map(o => {
      if (o.id!==orderId) return o;
      const items = o.items.map((it,i)=>i===idx?{...it,vendido:!it.vendido}:it);
      return {...o, items, estado: items.every(it=>it.vendido)?"Completado":o.estado};
    })
  }));
  const completeOrder = (orderId) => setData(prev => ({
    ...prev, activeOrders: prev.activeOrders.map(o => o.id===orderId?{...o,estado:"Completado",items:o.items.map(i=>({...i,vendido:true}))}:o)
  }));

  // ── FINANCE ──────────────────────────────────────────────
  const saveTx = () => { setData(prev => ({...prev, transactions:[newTx,...prev.transactions]})); setNewTx(emptyTx); setShowAddTx(false); };
  const deleteTx = (idx) => setData(prev => ({...prev, transactions:prev.transactions.filter((_,i)=>i!==idx)}));

  // ── DERIVED ──────────────────────────────────────────────
  const invValueCUP   = data.inventory.reduce((a,i) => a+i.precio*i.stockActual*rates.usd, 0);
  const moneyInStreet = data.activeOrders.filter(o=>o.estado==="En camino").reduce((a,o)=>a+o.totalCUP, 0);
  const activeMess    = [...new Set(data.activeOrders.filter(o=>o.estado==="En camino").map(o=>o.mensajero))];

  // ── RENDER ───────────────────────────────────────────────
  return (
    <div style={{ display:"flex", height:"100vh", background:"#0b0b0e", color:"#e0e0e0", fontFamily:"'Segoe UI',sans-serif" }}>

      {/* SIDEBAR */}
      <aside style={{ width:"255px", minWidth:"255px", background:"#0f0f15", borderRight:"1px solid #1e1e28", display:"flex", flexDirection:"column" }}>
        <div style={{ padding:"28px 20px 16px", textAlign:"center" }}>
          <div style={{ color:G, fontWeight:"800", fontSize:"22px", letterSpacing:"4px" }}>BOUTIQUE</div>
          <div style={{ fontSize:"10px", color:"#333", letterSpacing:"2px", marginTop:"4px" }}>GESTIÓN COMERCIAL</div>
        </div>
        <nav style={{ flex:1 }}>
          {["Dashboard","Inventario","Ventas","Mensajeros","Caja y Finanzas"].map(tab => (
            <div key={tab} onClick={()=>setActiveTab(tab)} style={{
              padding:"13px 22px", cursor:"pointer",
              color: activeTab===tab ? G : "#555",
              background: activeTab===tab ? "#13131f" : "transparent",
              borderLeft: activeTab===tab ? `3px solid ${G}` : "3px solid transparent",
              fontSize:"13px", fontWeight: activeTab===tab?"600":"400",
            }}>{tab}</div>
          ))}
        </nav>
        {/* Rates */}
        <div style={{ margin:"10px", background:"#13131f", borderRadius:"10px", padding:"16px" }}>
          <div style={{ fontSize:"10px", color:G, letterSpacing:"1.5px", marginBottom:"12px" }}>TASAS DE CAMBIO</div>
          {[["USD → CUP","usd"],["EUR → CUP","eur"],["Transfe ×","transfe"]].map(([label,key])=>(
            <div key={key} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"8px" }}>
              <span style={{ fontSize:"11px", color:"#666" }}>{label}</span>
              <input type="number" value={rates[key]} onChange={e=>setRates({...rates,[key]:parseFloat(e.target.value)})}
                style={{ width:"70px", ...inputSt("70px") }} />
            </div>
          ))}
        </div>
        <div style={{ padding:"12px 14px 20px" }}>
          <label style={{ display:"block", padding:"11px", background:G, color:"#000", textAlign:"center", borderRadius:"8px", cursor:"pointer", fontWeight:"700", fontSize:"12px" }}>
            📂 Importar Excel
            <input type="file" onChange={handleImport} style={{ display:"none" }} accept=".xlsx,.xls" />
          </label>
        </div>
      </aside>

      {/* MAIN */}
      <main style={{ flex:1, overflowY:"auto", padding:"36px 40px" }}>

        {/* ══ DASHBOARD ══ */}
        {activeTab==="Dashboard" && (
          <div>
            <h2 style={{ color:G, margin:"0 0 22px", fontSize:"18px" }}>Panel Principal</h2>
            <div style={{ fontSize:"10px", color:"#444", letterSpacing:"2px", marginBottom:"10px" }}>BALANCE REAL</div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"14px", marginBottom:"22px" }}>
              <StatBox label="USD"           val={`$${Number(data.balance.usd).toLocaleString()}`}      color={G} />
              <StatBox label="Efectivo CUP"  val={`₱${Number(data.balance.efectivo).toLocaleString()}`} color="#4ade80" />
              <StatBox label="Transferencia" val={`₱${Number(data.balance.transfe).toLocaleString()}`}  color="#60a5fa" />
              <StatBox label="Euro"          val={`€${Number(data.balance.euro).toLocaleString()}`}      color="#f59e0b" />
            </div>
            <div style={{ fontSize:"10px", color:"#444", letterSpacing:"2px", marginBottom:"10px" }}>OPERACIONES</div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"14px", marginBottom:"30px" }}>
              <StatBox label="💰 En la calle (CUP)" val={`₱${moneyInStreet.toLocaleString()}`} color="#a78bfa" sub={`${data.activeOrders.filter(o=>o.estado==="En camino").length} orden(es) activa(s)`} />
              <StatBox label="📦 Valor Inventario"  val={`₱${invValueCUP.toLocaleString()}`}  color="#fb923c" sub={`${data.inventory.length} productos`} />
              <StatBox label="🛵 Mensajeros activos" val={activeMess.length}                   color="#34d399" sub={activeMess.join(", ") || "Ninguno"} />
            </div>
            {/* Daily history */}
            <div style={{ background:"#13131a", borderRadius:"14px", padding:"24px" }}>
              <h3 style={{ color:G, margin:"0 0 18px", fontSize:"13px", letterSpacing:"1px" }}>BALANCE DIARIO — ÚLTIMOS 10 DÍAS</h3>
              {data.dailyHistory.length>0 ? (
                <>
                  <ResponsiveContainer width="100%" height={200}>
                    <BarChart data={data.dailyHistory} barGap={2}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1e1e28" />
                      <XAxis dataKey="fecha" tick={{ fill:"#555", fontSize:10 }} />
                      <YAxis tick={{ fill:"#555", fontSize:10 }} />
                      <Tooltip contentStyle={{ background:"#13131a", border:"1px solid #2a2a35", fontSize:"12px" }} />
                      <Legend wrapperStyle={{ fontSize:"12px" }} />
                      <Bar dataKey="usd"      fill={G}         name="USD"            radius={[3,3,0,0]} />
                      <Bar dataKey="efectivo" fill="#4ade80"   name="Efectivo"       radius={[3,3,0,0]} />
                      <Bar dataKey="transfe"  fill="#60a5fa"   name="Transferencia"  radius={[3,3,0,0]} />
                      <Bar dataKey="euro"     fill="#f59e0b"   name="Euro"           radius={[3,3,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                  <table style={{ width:"100%", borderCollapse:"collapse", marginTop:"18px", fontSize:"13px" }}>
                    <thead>
                      <tr style={{ borderBottom:"1px solid #1e1e28" }}>
                        {[["Fecha","#888"],["USD",G],["Efectivo","#4ade80"],["Transferencia","#60a5fa"],["Euro","#f59e0b"]].map(([h,c])=>(
                          <th key={h} style={{ padding:"9px 12px", textAlign: h==="Fecha"?"left":"right", color:c, fontWeight:"600", fontSize:"11px", letterSpacing:"1px" }}>{h.toUpperCase()}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {data.dailyHistory.map((row,i)=>(
                        <tr key={i} style={{ borderBottom:"1px solid #0f0f15" }}>
                          <td style={{ padding:"9px 12px", color:"#666" }}>{row.fecha}</td>
                          <td style={{ padding:"9px 12px", textAlign:"right", color:G,        fontFamily:"monospace" }}>${Number(row.usd).toLocaleString()}</td>
                          <td style={{ padding:"9px 12px", textAlign:"right", color:"#4ade80", fontFamily:"monospace" }}>₱{Number(row.efectivo).toLocaleString()}</td>
                          <td style={{ padding:"9px 12px", textAlign:"right", color:"#60a5fa", fontFamily:"monospace" }}>₱{Number(row.transfe).toLocaleString()}</td>
                          <td style={{ padding:"9px 12px", textAlign:"right", color:"#f59e0b", fontFamily:"monospace" }}>€{Number(row.euro).toLocaleString()}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </>
              ) : (
                <div style={{ textAlign:"center", padding:"50px", color:"#2a2a35" }}>Importa el Excel para ver el historial</div>
              )}
            </div>
          </div>
        )}

        {/* ══ INVENTARIO ══ */}
        {activeTab==="Inventario" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"22px" }}>
              <h2 style={{ color:G, margin:0, fontSize:"18px" }}>Inventario</h2>
              <button style={btnGold} onClick={()=>setShowAddProd(!showAddProd)}>+ Nuevo Producto</button>
            </div>
            {showAddProd && (
              <div style={{ background:"#13131f", border:"1px solid #2a2a35", borderRadius:"12px", padding:"20px", marginBottom:"20px" }}>
                <div style={{ fontSize:"11px", color:"#555", letterSpacing:"1.5px", marginBottom:"14px" }}>NUEVO PRODUCTO</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"12px" }}>
                  {[["Código","codigo","text"],["Producto","producto","text"],["Descripción","descripcion","text"],
                    ["Precio (USD)","precio","number"],["Stock Inicial","stockInicial","number"],["Stock Actual","stockActual","number"]
                  ].map(([label,key,type])=>(
                    <div key={key}>
                      <div style={{ fontSize:"11px", color:"#666", marginBottom:"5px" }}>{label}</div>
                      <input type={type} value={newProd[key]}
                        onChange={e=>setNewProd({...newProd,[key]:type==="number"?parseFloat(e.target.value)||0:e.target.value})}
                        style={inputSt()} />
                    </div>
                  ))}
                </div>
                <div style={{ display:"flex", gap:"10px", marginTop:"16px" }}>
                  <button style={{...btnGold,background:"#4ade80"}} onClick={saveNewProduct}>✓ Guardar</button>
                  <button style={btnGray} onClick={()=>{setShowAddProd(false);setNewProd(emptyProd);}}>Cancelar</button>
                </div>
              </div>
            )}
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:"13px" }}>
                <thead>
                  <tr style={{ borderBottom:"1px solid #2a2a35" }}>
                    {["Código","Producto","Descripción","Precio USD","Stock Inicial","Stock Actual","Total CUP",""].map(h=>(
                      <th key={h} style={{ padding:"11px 12px", textAlign:"left", color:"#555", fontWeight:"600", fontSize:"11px", letterSpacing:"1px", whiteSpace:"nowrap" }}>{h.toUpperCase()}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {data.inventory.length===0 && <tr><td colSpan={8} style={{ padding:"40px", textAlign:"center", color:"#2a2a35" }}>Importa el Excel o añade productos manualmente</td></tr>}
                  {data.inventory.map(item=>(
                    <tr key={item.codigo} style={{ borderBottom:"1px solid #13131a" }}>
                      <td style={{ padding:"10px 12px", color:"#555", fontFamily:"monospace", fontSize:"12px" }}>{item.codigo}</td>
                      <td style={{ padding:"10px 12px" }}>
                        <input value={item.producto} onChange={e=>updateInv(item.codigo,"producto",e.target.value)}
                          style={{ background:"transparent", border:"none", color:"#e0e0e0", width:"100%", outline:"none", fontSize:"13px" }} />
                      </td>
                      <td style={{ padding:"10px 12px" }}>
                        <input value={item.descripcion} onChange={e=>updateInv(item.codigo,"descripcion",e.target.value)}
                          style={{ background:"transparent", border:"none", color:"#666", width:"100%", outline:"none", fontSize:"12px" }} />
                      </td>
                      <td style={{ padding:"10px 12px" }}>
                        <input type="number" value={item.precio} onChange={e=>updateInv(item.codigo,"precio",parseFloat(e.target.value)||0)}
                          style={{...inputSt("75px"),color:G}} />
                      </td>
                      <td style={{ padding:"10px 12px" }}>
                        <input type="number" value={item.stockInicial} onChange={e=>updateInv(item.codigo,"stockInicial",parseInt(e.target.value)||0)}
                          style={inputSt("60px")} />
                      </td>
                      <td style={{ padding:"10px 12px" }}>
                        <input type="number" value={item.stockActual} onChange={e=>updateInv(item.codigo,"stockActual",parseInt(e.target.value)||0)}
                          style={{...inputSt("60px"), color:item.stockActual<=2?"#f87171":"#e0e0e0"}} />
                      </td>
                      <td style={{ padding:"10px 12px", color:"#4ade80", fontFamily:"monospace", whiteSpace:"nowrap" }}>
                        ₱{(item.precio*item.stockActual*rates.usd).toLocaleString()}
                      </td>
                      <td style={{ padding:"10px 12px" }}>
                        <button onClick={()=>deleteProduct(item.codigo)}
                          style={{ background:"none", border:"none", color:"#f87171", cursor:"pointer", fontSize:"16px" }}>🗑</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ══ VENTAS ══ */}
        {activeTab==="Ventas" && (
          <div style={{ display:"grid", gridTemplateColumns:"1.3fr 0.9fr", gap:"26px" }}>
            {/* Catalogo */}
            <div style={{ background:"#13131a", borderRadius:"14px", padding:"22px" }}>
              <h3 style={{ color:G, margin:"0 0 14px", fontSize:"13px", letterSpacing:"1px" }}>CATÁLOGO</h3>
              <div style={{ maxHeight:"70vh", overflowY:"auto" }}>
                {data.inventory.filter(i=>i.stockActual>0).map(item=>(
                  <div key={item.codigo} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"12px 0", borderBottom:"1px solid #1e1e28", gap:"12px" }}>
                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ display:"flex", gap:"8px", alignItems:"center", marginBottom:"3px" }}>
                        <span style={{ color:"#555", fontSize:"11px", fontFamily:"monospace" }}>{item.codigo}</span>
                        <span style={{ fontWeight:"600", fontSize:"13px" }}>{item.producto}</span>
                      </div>
                      <div style={{ fontSize:"12px", color:"#666", marginBottom:"3px", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{item.descripcion}</div>
                      <div style={{ fontSize:"12px", display:"flex", gap:"12px" }}>
                        <span style={{ color:G }}>${item.precio} USD</span>
                        <span style={{ color:"#4ade80" }}>₱{(item.precio*rates.usd).toLocaleString()}</span>
                        <span style={{ color:"#555" }}>Stock: {item.stockActual}</span>
                      </div>
                    </div>
                    <button onClick={()=>addToCart(item)} style={{...btnGold,padding:"7px 14px",whiteSpace:"nowrap",flexShrink:0}}>+ Añadir</button>
                  </div>
                ))}
                {data.inventory.filter(i=>i.stockActual>0).length===0 && (
                  <div style={{ textAlign:"center", padding:"40px", color:"#333" }}>Sin productos con stock disponible</div>
                )}
              </div>
            </div>
            {/* Order */}
            <div style={{ background:"#13131f", borderRadius:"14px", padding:"22px", display:"flex", flexDirection:"column" }}>
              <h3 style={{ color:G, margin:"0 0 14px", fontSize:"13px", letterSpacing:"1px" }}>ORDEN</h3>
              <div style={{ flex:1, overflowY:"auto", maxHeight:"28vh", marginBottom:"12px" }}>
                {cart.length===0 ? <div style={{ textAlign:"center", padding:"24px 0", color:"#333", fontSize:"13px" }}>Sin productos añadidos</div>
                : cart.map(item=>(
                  <div key={item.codigo} style={{ display:"flex", alignItems:"center", gap:"8px", padding:"8px 0", borderBottom:"1px solid #1e1e28" }}>
                    <span style={{ flex:1, fontSize:"13px", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{item.producto}</span>
                    <button onClick={()=>setCartQty(item.codigo,item.qty-1)} style={{ background:"#2a2a35", border:"none", color:"#fff", width:"24px", height:"24px", cursor:"pointer", borderRadius:"4px", fontWeight:"700" }}>−</button>
                    <span style={{ minWidth:"22px", textAlign:"center", fontSize:"13px" }}>{item.qty}</span>
                    <button onClick={()=>setCartQty(item.codigo,item.qty+1)} style={{ background:"#2a2a35", border:"none", color:"#fff", width:"24px", height:"24px", cursor:"pointer", borderRadius:"4px", fontWeight:"700" }}>+</button>
                    <button onClick={()=>removeFromCart(item.codigo)} style={{ background:"none", border:"none", color:"#f87171", cursor:"pointer", fontSize:"15px", padding:"0 2px" }}>✕</button>
                  </div>
                ))}
              </div>

              <select value={selMessenger} onChange={e=>setSelMessenger(e.target.value)}
                style={{...inputSt(), marginBottom:"10px", color:selMessenger?"#e0e0e0":"#555"}}>
                <option value="">👤 Seleccionar Mensajero...</option>
                {data.messengers.map(m=><option key={m} value={m}>{m}</option>)}
              </select>
              <input placeholder="📍 Municipio de destino" value={orderDest} onChange={e=>setOrderDest(e.target.value)}
                style={{...inputSt(), marginBottom:"12px"}} />

              {cart.length>0 && (
                <div style={{ background:"#0b0b0e", borderRadius:"8px", padding:"14px", marginBottom:"14px" }}>
                  <div style={{ fontSize:"10px", color:"#555", letterSpacing:"1.5px", marginBottom:"10px" }}>TOTAL A PAGAR</div>
                  {[
                    {label:"Efectivo CUP",  val:`₱${cartCUP.toLocaleString()}`,      color:"#4ade80"},
                    {label:"Transferencia", val:`₱${cartTX.toLocaleString()}`,        color:"#60a5fa"},
                    {label:"USD",           val:`$${cartUSD.toFixed(2)}`,             color:G},
                    {label:"Euro",          val:`€${cartEUR.toFixed(2)}`,             color:"#f59e0b"},
                  ].map(({label,val,color})=>(
                    <div key={label} style={{ display:"flex", justifyContent:"space-between", fontSize:"13px", marginBottom:"6px" }}>
                      <span style={{ color:"#555" }}>{label}</span>
                      <span style={{ color, fontWeight:"700", fontFamily:"monospace" }}>{val}</span>
                    </div>
                  ))}
                </div>
              )}

              <button onClick={handleFinalizeOrder} disabled={!selMessenger||cart.length===0}
                style={{ width:"100%", padding:"14px", border:"none", fontWeight:"700", borderRadius:"8px", fontSize:"13px",
                  cursor:(!selMessenger||cart.length===0)?"not-allowed":"pointer",
                  background:(!selMessenger||cart.length===0)?"#2a2a35":G,
                  color:(!selMessenger||cart.length===0)?"#555":"#000",
                }}>
                {!selMessenger?"Selecciona un mensajero":cart.length===0?"Añade productos":"✅ FINALIZAR ORDEN"}
              </button>
            </div>
          </div>
        )}

        {/* ══ MENSAJEROS ══ */}
        {activeTab==="Mensajeros" && (
          <div>
            <h2 style={{ color:G, margin:"0 0 22px", fontSize:"18px" }}>Mensajeros</h2>
            {data.activeOrders.length>0 && (
              <>
                <div style={{ fontSize:"10px", color:"#444", letterSpacing:"2px", marginBottom:"14px" }}>PEDIDOS DEL DÍA</div>
                <div style={{ display:"grid", gap:"14px", marginBottom:"32px" }}>
                  {data.messengers.map(messenger=>{
                    const orders=data.activeOrders.filter(o=>o.mensajero===messenger);
                    if (orders.length===0) return null;
                    return (
                      <div key={messenger} style={{ background:"#13131a", borderRadius:"14px", padding:"18px", border:"1px solid #2a2a35" }}>
                        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"12px" }}>
                          <div style={{ color:G, fontWeight:"700", fontSize:"14px" }}>👤 {messenger}</div>
                          <span style={{ fontSize:"12px", color:"#555" }}>{orders.filter(o=>o.estado==="En camino").length} en camino</span>
                        </div>
                        {orders.map(order=>(
                          <div key={order.id} style={{ background:"#0b0b0e", borderRadius:"10px", padding:"14px", marginBottom:"8px" }}>
                            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"10px" }}>
                              <span style={{ color:"#888", fontSize:"12px" }}>📍 {order.destino||"Sin destino"} · {order.fecha}</span>
                              <span style={{ fontSize:"11px", padding:"3px 10px", borderRadius:"20px",
                                background:order.estado==="Completado"?"#14532d":"#1e3a5f",
                                color:order.estado==="Completado"?"#4ade80":"#60a5fa" }}>
                                {order.estado}
                              </span>
                            </div>
                            {order.items.map((it,idx)=>(
                              <div key={idx} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"7px 0", borderBottom:"1px solid #1a1a1a" }}>
                                <span style={{ fontSize:"13px", textDecoration:it.vendido?"line-through":"none", color:it.vendido?"#555":"#e0e0e0" }}>
                                  {it.producto}{it.qty>1?` ×${it.qty}`:""}
                                </span>
                                <label style={{ display:"flex", alignItems:"center", gap:"6px", cursor:"pointer", fontSize:"12px", color:it.vendido?"#4ade80":"#555" }}>
                                  <input type="checkbox" checked={it.vendido} onChange={()=>toggleItem(order.id,idx)} style={{ cursor:"pointer", accentColor:"#4ade80" }} />
                                  {it.vendido?"Vendido ✓":"Marcar vendido"}
                                </label>
                              </div>
                            ))}
                            {order.estado!=="Completado" && (
                              <button onClick={()=>completeOrder(order.id)}
                                style={{...btnGold,background:"#4ade80",marginTop:"12px",fontSize:"12px",padding:"8px 16px"}}>
                                ✅ Completar pedido
                              </button>
                            )}
                          </div>
                        ))}
                      </div>
                    );
                  })}
                </div>
              </>
            )}
            <div style={{ fontSize:"10px", color:"#444", letterSpacing:"2px", marginBottom:"14px" }}>TODOS LOS MENSAJEROS</div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))", gap:"12px" }}>
              {data.messengers.map(m=>{
                const cnt=data.activeOrders.filter(o=>o.mensajero===m&&o.estado==="En camino").length;
                return (
                  <div key={m} style={{ background:"#13131a", padding:"20px", borderRadius:"12px",
                    border:`1px solid ${cnt>0?G:"#1e1e28"}`, textAlign:"center" }}>
                    <div style={{ fontSize:"26px", marginBottom:"8px" }}>👤</div>
                    <div style={{ fontWeight:"700", color:G, fontSize:"12px", letterSpacing:"0.5px" }}>{m}</div>
                    <div style={{ fontSize:"11px", marginTop:"6px", color:cnt>0?"#4ade80":"#444" }}>
                      {cnt>0?`${cnt} pedido(s) activo(s)`:"Disponible"}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ══ CAJA Y FINANZAS ══ */}
        {activeTab==="Caja y Finanzas" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"22px" }}>
              <h2 style={{ color:G, margin:0, fontSize:"18px" }}>Caja y Finanzas</h2>
              <button style={btnGold} onClick={()=>setShowAddTx(!showAddTx)}>+ Añadir Movimiento</button>
            </div>
            {showAddTx && (
              <div style={{ background:"#13131f", border:"1px solid #2a2a35", borderRadius:"12px", padding:"20px", marginBottom:"20px" }}>
                <div style={{ fontSize:"11px", color:"#555", letterSpacing:"1.5px", marginBottom:"14px" }}>NUEVO MOVIMIENTO</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"12px" }}>
                  <div>
                    <div style={{ fontSize:"11px", color:"#666", marginBottom:"5px" }}>Fecha</div>
                    <input type="date" value={newTx.fecha} onChange={e=>setNewTx({...newTx,fecha:e.target.value})} style={inputSt()} />
                  </div>
                  <div>
                    <div style={{ fontSize:"11px", color:"#666", marginBottom:"5px" }}>Tipo</div>
                    <select value={newTx.tipo} onChange={e=>setNewTx({...newTx,tipo:e.target.value})} style={inputSt()}>
                      <option>Ingreso</option><option>Gasto</option>
                    </select>
                  </div>
                  <div>
                    <div style={{ fontSize:"11px", color:"#666", marginBottom:"5px" }}>Moneda</div>
                    <select value={newTx.moneda} onChange={e=>setNewTx({...newTx,moneda:e.target.value})} style={inputSt()}>
                      <option>Efectivo</option><option>USD</option><option>Euro</option><option>Transferencia</option>
                    </select>
                  </div>
                  <div style={{ gridColumn:"span 2" }}>
                    <div style={{ fontSize:"11px", color:"#666", marginBottom:"5px" }}>Descripción</div>
                    <input value={newTx.descripcion} onChange={e=>setNewTx({...newTx,descripcion:e.target.value})} style={inputSt()} />
                  </div>
                  <div>
                    <div style={{ fontSize:"11px", color:"#666", marginBottom:"5px" }}>Monto</div>
                    <input type="number" value={newTx.monto} onChange={e=>setNewTx({...newTx,monto:parseFloat(e.target.value)||0})} style={inputSt()} />
                  </div>
                </div>
                <div style={{ display:"flex", gap:"10px", marginTop:"16px" }}>
                  <button style={{...btnGold,background:"#4ade80"}} onClick={saveTx}>✓ Guardar</button>
                  <button style={btnGray} onClick={()=>{setShowAddTx(false);setNewTx(emptyTx);}}>Cancelar</button>
                </div>
              </div>
            )}
            <table style={{ width:"100%", borderCollapse:"collapse", fontSize:"13px" }}>
              <thead>
                <tr style={{ borderBottom:"1px solid #2a2a35" }}>
                  {["Fecha","Tipo","Descripción","Monto","Moneda",""].map(h=>(
                    <th key={h} style={{ padding:"11px 12px", textAlign:"left", color:"#555", fontWeight:"600", fontSize:"11px", letterSpacing:"1px" }}>{h.toUpperCase()}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {data.transactions.length===0 && <tr><td colSpan={6} style={{ padding:"40px", textAlign:"center", color:"#2a2a35" }}>Sin movimientos registrados</td></tr>}
                {data.transactions.map((row,i)=>(
                  <tr key={i} style={{ borderBottom:"1px solid #13131a" }}>
                    <td style={{ padding:"10px 12px", color:"#666", fontFamily:"monospace", fontSize:"12px" }}>{row.fecha||todayISO}</td>
                    <td style={{ padding:"10px 12px" }}>
                      <span style={{ padding:"3px 10px", borderRadius:"20px", fontSize:"11px",
                        background:row.tipo==="Gasto"?"#3b0f0f":"#0f2b1a",
                        color:row.tipo==="Gasto"?"#f87171":"#34d399" }}>
                        {row.tipo}
                      </span>
                    </td>
                    <td style={{ padding:"10px 12px", color:"#ccc" }}>{row.descripcion}</td>
                    <td style={{ padding:"10px 12px", textAlign:"right", fontFamily:"monospace", fontWeight:"700",
                      color:row.tipo==="Gasto"?"#f87171":"#34d399" }}>
                      {Number(row.monto).toLocaleString()}
                    </td>
                    <td style={{ padding:"10px 12px", color:"#666" }}>{row.moneda}</td>
                    <td style={{ padding:"10px 12px" }}>
                      <button onClick={()=>deleteTx(i)} style={{ background:"none", border:"none", color:"#f87171", cursor:"pointer", fontSize:"15px" }}>🗑</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

      </main>
    </div>
  );
}