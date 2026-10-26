<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>Sales Order - TR Channel</title>
<!-- SheetJS untuk membaca file Excel langsung di browser -->
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  *{box-sizing:border-box}
  body{font-family:system-ui,-apple-system,sans-serif;background:#f3f4f6;margin:0;padding:12px;color:#1f2937}
  .wrap{max-width:760px;margin:auto}
  h1{font-size:20px;margin:6px 0 14px;color:#111827;display:flex;align-items:center;gap:8px}
  .card{background:#fff;border-radius:12px;padding:14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08)}
  label{display:block;font-size:12px;font-weight:600;color:#4b5563;margin:8px 0 4px}
  input,textarea,select{width:100%;padding:10px;border:1px solid #d1d5db;border-radius:8px;font-size:14px;background:#fff}
  input:focus,select:focus,textarea:focus{outline:none;border-color:#2563eb}
  .row{display:flex;gap:8px}.row>*{flex:1}
  
  /* Item Order Styling */
  .item-box{background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:10px;margin-bottom:8px;position:relative}
  .item-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
  .item-code{font-size:11px;color:#6b7280;font-family:monospace}
  .item-row{display:grid;grid-template-columns:1fr 1fr 1.2fr;gap:6px}
  .btn-del{background:#fee2e2;color:#b91c1c;border:0;border-radius:6px;padding:4px 8px;font-size:12px;cursor:pointer;font-weight:bold}
  
  button{border:0;border-radius:8px;padding:11px 14px;font-size:14px;cursor:pointer;font-weight:600;transition:.15s}
  button:active{transform:scale(0.98)}
  .primary{background:#16a34a;color:#fff}
  .tg{background:#0284c7;color:#fff}
  .ghost{background:#e5e7eb;color:#374151}
  .btn-db{background:#4f46e5;color:#fff;width:100%}
  .actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
  .actions button{flex:1;min-width:130px}
  
  pre{white-space:pre-wrap;background:#f8fafc;border:1px dashed #cbd5e1;padding:12px;border-radius:8px;font-size:12.5px;color:#0f172a;line-height:1.45}
  .total-box{background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;padding:10px 14px;margin-top:10px;display:flex;justify-content:space-between;align-items:center}
  .total-num{font-size:18px;font-weight:700;color:#1e40af}
  .badge{display:inline-block;padding:2px 6px;border-radius:4px;font-size:11px;font-weight:600;background:#e0e7ff;color:#3730a3}
  .hist{display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #f1f5f9;padding:8px 0;font-size:13px}
  .hist button{padding:5px 8px;font-size:11px}
  details summary{cursor:pointer;font-weight:600;color:#374151}
  
  /* Autocomplete suggestion dropdown */
  .suggest-wrap{position:relative}
  .suggestions{position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid #d1d5db;border-radius:8px;max-height:180px;overflow-y:auto;z-index:99;box-shadow:0 4px 6px rgba(0,0,0,0.1)}
  .suggest-item{padding:8px 10px;font-size:13px;cursor:pointer;border-bottom:1px solid #f3f4f6}
  .suggest-item:hover{background:#f0fdf4}
  .suggest-item b{color:#15803d}
</style>
</head>
<body>
<div class="wrap">
  <h1>📦 Sales Order TR Channel</h1>

  <!-- Database & Excel Management -->
  <div class="card" style="border-left:4px solid #4f46e5">
    <div style="display:flex;justify-content:space-between;align-items:center">
      <b>📚 Database Produk & Harga</b>
      <span class="badge" id="dbCount">Memuat...</span>
    </div>
    <div style="font-size:12px;color:#6b7280;margin:6px 0 10px">
      Menggunakan database <i>PRICE WATERFALL TR CHANNEL (Signify/Philips/Ecolink)</i>.
    </div>
    <input type="file" id="excelFile" accept=".xlsx, .xls" style="display:none" onchange="handleExcelUpload(event)">
    <button class="btn-db" onclick="$('excelFile').click()">📂 Upload / Perbarui dari Excel (.xlsx)</button>
    <div id="uploadStatus" style="font-size:12px;color:#16a34a;margin-top:6px;display:none"></div>
  </div>

  <!-- Header Order -->
  <div class="card">
    <div class="row">
      <div><label>Nama Sales</label><input id="sales" placeholder="Nama sales"></div>
      <div><label>Tanggal</label><input id="tgl" type="date"></div>
    </div>
    <div class="row">
      <div>
        <label>Tipe Harga (Tier)</label>
        <select id="tier" onchange="gantiTier()">
          <option value="toko" selected>Harga ke Toko (TR)</option>
          <option value="retail">Harga Retail</option>
          <option value="rekom">Rekom. Distributor</option>
        </select>
      </div>
      <div>
        <label>Pembayaran</label>
        <select id="bayar">
          <option>Cash / Tunai</option>
          <option>Transfer</option>
          <option>Tempo 14 Hari</option>
          <option>Tempo 30 Hari</option>
        </select>
      </div>
    </div>
    <label>Nama Pelanggan / Toko</label>
    <input id="pelanggan" placeholder="Contoh: Toko Terang Jaya">
    <div class="row">
      <div><label>No. Telepon / WA</label><input id="hp" inputmode="tel" placeholder="08xxxxxxxxxx"></div>
      <div><label>Alamat / Kota</label><input id="alamat" placeholder="Jl. Raya No. 10"></div>
    </div>
  </div>

  <!-- Detail Order -->
  <div class="card">
    <div style="display:flex;justify-content:space-between;align-items:center">
      <b>Item Pesanan</b>
      <button class="ghost" style="padding:6px 12px;font-size:13px" onclick="addItem()">+ Tambah Baris</button>
    </div>

    <div id="items" style="margin-top:10px"></div>

    <div class="row" style="margin-top:10px">
      <div><label>Diskon Tambahan (Rp)</label><input id="diskon" type="number" inputmode="numeric" value="0"></div>
      <div><label>Catatan Order</label><input id="catatan" placeholder="Catatan pengiriman/spesial"></div>
    </div>

    <div class="total-box">
      <span>Total Tagihan:</span>
      <span class="total-num" id="totalDisplay">Rp 0</span>
    </div>
  </div>

  <!-- Preview & Kirim -->
  <div class="card">
    <b>Format Pesan Group Order</b>
    <pre id="preview"></pre>
    <div class="actions">
      <button class="primary" onclick="kirimWA()">📲 Kirim ke WhatsApp</button>
      <button class="tg" onclick="kirimTG()">✈️ Kirim ke Telegram</button>
      <button class="ghost" onclick="salin()">📋 Salin Pesan</button>
    </div>
    <div class="actions">
      <button class="ghost" style="background:#e0e7ff;color:#3730a3" onclick="simpan(true)">➕ Simpan & Buat Order Baru</button>
    </div>
  </div>

  <!-- Konfigurasi Telegram -->
  <div class="card">
    <details>
      <summary>⚙️ Pengaturan Bot Telegram Group</summary>
      <div style="font-size:12px;color:#6b7280;margin:6px 0">Masukkan Token bot dan Chat ID group sales untuk auto-kirim tanpa membuka aplikasi lain.</div>
      <label>Bot Token</label><input id="tgToken" placeholder="123456:ABC-DEF...">
      <label>Chat ID Group</label><input id="tgChat" placeholder="-1001234567890">
      <div class="actions"><button class="ghost" onclick="simpanSetting()">Simpan Pengaturan</button></div>
    </details>
  </div>

  <!-- Riwayat -->
  <div class="card">
    <b>Riwayat Order Terakhir</b>
    <div id="riwayat" style="margin-top:6px"></div>
  </div>
</div>

<script>
const $ = id => document.getElementById(id);
const rp = n => 'Rp ' + (Number(n)||0).toLocaleString('id-ID');

let productsDB = [];
let items = [];
let noOrder = '';

// Database bawaan Philips & Ecolink TR Channel
const defaultProducts = [
  { c: "929002299107", n: "Philips LEDBulb 8W E27 6500K 230V 1CT/12 APR", toko: 25000, retail: 31000, rekom: 22000, moq: 12 },
  { c: "929002299207", n: "Philips LEDBulb 10W E27 6500K 230V 1CT/12 APR", toko: 32000, retail: 39000, rekom: 28000, moq: 12 },
  { c: "929002299307", n: "Philips LEDBulb 12W E27 6500K 230V 1CT/12 APR", toko: 39500, retail: 49000, rekom: 35000, moq: 12 },
  { c: "929002299407", n: "Philips LEDBulb 14.5W E27 6500K 230V 1CT/12 APR", toko: 49000, retail: 61000, rekom: 43000, moq: 12 },
  { c: "929002412808", n: "Philips TrueForce Core LED HPL 30W E27 6500K", toko: 82000, retail: 102000, rekom: 72000, moq: 6 },
  { c: "929002412908", n: "Philips TrueForce Core LED HPL 40W E27 6500K", toko: 110000, retail: 138000, rekom: 97000, moq: 6 },
  { c: "929002413008", n: "Philips TrueForce Core LED HPL 50W E27 6500K", toko: 135000, retail: 169000, rekom: 119000, moq: 6 },
  { c: "929001380802", n: "Philips LEDTube Ecofit 16W 1200mm 765 T8 G13", toko: 36000, retail: 45000, rekom: 31500, moq: 20 },
  { c: "929001380702", n: "Philips LEDTube Ecofit 8W 600mm 765 T8 G13", toko: 28500, retail: 35000, rekom: 25000, moq: 20 },
  { c: "911401824182", n: "Philips SmartBright Downlight DN027B G3 10W D125 6500K", toko: 68000, retail: 85000, rekom: 60000, moq: 10 },
  { c: "911401824282", n: "Philips SmartBright Downlight DN027B G3 15W D150 6500K", toko: 88000, retail: 110000, rekom: 78000, moq: 10 },
  { c: "911401844482", n: "Philips Floodlight Essential SmartBright FL007 30W CW", toko: 145000, retail: 180000, rekom: 128000, moq: 8 },
  { c: "911401844582", n: "Philips Floodlight Essential SmartBright FL007 50W CW", toko: 195000, retail: 245000, rekom: 172000, moq: 8 },
  { c: "929002010110", n: "Ecolink LEDBulb 7W E27 6500K 220V", toko: 14000, retail: 17500, rekom: 12200, moq: 12 },
  { c: "929002010210", n: "Ecolink LEDBulb 9W E27 6500K 220V", toko: 17000, retail: 21500, rekom: 14900, moq: 12 },
  { c: "929002010310", n: "Ecolink LEDBulb 11W E27 6500K 220V", toko: 21000, retail: 26500, rekom: 18500, moq: 12 },
  { c: "911401800180", n: "Ecolink Downlight LED 9W 6500K D105", toko: 34000, retail: 42500, rekom: 30000, moq: 10 },
  { c: "911401800280", n: "Ecolink Downlight LED 12W 6500K D125", toko: 42000, retail: 52500, rekom: 37000, moq: 10 }
];

/* IndexedDB Management */
const DB_NAME = 'SalesAppDB';
const DB_VERSION = 1;

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('products')) {
        db.createObjectStore('products', { keyPath: 'id', autoIncrement: true });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function loadProducts() {
  try {
    const db = await openDB();
    const tx = db.transaction('products', 'readonly');
    const store = tx.objectStore('products');
    const all = await new Promise(res => {
      const r = store.getAll();
      r.onsuccess = () => res(r.result);
    });
    if (all && all.length > 0) {
      productsDB = all;
    } else {
      productsDB = defaultProducts;
      await saveProductsToDB(defaultProducts);
    }
  } catch (e) {
    productsDB = defaultProducts;
  }
  updateDBBadge();
}

async function saveProductsToDB(list) {
  const db = await openDB();
  const tx = db.transaction('products', 'readwrite');
  const store = tx.objectStore('products');
  store.clear();
  list.forEach(item => store.add(item));
  return new Promise(res => { tx.oncomplete = () => res(); });
}

function updateDBBadge() {
  $('dbCount').textContent = `${productsDB.length.toLocaleString('id-ID')} Produk`;
}

/* Upload & Parse Excel */
function handleExcelUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const st = $('uploadStatus');
  st.style.display = 'block';
  st.textContent = '⏳ Membaca file Excel...';

  const reader = new FileReader();
  reader.onload = async (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const parsed = [];

      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        if (rows.length < 2) return;

        // Cari baris header
        let hRowIdx = -1;
        let cIdx = {}, nIdx = -1, tokoIdx = -1, retIdx = -1, rekomIdx = -1, moqIdx = -1;

        for (let r = 0; r < Math.min(rows.length, 15); r++) {
          const rowStr = rows[r].map(c => String(c).toLowerCase().trim());
          const hasDesc = rowStr.some(c => c.includes('desc') || c.includes('nama') || c.includes('uraian'));
          const has12NC = rowStr.some(c => c.includes('12nc') || c.includes('material') || c.includes('kode') || c.includes('code'));
          if (hasDesc || has12NC) {
            hRowIdx = r;
            rowStr.forEach((val, colIdx) => {
              if (val.includes('12nc') || val.includes('kode') || val.includes('code') || val.includes('material')) cIdx = colIdx;
              if (val.includes('desc') || val.includes('nama') || val.includes('uraian') || val.includes('product')) nIdx = colIdx;
              if (val.includes('toko') || val.includes('store')) tokoIdx = colIdx;
              if (val.includes('retail') || val.includes('het') || val.includes('rrp')) retIdx = colIdx;
              if (val.includes('rekom') || val.includes('dist')) rekomIdx = colIdx;
              if (val.includes('moq') || val.includes('box') || val.includes('isi') || val.includes('qty')) moqIdx = colIdx;
            });
            break;
          }
        }

        if (hRowIdx !== -1 && (nIdx !== -1 || typeof cIdx === 'number')) {
          for (let r = hRowIdx + 1; r < rows.length; r++) {
            const row = rows[r];
            const name = String(row[nIdx] || '').trim();
            const code = String(row[cIdx] || '').trim();
            if (!name && !code) continue;

            const parseNum = v => {
              if (typeof v === 'number') return v;
              const s = String(v).replace(/[^0-9]/g, '');
              return Number(s) || 0;
            };

            const pToko = tokoIdx !== -1 ? parseNum(row[tokoIdx]) : 0;
            const pRet = retIdx !== -1 ? parseNum(row[retIdx]) : 0;
            const pRekom = rekomIdx !== -1 ? parseNum(row[rekomIdx]) : 0;
            const moq = moqIdx !== -1 ? parseNum(row[moqIdx]) : 1;

            parsed.push({
              c: code,
              n: name || code,
              toko: pToko || pRet || pRekom,
              retail: pRet || pToko,
              rekom: pRekom || pToko,
              moq: moq || 1
            });
          }
        }
      });

      if (parsed.length > 0) {
        productsDB = parsed;
        await saveProductsToDB(parsed);
        updateDBBadge();
        st.textContent = `✅ Berhasil memuat ${parsed.length.toLocaleString('id-ID')} produk dari Excel!`;
      } else {
        st.textContent = '⚠️ Tidak dapat menemukan kolom produk/harga. Format sheet dicek ulang.';
      }
    } catch (err) {
      st.textContent = '❌ Gagal memproses file: ' + err.message;
    }
  };
  reader.readAsArrayBuffer(file);
}

/* Generator & Rendering */
function genNo(){
  const d = new Date();
  const ymd = d.toISOString().slice(0,10).replace(/-/g,'');
  let c = JSON.parse(localStorage.getItem('so_counter')||'{}');
  const n = (c[ymd]||0) + 1;
  return {no:`ORD-TR-${ymd}-${String(n).padStart(3,'0')}`, ymd, n};
}
function peekNo(){ return genNo().no; }

function getActivePrice(prod) {
  const tier = $('tier').value;
  if (tier === 'retail') return prod.retail || prod.toko || 0;
  if (tier === 'rekom') return prod.rekom || prod.toko || 0;
  return prod.toko || 0;
}

function gantiTier() {
  items.forEach(it => {
    if (it.ref) it.harga = getActivePrice(it.ref);
  });
  renderItems();
}

function addItem(nama='', code='', qty=1, harga=0, ref=null){
  items.push({nama, code, qty, harga, ref});
  renderItems();
}

function delItem(i){
  items.splice(i,1);
  if(!items.length) addItem();
  else renderItems();
}

function renderItems(){
  $('items').innerHTML = items.map((it,i)=>`
    <div class="item-box">
      <div class="item-header">
        <span class="item-code">${it.code ? '12NC: ' + it.code : 'Item #' + (i+1)}</span>
        <button class="btn-del" onclick="delItem(${i})">Hapus</button>
      </div>
      <div class="suggest-wrap">
        <input placeholder="Cari nama produk / kode 12NC..." 
               value="${(it.nama||'').replace(/"/g,'&quot;')}" 
               oninput="onSearchInput(${i}, this.value)"
               autocomplete="off">
        <div id="sug-${i}" class="suggestions" style="display:none"></div>
      </div>
      <div class="item-row" style="margin-top:6px">
        <div>
          <label style="margin:2px 0">Qty (Pcs)</label>
          <input type="number" inputmode="numeric" min="1" value="${it.qty}" oninput="upd(${i},'qty',this.value)">
        </div>
        <div>
          <label style="margin:2px 0">Harga Satuan</label>
          <input type="number" inputmode="numeric" value="${it.harga}" oninput="upd(${i},'harga',this.value)">
        </div>
        <div>
          <label style="margin:2px 0">Subtotal</label>
          <input value="${rp((Number(it.qty)||0)*(Number(it.harga)||0))}" readonly style="background:#f1f5f9;font-weight:600">
        </div>
      </div>
    </div>`).join('');
  refresh();
}

function upd(i,k,v){
  items[i][k]=v;
  refresh();
}

function onSearchInput(idx, query) {
  items[idx].nama = query;
  refresh();

  const box = $(`sug-${idx}`);
  if (!query || query.length < 2) {
    box.style.display = 'none';
    return;
  }

  const q = query.toLowerCase();
  const matches = productsDB.filter(p => 
    p.n.toLowerCase().includes(q) || (p.c && p.c.toLowerCase().includes(q))
  ).slice(0, 10);

  if (!matches.length) {
    box.style.display = 'none';
    return;
  }

  box.innerHTML = matches.map(m => `
    <div class="suggest-item" onclick="pilihProduk(${idx}, '${m.c.replace(/'/g,"\\'")}')">
      <div><b>${m.n}</b></div>
      <div style="font-size:11px;color:#6b7280">12NC: ${m.c || '-'} | Harga: ${rp(getActivePrice(m))} ${m.moq ? '| Isi Box: ' + m.moq : ''}</div>
    </div>
  `).join('');
  box.style.display = 'block';
}

function pilihProduk(idx, code) {
  const p = productsDB.find(x => x.c === code);
  if (!p) return;
  items[idx].nama = p.n;
  items[idx].code = p.c;
  items[idx].harga = getActivePrice(p);
  items[idx].ref = p;
  $(`sug-${idx}`).style.display = 'none';
  renderItems();
}

function hitung(){
  const sub = items.reduce((s,i)=>s+(Number(i.qty)||0)*(Number(i.harga)||0),0);
  const disk = Number($('diskon').value)||0;
  return {sub, disk, total: Math.max(sub-disk,0)};
}

function getTierLabel(){
  const t = $('tier').value;
  if(t === 'retail') return 'Harga Retail';
  if(t === 'rekom') return 'Rekom. Distributor';
  return 'Harga ke Toko (TR)';
}

function buildMsg(){
  const h = hitung();
  const valid = items.filter(i=>i.nama.trim());
  const tgl = $('tgl').value ? new Date($('tgl').value).toLocaleDateString('id-ID',{day:'2-digit',month:'long',year:'numeric'}) : '-';

  let t = `*PESANAN ORDER SALES - TR CHANNEL*\n`;
  t += `No: *${noOrder}*\n`;
  t += `Tanggal: ${tgl}\n`;
  t += `Sales: ${$('sales').value||'-'}\n`;
  t += `Tier: ${getTierLabel()}\n\n`;

  t += `*INFORMASI PELANGGAN / TOKO*\n`;
  t += `Toko: *${$('pelanggan').value||'-'}*\n`;
  t += `No. HP: ${$('hp').value||'-'}\n`;
  t += `Alamat: ${$('alamat').value||'-'}\n`;
  t += `Pembayaran: ${$('bayar').value}\n\n`;

  t += `*RINCIAN ITEM*\n`;
  if (valid.length) {
    valid.forEach((i, idx) => {
      const sub = (Number(i.qty)||0)*(Number(i.harga)||0);
      t += `${idx+1}. ${i.code ? '['+i.code+'] ' : ''}${i.nama}\n`;
      t += `   ${i.qty} pcs x ${rp(i.harga)} = *${rp(sub)}*\n`;
    });
  } else {
    t += `- Belum ada item -\n`;
  }

  t += `\n-----------------------------\n`;
  t += `Subtotal: ${rp(h.sub)}\n`;
  if(h.disk > 0) t += `Diskon: -${rp(h.disk)}\n`;
  t += `*TOTAL ORDER: ${rp(h.total)}*\n`;

  if($('catatan').value.trim()) {
    t += `\nCatatan: ${$('catatan').value.trim()}\n`;
  }
  return t;
}

function refresh(){
  $('totalDisplay').textContent = rp(hitung().total);
  $('preview').textContent = buildMsg();
}

function validasi(){
  if(!$('sales').value.trim()) return alert('Silakan isi Nama Sales'), false;
  if(!$('pelanggan').value.trim()) return alert('Silakan isi Nama Pelanggan / Toko'), false;
  if(!items.some(i=>i.nama.trim() && Number(i.qty)>0)) return alert('Isi minimal 1 item produk'), false;
  return true;
}

function kirimWA(){
  if(!validasi()) return;
  simpan(false);
  window.open('https://wa.me/?text=' + encodeURIComponent(buildMsg()), '_blank');
}

async function kirimTG(){
  if(!validasi()) return;
  const token = localStorage.getItem('tg_token'), chat = localStorage.getItem('tg_chat');
  if(!token || !chat) return alert('Silakan atur Bot Token & Chat ID di menu Pengaturan Telegram di bawah');
  try{
    const r = await fetch(`https://api.telegram.org/bot${token}/sendMessage`,{
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({chat_id:chat, text: buildMsg().replace(/\*/g,'')})
    });
    const j = await r.json();
    if(j.ok){ simpan(false); alert('✅ Order berhasil dikirim ke group Telegram!'); }
    else alert('Gagal kirim Telegram: ' + j.description);
  }catch(e){ alert('Koneksi Telegram error: ' + e.message); }
}

function salin(){
  navigator.clipboard.writeText(buildMsg()).then(()=>alert('Teks pesanan berhasil disalin!'));
}

function simpan(reset){
  if(reset && !validasi()) return;
  let hist = JSON.parse(localStorage.getItem('so_hist')||'[]');
  if(!hist.find(h=>h.no===noOrder)){
    const g = genNo();
    let c = JSON.parse(localStorage.getItem('so_counter')||'{}');
    c[g.ymd] = g.n; localStorage.setItem('so_counter', JSON.stringify(c));
    hist.unshift({no:noOrder, pelanggan:$('pelanggan').value, total:hitung().total, msg:buildMsg()});
    localStorage.setItem('so_hist', JSON.stringify(hist.slice(0,30)));
  }
  renderRiwayat();
  if(reset) orderBaru();
}

function orderBaru(){
  ['pelanggan','hp','alamat','catatan'].forEach(k=>$(k).value='');
  $('diskon').value=0; items=[]; addItem();
  noOrder = peekNo(); refresh();
}

function renderRiwayat(){
  const hist = JSON.parse(localStorage.getItem('so_hist')||'[]');
  $('riwayat').innerHTML = hist.length ? hist.map((h,i)=>`
    <div class="hist"><div><b>${h.no}</b><br>${h.pelanggan} — ${rp(h.total)}</div>
    <div><button class="ghost" onclick="salinRiwayat(${i})">Salin</button>
    <button class="primary" onclick="waRiwayat(${i})">WA</button></div></div>`).join('')
    : '<small style="color:#6b7280">Belum ada riwayat pesanan.</small>';
}
function getH(i){ return JSON.parse(localStorage.getItem('so_hist'))[i]; }
function salinRiwayat(i){ navigator.clipboard.writeText(getH(i).msg).then(()=>alert('Pesan disalin')); }
function waRiwayat(i){ window.open('https://wa.me/?text='+encodeURIComponent(getH(i).msg),'_blank'); }

function simpanSetting(){
  localStorage.setItem('tg_token',$('tgToken').value.trim());
  localStorage.setItem('tg_chat',$('tgChat').value.trim());
  alert('Pengaturan Telegram berhasil disimpan.');
}

// Inisialisasi Aplikasi
window.onload = async () => {
  $('tgToken').value = localStorage.getItem('tg_token')||'';
  $('tgChat').value = localStorage.getItem('tg_chat')||'';
  $('tgl').value = new Date().toISOString().slice(0,10);
  $('sales').value = localStorage.getItem('so_sales')||'';

  ['sales','tgl','pelanggan','hp','bayar','alamat','diskon','catatan'].forEach(k=>$(k).addEventListener('input',()=>{
    if(k==='sales') localStorage.setItem('so_sales',$('sales').value);
    refresh();
  }));

  noOrder = peekNo();
  await loadProducts();
  addItem();
  renderRiwayat();
};
</script>
</body>
</html>
