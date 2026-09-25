const input = document.getElementById("fileInput");
const dropzone = document.getElementById("dropzone");
const fileList = document.getElementById("fileList");
const counter = document.getElementById("counter");
const statusEl = document.getElementById("status");
const convertBtn = document.getElementById("convertBtn");
const clearBtn = document.getElementById("clearBtn");

let files = [];
const allowed = /\.(xls|xlsx|xml|csv|txt|tsv)$/i;

dropzone.addEventListener("click", () => input.click());
dropzone.addEventListener("keydown", e => {
  if (e.key === "Enter" || e.key === " ") input.click();
});
input.addEventListener("change", e => addFiles([...e.target.files]));
["dragenter","dragover"].forEach(ev => dropzone.addEventListener(ev, e => {
  e.preventDefault(); dropzone.classList.add("dragover");
}));
["dragleave","drop"].forEach(ev => dropzone.addEventListener(ev, e => {
  e.preventDefault(); dropzone.classList.remove("dragover");
}));
dropzone.addEventListener("drop", e => addFiles([...e.dataTransfer.files]));
clearBtn.addEventListener("click", () => { files=[]; input.value=""; render(); });
convertBtn.addEventListener("click", convertAll);

function addFiles(newFiles) {
  for (const f of newFiles) {
    if (allowed.test(f.name) && !files.some(x => x.name === f.name && x.size === f.size)) {
      files.push(f);
    }
  }
  render();
}
function render() {
  fileList.innerHTML = "";
  counter.textContent = files.length ? `${files.length} arquivo(s) selecionado(s)` : "Nenhum arquivo selecionado";
  convertBtn.disabled = !files.length;
  clearBtn.disabled = !files.length;
  files.forEach((f,i) => {
    const card = document.createElement("article");
    card.className = "file-card";
    card.id = `file-${i}`;
    card.innerHTML = `<div class="file-meta"><div class="file-name">${escapeHtml(f.name)}</div>
      <div class="file-detail">${formatBytes(f.size)}</div></div>
      <span class="badge">Aguardando</span>`;
    fileList.appendChild(card);
  });
}
async function convertAll() {
  convertBtn.disabled = true; clearBtn.disabled = true;
  statusEl.textContent = "Convertendo...";
  let ok=0, errors=0;

  for (let i=0; i<files.length; i++) {
    const badge = document.querySelector(`#file-${i} .badge`);
    badge.textContent = "Processando";
    try {
      const wb = await fileToWorkbook(files[i]);
      const out = XLSX.write(wb, {bookType:"xlsx", type:"array"});
      downloadBlob(out, outputName(files[i].name));
      badge.textContent = "Convertido";
      badge.className = "badge ok";
      ok++;
    } catch (err) {
      console.error(err);
      badge.textContent = "Erro";
      badge.className = "badge error";
      const detail = document.querySelector(`#file-${i} .file-detail`);
      detail.textContent += ` · ${err.message || err}`;
      errors++;
    }
  }
  statusEl.textContent = `Finalizado: ${ok} convertido(s), ${errors} erro(s).`;
  convertBtn.disabled = false; clearBtn.disabled = false;
}
async function fileToWorkbook(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  const buffer = await file.arrayBuffer();

  // XLS/XLSX: SheetJS detecta o formato pelo conteúdo.
  if (ext === "xls" || ext === "xlsx") {
    try {
      return XLSX.read(buffer, {type:"array", cellDates:true});
    } catch (_) {
      // Alguns sistemas entregam texto tabulado com extensão .xls.
      const text = decodeBuffer(buffer);
      return delimitedTextToWorkbook(text, ext);
    }
  }

  if (ext === "xml") {
    const text = decodeBuffer(buffer);
    return xmlToWorkbook(text);
  }

  const text = decodeBuffer(buffer);
  return delimitedTextToWorkbook(text, ext);
}
function decodeBuffer(buffer) {
  const bytes = new Uint8Array(buffer);

  // BOM UTF-32 LE
  if (bytes.length >= 4 && bytes[0]===0xFF && bytes[1]===0xFE && bytes[2]===0x00 && bytes[3]===0x00)
    return decodeUtf32(bytes, true, 4);

  // BOM UTF-32 BE
  if (bytes.length >= 4 && bytes[0]===0x00 && bytes[1]===0x00 && bytes[2]===0xFE && bytes[3]===0xFF)
    return decodeUtf32(bytes, false, 4);

  // Heurística UTF-32 LE sem BOM
  if (bytes.length >= 8 && bytes[1]===0 && bytes[2]===0 && bytes[3]===0)
    return decodeUtf32(bytes, true, 0);

  if (bytes.length >= 2 && bytes[0]===0xFF && bytes[1]===0xFE)
    return new TextDecoder("utf-16le").decode(bytes.subarray(2));

  if (bytes.length >= 2 && bytes[0]===0xFE && bytes[1]===0xFF) {
    const swapped = new Uint8Array(bytes.length-2);
    for(let i=2;i+1<bytes.length;i+=2){ swapped[i-2]=bytes[i+1]; swapped[i-1]=bytes[i]; }
    return new TextDecoder("utf-16le").decode(swapped);
  }

  try { return new TextDecoder("utf-8", {fatal:true}).decode(bytes).replace(/^\uFEFF/,""); }
  catch { return new TextDecoder("windows-1252").decode(bytes); }
}
function decodeUtf32(bytes, little, start) {
  let s="";
  for(let i=start;i+3<bytes.length;i+=4){
    let cp = little
      ? (bytes[i] | bytes[i+1]<<8 | bytes[i+2]<<16 | bytes[i+3]<<24) >>> 0
      : (bytes[i]<<24 | bytes[i+1]<<16 | bytes[i+2]<<8 | bytes[i+3]) >>> 0;
    if(cp===0) continue;
    if(cp<=0x10FFFF) s += String.fromCodePoint(cp);
  }
  return s.replace(/^\uFEFF/,"");
}
function detectDelimiter(text, ext) {
  if (ext === "tsv") return "\t";
  const sample = text.split(/\r?\n/).slice(0,30);
  const candidates = ["\t",";",",","|"];
  let best=null, bestScore=-1;
  for(const d of candidates){
    const counts=sample.filter(Boolean).map(line => splitDelimited(line,d).length-1);
    if(!counts.length) continue;
    const nonzero=counts.filter(n=>n>0);
    if(!nonzero.length) continue;
    const mode = nonzero.sort((a,b)=>nonzero.filter(v=>v===a).length-nonzero.filter(v=>v===b).length).pop();
    const score = nonzero.filter(n=>n===mode).length*100 + mode;
    if(score>bestScore){bestScore=score;best=d;}
  }
  if(!best) throw new Error("Delimitador não reconhecido");
  return best;
}
function splitDelimited(line, delimiter) {
  const out=[]; let cur="", quoted=false;
  for(let i=0;i<line.length;i++){
    const ch=line[i];
    if(ch === '"'){
      if(quoted && line[i+1] === '"'){ cur+='"'; i++; }
      else quoted=!quoted;
    } else if(ch===delimiter && !quoted){ out.push(cur); cur=""; }
    else cur+=ch;
  }
  out.push(cur); return out;
}
function delimitedTextToWorkbook(text, ext) {
  const delimiter=detectDelimiter(text,ext);
  const rows=text.split(/\r?\n/).filter(x=>x.trim().length).map(x=>splitDelimited(x,delimiter));
  if(!rows.length) throw new Error("Arquivo vazio");
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Dados");
  return wb;
}
function xmlToWorkbook(text) {
  const doc = new DOMParser().parseFromString(text,"application/xml");
  if(doc.querySelector("parsererror")) throw new Error("XML inválido");

  const all=[...doc.getElementsByTagName("*")];
  let best=null;
  for(const parent of all){
    const children=[...parent.children];
    const groups={};
    children.forEach(c => (groups[c.localName] ??= []).push(c));
    for(const [name,group] of Object.entries(groups)){
      if(group.length<2) continue;
      const richness=group.reduce((s,e)=>s+e.children.length+e.attributes.length,0);
      const score=group.length*1000+richness;
      if(!best || score>best.score) best={name,group,score};
    }
  }
  const records = best ? best.group : [doc.documentElement];
  const flattened=records.map(e=>flattenXml(e));
  const cols=[];
  const seen=new Set();
  flattened.forEach(r=>Object.keys(r).forEach(k=>{if(!seen.has(k)){seen.add(k);cols.push(k);}}));
  if(!cols.length) throw new Error("XML sem estrutura tabular reconhecível");
  const rows=[cols,...flattened.map(r=>cols.map(c=>r[c] ?? ""))];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Dados");
  return wb;
}
function flattenXml(el,prefix="") {
  const obj={};
  [...el.attributes].forEach(a=>obj[prefix ? `${prefix}.@${a.localName}` : `@${a.localName}`]=a.value);
  const children=[...el.children];
  if(!children.length){
    if(prefix) obj[prefix]=(el.textContent||"").trim();
    return obj;
  }
  const counts={}; children.forEach(c=>counts[c.localName]=(counts[c.localName]||0)+1);
  const idx={};
  children.forEach(c=>{
    idx[c.localName]=(idx[c.localName]||0)+1;
    let name=c.localName;
    if(counts[name]>1) name=`${name}_${idx[c.localName]}`;
    const path=prefix ? `${prefix}.${name}` : name;
    Object.assign(obj,flattenXml(c,path));
  });
  return obj;
}
function outputName(name){ return name.replace(/\.[^.]+$/, "") + ".xlsx"; }
function downloadBlob(data,name){
  const blob=new Blob([data],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a"); a.href=url; a.download=name; a.click();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
}
function formatBytes(n){
  if(n<1024)return `${n} B`;
  if(n<1024**2)return `${(n/1024).toFixed(1)} KB`;
  return `${(n/1024**2).toFixed(1)} MB`;
}
function escapeHtml(s){const d=document.createElement("div");d.textContent=s;return d.innerHTML;}
render();

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  navigator.serviceWorker.register("./service-worker.js").catch(console.error);
}
