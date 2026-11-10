/* Local XLSX reader; JSZip bundled alongside. */
(function(){'use strict';
const parse=s=>new DOMParser().parseFromString(s,'application/xml');
const elems=(n,t)=>Array.from(n.getElementsByTagName(t));
const txt=n=>n?.textContent||'';
const attr=(n,k)=>n?.getAttribute(k)||'';
const colIndex=ref=>{let n=0;for(const c of (ref.match(/^[A-Z]+/i)||[''])[0].toUpperCase())n=n*26+c.charCodeAt(0)-64;return n-1};
async function readLocal(data){
 if(!window.JSZip)throw Error('مكتبة فك ضغط Excel غير موجودة.');
 const zip=await JSZip.loadAsync(data);
 const get=async p=>{const f=zip.file(p);if(!f)throw Error('ملف Excel غير مكتمل: '+p);return parse(await f.async('string'))};
 const workbook=await get('xl/workbook.xml');
 const rels=await get('xl/_rels/workbook.xml.rels');
 const targets={};for(const r of elems(rels,'Relationship'))targets[attr(r,'Id')]=attr(r,'Target');
 let strings=[];
 if(zip.file('xl/sharedStrings.xml')){const ss=await get('xl/sharedStrings.xml');strings=elems(ss,'si').map(si=>elems(si,'t').map(txt).join(''))}
 const sheets={},names=[];
 for(const sheet of elems(workbook,'sheet')){
  const name=attr(sheet,'name'),rid=sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')||attr(sheet,'r:id');
  let target=targets[rid];if(!target)continue;
  target=target.replace(/^\/+/, '');const path=target.startsWith('xl/')?target:'xl/'+target.replace(/^\.\//,'');
  const doc=await get(path);const rows=[];
  for(const row of elems(doc,'row')){
   const ri=Math.max(0,Number(attr(row,'r')||rows.length+1)-1);const values=[];
   for(const c of Array.from(row.children).filter(x=>x.localName==='c')){
    const ci=colIndex(attr(c,'r'));if(ci<0||ci>10000)continue;
    const type=attr(c,'t'),v=Array.from(c.children).find(x=>x.localName==='v');let value='';
    if(type==='inlineStr')value=elems(c,'t').map(txt).join('');
    else if(type==='s')value=strings[Number(txt(v))]??'';
    else if(type==='b')value=txt(v)==='1';
    else if(type==='e')value='';
    else if(v)value=type==='str'?txt(v):(Number.isFinite(Number(txt(v)))?Number(txt(v)):txt(v));
    values[ci]=value;
   }rows[ri]=values;
  }names.push(name);sheets[name]={rows};
 }return {SheetNames:names,Sheets:sheets};
}
window.XLSX={readLocal,utils:{sheet_to_json:(sheet,opts)=>sheet.rows.filter(Boolean).map(r=>Array.from({length:r.length},(_,i)=>r[i]??''))}};
})();
