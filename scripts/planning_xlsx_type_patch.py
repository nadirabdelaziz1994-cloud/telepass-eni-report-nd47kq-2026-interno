from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH = r'''
<script>
(function(){
  function clsType(p){
    const g=!!(p&&p.is_grab), t=!!(p&&p.is_tp);
    if(g && t) return 'dual';
    if(g) return 'grab';
    return 'tp';
  }
  window.isGrabOnly = isGrabOnly = function(p){return clsType(p)==='grab';};
  window.isDual = isDual = function(p){return clsType(p)==='dual';};
  window.isGrab = isGrab = function(p){return clsType(p)==='grab' || clsType(p)==='dual';};
  window.typeLabel = typeLabel = function(p){const c=clsType(p);return c==='dual'?'GRAB & GO + TELEPASS POINT':(c==='grab'?'GRAB & GO':'TELEPASS POINT');};
  window.visitMin = visitMin = function(p){const c=clsType(p);return c==='grab'?15:(c==='dual'?60:45);};

  const _oldApplyRemote = window.applyRemote;
  window.applyRemote = applyRemote = function(){
    const rows = _oldApplyRemote ? _oldApplyRemote() : [];
    return rows.map(p=>{
      const q=Object.assign({},p);
      const c=clsType(q);
      q.is_grab = c==='grab' || c==='dual';
      q.is_tp = c==='tp' || c==='dual';
      return q;
    });
  };

  window.generatePlanning = generatePlanning = function(){
    const agent=document.getElementById('agent').value,mv=document.getElementById('month').value,inc=document.getElementById('grab').value==='yes';
    if(!agent||!mv){alert('Scegli agente e mese');return;}
    let pts=allPoints().filter(p=>p.agent_display===agent && (inc || !isGrabOnly(p)) && p.lat!=null && p.lng!=null);
    if(!pts.length){alert('Nessun PV con coordinate per questo agente');return;}
    const days=workdays(mv),start=startPoint(pts,document.getElementById('start').value);
    PLAN=assign(order(pts,start,mv),days,start);
    renderTable(days.length);
  };

  window.renderTable = renderTable = function(workdayCount){
    const days=[...new Set(PLAN.map(p=>p.date))].length;
    const grabOnly=PLAN.filter(isGrabOnly).length, dual=PLAN.filter(isDual).length;
    const rows=PLAN.map((p,i)=>{
      const c=clsType(p);
      const trClass=c==='dual'?'grab-dual':(c==='grab'?'grab':'tp');
      const pill=c==='dual'?'<span class="pill warn">Grab & Go + TPoint</span>':(c==='grab'?'<span class="pill grab">Solo Grab & Go</span>':'<span class="pill">Solo TPoint</span>');
      return '<tr class="'+trClass+'"><td><button class="btn light small" onclick="moveRow('+i+',-1)">↑</button> <button class="btn light small" onclick="moveRow('+i+',1)">↓</button> <button class="btn bad small" onclick="delRow('+i+')">x</button></td><td>'+esc(p.dateLabel)+'</td><td>'+esc(p.pdv)+'</td><td>'+pill+'</td><td class="city"><b>'+esc(p.city)+'</b><span>'+esc(p.address||'')+'</span></td><td>'+esc(p.province||'')+'</td><td>'+esc(p.region||'')+'</td><td>'+esc(p.rzv||'')+'</td><td>'+esc(p.cr||'')+'</td><td>'+esc(p.focal||'')+'</td><td class="num">'+fmt(p.travel_km||0)+' km</td><td class="num">'+fmt(p.visit_min||0)+' min</td><td>'+esc(p.day_load||'')+'</td></tr>';
    }).join('');
    document.getElementById('result').innerHTML='<style>.grab-dual{box-shadow:inset 5px 0 0 #eab308;background:#fff7d6}.grab{box-shadow:inset 5px 0 0 #2563eb;background:#eff6ff}</style><section class="metric"><div class="box"><h4>PV pianificati</h4><div class="big">'+fmt(PLAN.length)+'</div></div><div class="box"><h4>Solo Grab & Go</h4><div class="big">'+fmt(grabOnly)+'</div></div><div class="box"><h4>Grab & Go + TPoint</h4><div class="big">'+fmt(dual)+'</div></div><div class="box"><h4>Giorni usati</h4><div class="big">'+fmt(days)+'</div><div class="muted">Lavorativi: '+fmt(workdayCount||0)+'</div></div></section><section class="card"><div class="muted" style="margin-bottom:8px">Blu = solo Grab & Go · Giallo = Grab & Go + Telepass Point · Normale = solo Telepass Point.</div><div class="table-wrap"><table><thead><tr><th>Modifica</th><th>Data</th><th>PV</th><th>Tipo</th><th>Città / indirizzo</th><th>Prov.</th><th>Regione</th><th>RZV</th><th>CR</th><th>Focal Point ENI</th><th>Km</th><th>Visita</th><th>Carico giorno</th></tr></thead><tbody>'+rows+'</tbody></table></div></section>';
  };

  window.exportRows = exportRows = function(){
    return [['n° WEEK','DATA','ORA ','n° PV','AC','RZV','CR','Regione','Provincia','Città ','Indirizzo','FOCAL POINT ENI','MY WORLD','CONFERMA ENI']]
      .concat(PLAN.map(p=>[weekNum(p.date),p.dateOnly,'',p.pdv,acVal(p),p.rzv||'',p.cr||'',p.region||'',p.province||'',p.city||'',p.address||'',p.focal||'',p.agent_display||document.getElementById('agent').value,'']));
  };

  function xmlEsc(v){return String(v==null?'':v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));}
  function colName(n){let s='';while(n>0){let m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
  function crc32(str){let crc=-1;for(let i=0;i<str.length;i++){crc^=str.charCodeAt(i);for(let j=0;j<8;j++)crc=(crc>>>1)^((crc&1)?0xEDB88320:0);}return (crc^(-1))>>>0;}
  function strToU8(s){return new TextEncoder().encode(s);}
  function u16(n){return [n&255,(n>>>8)&255];}
  function u32(n){return [n&255,(n>>>8)&255,(n>>>16)&255,(n>>>24)&255];}
  function zip(files){
    const chunks=[], central=[];let offset=0;
    files.forEach(f=>{
      const name=strToU8(f.name), data=strToU8(f.data), crc=crc32(f.data), size=data.length;
      const local=new Uint8Array([0x50,0x4b,0x03,0x04,20,0,0,0,0,0,0,0,0,0,...u32(crc),...u32(size),...u32(size),...u16(name.length),0,0]);
      chunks.push(local,name,data);
      central.push({f,name,crc,size,offset});
      offset+=local.length+name.length+data.length;
    });
    const centralStart=offset;
    central.forEach(c=>{
      const h=new Uint8Array([0x50,0x4b,0x01,0x02,20,0,20,0,0,0,0,0,0,0,0,0,...u32(c.crc),...u32(c.size),...u32(c.size),...u16(c.name.length),0,0,0,0,0,0,0,0,0,0,...u32(c.offset)]);
      chunks.push(h,c.name);offset+=h.length+c.name.length;
    });
    const centralSize=offset-centralStart;
    chunks.push(new Uint8Array([0x50,0x4b,0x05,0x06,0,0,0,0,...u16(central.length),...u16(central.length),...u32(centralSize),...u32(centralStart),0,0]));
    return new Blob(chunks,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }
  function xlsxBlob(){
    const rows=exportRows();
    const widths=rows[0].map((_,c)=>Math.min(55,Math.max(10,...rows.map(r=>String(r[c]??'').length+3))));
    const cols='<cols>'+widths.map((w,i)=>'<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>').join('')+'</cols>';
    function styleFor(i){if(i===0)return 1;const p=PLAN[i-1];if(!p)return 0;if(isGrabOnly(p))return 2;if(isDual(p))return 3;return 0;}
    const sheetRows=rows.map((r,ri)=>'<row r="'+(ri+1)+'">'+r.map((v,ci)=>{const ref=colName(ci+1)+(ri+1),s=styleFor(ri);const isNum=(ci===0&&ri>0);return '<c r="'+ref+'" s="'+s+'"'+(isNum?'':' t="inlineStr"')+'>'+(isNum?'<v>'+xmlEsc(v)+'</v>':'<is><t>'+xmlEsc(v)+'</t></is>')+'</c>';}).join('')+'</row>').join('');
    const sheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+cols+'<sheetData>'+sheetRows+'</sheetData><autoFilter ref="A1:N'+rows.length+'"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
    const workbook='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Planning" sheetId="1" r:id="rId1"/></sheets></workbook>';
    const rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    const wbRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    const types='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
    const styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF999999"/></left><right style="thin"><color rgb="FF999999"/></right><top style="thin"><color rgb="FF999999"/></top><bottom style="thin"><color rgb="FF999999"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    return zip([{name:'[Content_Types].xml',data:types},{name:'_rels/.rels',data:rels},{name:'xl/workbook.xml',data:workbook},{name:'xl/_rels/workbook.xml.rels',data:wbRels},{name:'xl/worksheets/sheet1.xml',data:sheet},{name:'xl/styles.xml',data:styles}]);
  }
  window.downloadXls = downloadXls = function(){
    if(!PLAN.length){alert('Prima crea il planning');return;}
    const a=document.createElement('a');a.href=URL.createObjectURL(xlsxBlob());a.download='planning_'+(document.getElementById('agent').value||'agente')+'_'+(document.getElementById('month').value||'mese')+'.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),3000);
  };
  if(typeof renderAll==='function') renderAll();
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, patch saltata")
        return
    html = path.read_text(encoding="utf-8")
    if "function xlsxBlob()" in html:
        print("XLSX/type patch già presente")
        return
    html = html.replace("</body>", "<!-- XLSX export and dual type handling -->\n" + PATCH + "\n</body>", 1)
    path.write_text(html, encoding="utf-8")
    print("XLSX/type patch applicata")


if __name__ == "__main__":
    main()
