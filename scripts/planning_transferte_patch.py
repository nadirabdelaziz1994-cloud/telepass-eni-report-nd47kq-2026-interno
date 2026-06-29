from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

DOWNLOAD_XLS = r'''function downloadXls(){
  if(!PLAN.length){alert('Prima crea il planning');return;}
  const meta=exportRowsMeta();
  const rows=meta.map(x=>x.row);
  function xml(v){return String(v==null?'':v).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,' ').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));}
  function enc(s){return new TextEncoder().encode(s);}
  function u16(n){return [n&255,(n>>>8)&255];}
  function u32(n){return [n&255,(n>>>8)&255,(n>>>16)&255,(n>>>24)&255];}
  function colName(n){let s='';while(n>0){let m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
  function crc32buf(buf){let crc=-1;for(let i=0;i<buf.length;i++){crc^=buf[i];for(let j=0;j<8;j++)crc=(crc>>>1)^((crc&1)?0xEDB88320:0);}return (crc^(-1))>>>0;}
  function zip(files){
    const chunks=[], central=[];let offset=0;
    files.forEach(f=>{const name=enc(f.name),data=enc(f.data),crc=crc32buf(data),size=data.length;const local=new Uint8Array([0x50,0x4b,0x03,0x04,20,0,0,0,0,0,0,0,0,0,...u32(crc),...u32(size),...u32(size),...u16(name.length),0,0]);chunks.push(local,name,data);central.push({name,crc,size,offset});offset+=local.length+name.length+data.length;});
    const centralStart=offset;
    central.forEach(c=>{const h=new Uint8Array([0x50,0x4b,0x01,0x02,20,0,20,0,0,0,0,0,0,0,0,0,...u32(c.crc),...u32(c.size),...u32(c.size),...u16(c.name.length),0,0,0,0,0,0,0,0,0,0,...u32(c.offset)]);chunks.push(h,c.name);offset+=h.length+c.name.length;});
    const centralSize=offset-centralStart;
    chunks.push(new Uint8Array([0x50,0x4b,0x05,0x06,0,0,0,0,...u16(central.length),...u16(central.length),...u32(centralSize),...u32(centralStart),0,0]));
    return new Blob(chunks,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }
  const widths=rows[0].map((_,c)=>Math.min(45,Math.max(10,...rows.map(r=>String(r[c]??'').length+3))));
  const cols='<cols>'+widths.map((w,i)=>'<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>').join('')+'</cols>';
  function styleFor(m){if(m.kind==='header')return 1;if(m.kind==='marker')return 4;const p=m.plan;if(p&&typeof isGrabOnly==='function'&&isGrabOnly(p))return 2;if(p&&typeof isDual==='function'&&isDual(p))return 3;return 0;}
  const merges=[];
  const sheetRows=meta.map((m,ri)=>{const rowNum=ri+1,s=styleFor(m);if(m.kind==='marker')merges.push('A'+rowNum+':N'+rowNum);return '<row r="'+rowNum+'">'+m.row.map((v,ci)=>{const ref=colName(ci+1)+rowNum;return '<c r="'+ref+'" s="'+s+'" t="inlineStr"><is><t>'+xml(v)+'</t></is></c>';}).join('')+'</row>';}).join('');
  const mergeXml=merges.length?'<mergeCells count="'+merges.length+'">'+merges.map(r=>'<mergeCell ref="'+r+'"/>').join('')+'</mergeCells>':'';
  const sheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+cols+'<sheetData>'+sheetRows+'</sheetData>'+mergeXml+'<autoFilter ref="A1:N'+rows.length+'"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
  const workbook='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Planning" sheetId="1" r:id="rId1"/></sheets></workbook>';
  const rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
  const wbRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
  const types='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
  const styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF548235"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF999999"/></left><right style="thin"><color rgb="FF999999"/></right><top style="thin"><color rgb="FF999999"/></top><bottom style="thin"><color rgb="FF999999"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"><alignment horizontal="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  const blob=zip([{name:'[Content_Types].xml',data:types},{name:'_rels/.rels',data:rels},{name:'xl/workbook.xml',data:workbook},{name:'xl/_rels/workbook.xml.rels',data:wbRels},{name:'xl/worksheets/sheet1.xml',data:sheet},{name:'xl/styles.xml',data:styles}]);
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='planning_'+(document.getElementById('agent').value||'agente')+'_'+(document.getElementById('month').value||'mese')+'.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),3000);
}'''

PATCH_JS = r'''
<script>
(function(){
  function maxTripDays(){const el=document.getElementById('tripDays');const n=parseInt(el&&el.value?el.value:'0',10);return Number.isFinite(n)&&n>0?n:0;}
  window.isTrasfertaPoint = isTrasfertaPoint = function(p){
    if(!p)return false;
    const home=startPoint(PLAN.length?PLAN:allPoints(),document.getElementById('start').value);
    if(!home||!home.lat||!home.lng)return false;
    const sameRegion=norm(home.region||'') && norm(home.region||'')===norm(p.region||'');
    const dist=km(home,p);
    return (!sameRegion && norm(p.region||'')) || dist>=180;
  };
  window.exportRowsMeta = exportRowsMeta = function(){
    const header=['n° WEEK','DATA','ORA ','n° PV','AC','RZV','CR','Regione','Provincia','Città ','Indirizzo','FOCAL POINT ENI','MY WORLD','CONFERMA ENI'];
    const meta=[{kind:'header',row:header}];
    const max=maxTripDays();
    let open=false,daysInBlock=new Set();
    function marker(txt){return {kind:'marker',row:[txt,'','','','','','','','','','','','','']};}
    PLAN.forEach((p,idx)=>{
      const tr=isTrasfertaPoint(p);
      if(tr){
        if(!open){meta.push(marker('INIZIO TRASFERTA'));open=true;daysInBlock=new Set();}
        daysInBlock.add(p.date);
        if(max && daysInBlock.size>max){meta.push(marker('FINE TRASFERTA'));meta.push(marker('INIZIO TRASFERTA'));daysInBlock=new Set([p.date]);}
      }else if(open){meta.push(marker('FINE TRASFERTA'));open=false;daysInBlock=new Set();}
      meta.push({kind:'data',plan:p,row:[weekNum(p.date),p.dateOnly,'',p.pdv,acVal(p),p.rzv||'',p.cr||'',p.region||'',p.province||'',p.city||'',p.address||'',p.focal||'',p.agent_display||document.getElementById('agent').value,'']});
      if(open && idx===PLAN.length-1){meta.push(marker('FINE TRASFERTA'));open=false;}
    });
    return meta;
  };
})();
</script>
'''


def replace_between(text, start_marker, end_marker, replacement):
    start = text.find(start_marker)
    if start == -1:
        raise RuntimeError("downloadXls start marker not found")
    end = text.find(end_marker, start)
    if end == -1:
        raise RuntimeError("downloadXls end marker not found")
    return text[:start] + replacement + text[end:]


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, patch saltata")
        return
    html = path.read_text(encoding="utf-8")
    html = html.replace('<button class="btn light" onclick="downloadCsv()">Scarica CSV</button>', '')
    target = '<div><label>Planning precedente</label><input id="prev" type="file" accept=".csv,.txt,.xls"></div>'
    if 'id="tripDays"' not in html and target in html:
        html = html.replace(target, '<div><label>Giorni max trasferta</label><input id="tripDays" type="number" min="1" max="31" value="5" placeholder="es. 5"></div>' + target, 1)
    html = replace_between(html, "function downloadXls(){", "function parsePrev(t)", DOWNLOAD_XLS)
    if "function maxTripDays()" not in html:
        html = html.replace("</body>", PATCH_JS + "\n</body>", 1)
    path.write_text(html, encoding="utf-8")
    print("Patch trasferte/excel applicata")


if __name__ == "__main__":
    main()
