from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

XLSX_DOWNLOAD = r'''function downloadXls(){
  if(!PLAN.length){alert('Prima crea il planning');return;}
  const rows=exportRows();
  function xml(v){return String(v==null?'':v).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,' ').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));}
  function enc(s){return new TextEncoder().encode(s);}
  function u16(n){return [n&255,(n>>>8)&255];}
  function u32(n){return [n&255,(n>>>8)&255,(n>>>16)&255,(n>>>24)&255];}
  function colName(n){let s='';while(n>0){let m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
  function crc32buf(buf){let crc=-1;for(let i=0;i<buf.length;i++){crc^=buf[i];for(let j=0;j<8;j++)crc=(crc>>>1)^((crc&1)?0xEDB88320:0);}return (crc^(-1))>>>0;}
  function zip(files){
    const chunks=[], central=[];let offset=0;
    files.forEach(f=>{
      const name=enc(f.name), data=enc(f.data), crc=crc32buf(data), size=data.length;
      const local=new Uint8Array([0x50,0x4b,0x03,0x04,20,0,0,0,0,0,0,0,0,0,...u32(crc),...u32(size),...u32(size),...u16(name.length),0,0]);
      chunks.push(local,name,data);
      central.push({name,crc,size,offset});
      offset+=local.length+name.length+data.length;
    });
    const centralStart=offset;
    central.forEach(c=>{
      const h=new Uint8Array([0x50,0x4b,0x01,0x02,20,0,20,0,0,0,0,0,0,0,0,0,...u32(c.crc),...u32(c.size),...u32(c.size),...u16(c.name.length),0,0,0,0,0,0,0,0,0,0,...u32(c.offset)]);
      chunks.push(h,c.name);
      offset+=h.length+c.name.length;
    });
    const centralSize=offset-centralStart;
    chunks.push(new Uint8Array([0x50,0x4b,0x05,0x06,0,0,0,0,...u16(central.length),...u16(central.length),...u32(centralSize),...u32(centralStart),0,0]));
    return new Blob(chunks,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }
  const widths=rows[0].map((_,c)=>Math.min(45,Math.max(10,...rows.map(r=>String(r[c]??'').length+3))));
  const cols='<cols>'+widths.map((w,i)=>'<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>').join('')+'</cols>';
  function rowStyle(i){if(i===0)return 1;const p=PLAN[i-1];if(p&&typeof isGrabOnly==='function'&&isGrabOnly(p))return 2;if(p&&typeof isDual==='function'&&isDual(p))return 3;return 0;}
  const sheetRows=rows.map((r,ri)=>{
    const s=rowStyle(ri);
    return '<row r="'+(ri+1)+'">'+r.map((v,ci)=>{
      const ref=colName(ci+1)+(ri+1);
      return '<c r="'+ref+'" s="'+s+'" t="inlineStr"><is><t>'+xml(v)+'</t></is></c>';
    }).join('')+'</row>';
  }).join('');
  const sheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+cols+'<sheetData>'+sheetRows+'</sheetData><autoFilter ref="A1:N'+rows.length+'"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
  const workbook='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Planning" sheetId="1" r:id="rId1"/></sheets></workbook>';
  const rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
  const wbRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
  const types='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
  const styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF999999"/></left><right style="thin"><color rgb="FF999999"/></right><top style="thin"><color rgb="FF999999"/></top><bottom style="thin"><color rgb="FF999999"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  const blob=zip([{name:'[Content_Types].xml',data:types},{name:'_rels/.rels',data:rels},{name:'xl/workbook.xml',data:workbook},{name:'xl/_rels/workbook.xml.rels',data:wbRels},{name:'xl/worksheets/sheet1.xml',data:sheet},{name:'xl/styles.xml',data:styles}]);
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='planning_'+(document.getElementById('agent').value||'agente')+'_'+(document.getElementById('month').value||'mese')+'.xlsx';
  a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),3000);
}'''


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
    if not ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" in html and "crc32buf" in html):
        html = replace_between(html, "function downloadXls(){", "function parsePrev(t)", XLSX_DOWNLOAD)
        path.write_text(html, encoding="utf-8")
        print("Export XLSX reale applicato")
    else:
        print("Export XLSX reale già presente")

    try:
        from planning_transferte_patch import main as transferte_main
        transferte_main()
    except Exception as exc:
        raise RuntimeError(f"Errore patch trasferte: {exc}") from exc


if __name__ == "__main__":
    main()
