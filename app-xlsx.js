/* =====================================================================
   COMEDOR PETROCENO · Generador XLSX nativo
   ---------------------------------------------------------------------
   Sustituye a SheetJS (CDN). Escribe un .xlsx real: ZIP(STORE)+XML a mano.
   Ventajas: 0 KB de CDN, funciona SIN INTERNET, y RESPETA el formato
   original (Calibri, merges, anchos, alturas y FORMULAS SUM).
   ===================================================================== */
(function (global) {
  'use strict';

  /* ---------------- CRC32 ---------------- */
  let _T = null;
  function crc32(b) {
    if (!_T) { _T = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); _T.push(c >>> 0); } }
    let c = 0xFFFFFFFF;
    for (let i = 0; i < b.length; i++) c = (_T[(c ^ b[i]) & 0xFF] ^ (c >>> 8)) >>> 0;
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /* ---------------- ZIP (STORE, sin compresion) ---------------- */
  function zip(files) {
    const enc = new TextEncoder(), parts = [], centrals = [];
    let offset = 0;
    const d = new Date();
    const t = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    const dt = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();

    for (const f of files) {
      const nb = enc.encode(f.n), crc = crc32(f.data);
      const lh = new DataView(new ArrayBuffer(30 + nb.length));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
      lh.setUint16(8, 0, true); lh.setUint16(10, t, true); lh.setUint16(12, dt, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, f.data.length, true); lh.setUint32(22, f.data.length, true);
      lh.setUint16(26, nb.length, true); lh.setUint16(28, 0, true);
      nb.forEach((b, i) => lh.setUint8(30 + i, b));
      const lhA = new Uint8Array(lh.buffer); parts.push(lhA, f.data);

      const ch = new DataView(new ArrayBuffer(46 + nb.length));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true); ch.setUint16(12, t, true); ch.setUint16(14, dt, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, f.data.length, true); ch.setUint32(24, f.data.length, true);
      ch.setUint16(28, nb.length, true); ch.setUint32(42, offset, true);
      nb.forEach((b, i) => ch.setUint8(46 + i, b));
      centrals.push(new Uint8Array(ch.buffer));
      offset += lhA.length + f.data.length;
    }
    const cdLen = centrals.reduce((a, c) => a + c.length, 0);
    const eo = new DataView(new ArrayBuffer(22));
    eo.setUint32(0, 0x06054b50, true); eo.setUint16(8, files.length, true); eo.setUint16(10, files.length, true);
    eo.setUint32(12, cdLen, true); eo.setUint32(16, offset, true);
    const out = new Uint8Array(offset + cdLen + 22);
    let p = 0;
    for (const x of parts) { out.set(x, p); p += x.length; }
    for (const c of centrals) { out.set(c, p); p += c.length; }
    out.set(new Uint8Array(eo.buffer), p);
    return out;
  }

  /* ---------------- Helpers ---------------- */
  const esc = s => String(s == null ? '' : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');

  function col(n) { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  /* ---------------- Estilos (replican el Excel original) ----------------
     0 Calibri 11 · 1 Calibri 11 bold · 2 Calibri 14 bold centro (titulo)
     3 Calibri 14 bold (subtitulo) · 4 Calibri 11 bold centro (encabezado)
     5 Calibri 11 centro (dato)                                       */
  const STYLES =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="4">' +
      '<font><sz val="11"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="14"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="11"/><name val="Calibri"/></font>' +
      '<font><sz val="14"/><name val="Calibri"/></font>' +
    '</fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf borderId="0" fillId="0" fontId="0" numFmtId="0"/></cellStyleXfs>' +
    '<cellXfs count="6">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>' +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>';

  /* ---------------- Constructor de hoja ----------------
     m = { name, cols:[{min,max,width}], merges:[], filas:[{r,ht,cells:[{c,s,v|f}]}] } */
  function sheetXml(m) {
    const cols = (m.cols && m.cols.length)
      ? '<cols>' + m.cols.map(c => '<col min="' + c.min + '" max="' + c.max + '" width="' + c.width + '" customWidth="1"/>').join('') + '</cols>' : '';

    const rows = m.filas.map(f => {
      const ht = f.ht ? ' ht="' + f.ht + '" customHeight="1"' : '';
      const cells = (f.cells || []).map(c => {
        const ref = col(c.c) + f.r, st = c.s ? ' s="' + c.s + '"' : '';
        if (c.f !== undefined) {
          // OBLIGATORIO: si la formula devuelve texto, la celda debe llevar
          // t="str". Sin esto Excel declara el archivo CORRUPTO.
          const esTexto = typeof c.v === 'string';
          const tAttr = esTexto ? ' t="str"' : '';
          return '<c r="' + ref + '"' + st + tAttr + '><f>' + esc(c.f) + '</f><v>' +
                 (c.v != null ? esc(c.v) : '') + '</v></c>';
        }
        if (c.v === null || c.v === undefined || c.v === '') return '<c r="' + ref + '"' + st + '/>';
        if (typeof c.v === 'number' && isFinite(c.v)) return '<c r="' + ref + '"' + st + '><v>' + c.v + '</v></c>';
        return '<c r="' + ref + '"' + st + ' t="inlineStr"><is><t xml:space="preserve">' + esc(c.v) + '</t></is></c>';
      }).join('');
      return '<row r="' + f.r + '"' + ht + '>' + cells + '</row>';
    }).join('');

    const mg = (m.merges && m.merges.length)
      ? '<mergeCells count="' + m.merges.length + '">' + m.merges.map(x => '<mergeCell ref="' + x + '"/>').join('') + '</mergeCells>' : '';

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetPr><pageSetUpPr/></sheetPr><sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
      '<sheetFormatPr defaultColWidth="14.43" defaultRowHeight="15"/>' + cols +
      '<sheetData>' + rows + '</sheetData>' + mg +
      '<printOptions/><pageMargins bottom="0.75" footer="0.0" header="0.0" left="0.7" right="0.7" top="0.75"/>' +
      '<pageSetup paperSize="9" orientation="portrait"/></worksheet>';
  }

  function workbook(sheetModels) {
    const e = s => new TextEncoder().encode(s);
    const files = [];
    let types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
    let rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    let wbx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>';

    sheetModels.forEach((m, i) => {
      const n = 'sheet' + (i + 1), rid = 'rId' + (i + 1);
      types += '<Override PartName="/xl/worksheets/' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      rels += '<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' + n + '.xml"/>';
      // Nombre de hoja valido: Excel REQUIERE 1..31 chars y prohibe : \ / ? * [ ]
      let nm = String(m.name == null ? '' : m.name).replace(/[\\\/\?\*\[\]:]/g, '').substring(0, 31).trim();
      if (!nm) nm = 'Hoja' + (i + 1);
      wbx += '<sheet name="' + esc(nm) + '" sheetId="' + (i + 1) + '" r:id="' + rid + '"/>';
      files.push({ n: 'xl/worksheets/' + n + '.xml', data: e(sheetXml(m)) });
    });

    types += '</Types>';
    rels += '<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    wbx += '</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>';

    files.unshift(
      { n: '[Content_Types].xml', data: e(types) },
      { n: '_rels/.rels', data: e('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>') },
      { n: 'xl/workbook.xml', data: e(wbx) },
      { n: 'xl/_rels/workbook.xml.rels', data: e(rels) },
      { n: 'xl/styles.xml', data: e(STYLES) }
    );
    return zip(files);
  }

  function descargar(bytes, nombre) {
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = document.createElement('a');
    a.href = url; a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  /* =====================================================================
     CONSTRUCTORES DE HOJA PARA ESTE APP
     Reciben el shape de tu app: { nombre, menus:[{nombre}], datos:{fecha:[..]}, personal:{fecha:[..]} }
     ===================================================================== */
  const TITULO = 'SOLICITUD DE COMIDAS EMPACADAS';
  const SUBTITULO = 'COMEDOR ADMINISTRATIVO PETROCEDENO';

  function fechaLarga(iso) {
    const p = String(iso || '').split('-');
    if (p.length !== 3) return iso || '';
    return p[2] + '-' + parseInt(p[1], 10) + '-' + p[0];
  }
  function fechaCorta(iso) {
    const p = String(iso || '').split('-');
    if (p.length !== 3) return (iso || 'Solicitud').substring(0, 31);
    return (p[2] + '-' + p[1]).substring(0, 31);
  }

  /** Hoja principal: replica el formato del Excel original */
  function hojaSolicitud({ areas, fecha, menusUnicos }) {
    const filas = areas.map(a => {
      const d = a.datos[fecha] || [];
      return menusUnicos.map(mn => {
        const i = menuIndex(a.menus || [], mn);
        return (i >= 0 && d[i] != null) ? Number(d[i]) || 0 : 0;
      });
    });
    const tot = menusUnicos.map((_, i) => filas.reduce((s, f) => s + (f[i] || 0), 0));
    const gran = tot.reduce((a, b) => a + b, 0);

    const rIni = 4, rFin = rIni + areas.length - 1, rTot = rFin + 1, rRes = rTot + 2;
    const m = {
      name: fechaCorta(fecha),
      cols: [{ min: 1, max: 1, width: 19.14 }, { min: 2, max: 2, width: 16.29 }, { min: 3, max: 12, width: 10.71 }],
      merges: ['B1:E1'],
      filas: [],
    };

    m.filas.push({ r: 1, ht: 42.75, cells: [{ c: 1, s: 2, v: TITULO }] });
    m.filas.push({ r: 2, ht: 42.75, cells: [{ c: 1, s: 3, v: SUBTITULO }] });

    const enc = [{ c: 0, s: 1, v: 'FECHA ' + fechaLarga(fecha) }];
    menusUnicos.forEach((mn, i) => enc.push({ c: 1 + i, s: 4, v: mn }));
    enc.push({ c: menusUnicos.length + 1, s: 4, v: 'TOTAL' });
    m.filas.push({ r: 3, ht: 30, cells: enc });

    areas.forEach((a, i) => {
      const r = rIni + i;
      const cells = [{ c: 0, s: 1, v: a.nombre }];
      for (let k = 0; k < menusUnicos.length; k++) cells.push({ c: 1 + k, s: 5, v: filas[i][k] });
      cells.push({ c: menusUnicos.length + 1, s: 5, f: 'SUM(B' + r + ':' + col(menusUnicos.length) + r + ')', v: filas[i].reduce((a, b) => a + b, 0) });
      m.filas.push({ r, ht: 30, cells });
    });

    const cTot = [{ c: 0, s: 1, v: 'TOTAL' }];
    menusUnicos.forEach((mn, k) => {
      const L = col(1 + k);
      cTot.push({ c: 1 + k, s: 5, f: 'SUM(' + L + rIni + ':' + L + rFin + ')', v: tot[k] });
    });
    cTot.push({ c: menusUnicos.length + 1, s: 5, f: 'SUM(' + col(menusUnicos.length + 1) + rIni + ':' + col(menusUnicos.length + 1) + rFin + ')', v: gran });
    m.filas.push({ r: rTot, ht: 30, cells: cTot });

    const letras = menusUnicos.map((_, k) => col(1 + k));
    m.filas.push({ r: rRes, ht: 30, cells: [{ c: 0, s: 1, f: '"ALMUERZOS EMPACADOS SOLICITADOS "&SUM(B' + rTot + ':' + col(menusUnicos.length) + rTot + ')', v: 'ALMUERZOS EMPACADOS SOLICITADOS ' + gran }] });
    menusUnicos.forEach((mn, k) => {
      m.filas.push({ r: rRes + 1 + k, ht: 30, cells: [{ c: 0, f: '"' + mn + ' = "&' + letras[k] + rTot, v: mn + ' = ' + tot[k] }] });
    });

    return m;
  }

  /** Hoja por area (opcional), con personal de entrega */
  function hojaArea(area, fecha) {
    const d = area.datos[fecha] || [];
    const pers = area.personal[fecha] || [];
    const m = { name: '', cols: [{ min: 1, max: 1, width: 22 }, { min: 2, max: 2, width: 16 }], merges: [], filas: [] };
    m.filas.push({ r: 1, ht: 42.75, cells: [{ c: 0, s: 2, v: TITULO }] });
    m.filas.push({ r: 2, ht: 42.75, cells: [{ c: 0, s: 3, v: SUBTITULO }] });
    m.filas.push({ r: 3, ht: 30, cells: [{ c: 0, s: 1, v: 'FECHA ' + fechaLarga(fecha) }] });
    m.filas.push({ r: 5, ht: 30, cells: [{ c: 0, s: 4, v: 'ÁREA' }, { c: 1, s: 4, v: area.nombre }] });
    m.filas.push({ r: 7, ht: 30, cells: [{ c: 0, s: 4, v: 'MENÚ' }, { c: 1, s: 4, v: 'CANTIDAD' }] });
    let r = 8;
    area.menus.forEach((mn, i) => { m.filas.push({ r: r++, ht: 30, cells: [{ c: 0, v: menuNombre(mn) }, { c: 1, s: 5, v: Number(d[i]) || 0 }] }); });
    const rTot = r;
    m.filas.push({ r: r++, ht: 30, cells: [{ c: 0, s: 1, v: 'TOTAL' }, { c: 1, s: 5, f: 'SUM(B8:B' + (rTot - 1) + ')', v: d.reduce((a, b) => a + (Number(b) || 0), 0) }] });
    if (pers.length) {
      m.filas.push({ r: r++, ht: 30, cells: [{ c: 0, s: 1, v: 'PERSONAL DE LA ENTREGA' }] });
      m.filas.push({ r: r++, ht: 30, cells: [{ c: 0, s: 4, v: 'NOMBRE' }, { c: 1, s: 4, v: 'CARGO' }] });
      pers.forEach(p => { m.filas.push({ r: r++, ht: 30, cells: [{ c: 0, v: p.nombre || '' }, { c: 1, v: p.rol || '' }] }); });
    }
    return m;
  }

  /** Hoja de historico */
  // El snapshot guarda menus como array de TEXTO, pero areas[] usa objetos
  // {nombre}. Aceptamos ambas formas.
  function menuNombre(m) {
    return (m && typeof m === 'object') ? (m.nombre || '') : String(m == null ? '' : m);
  }
  function menuIndex(menus, nombre) {
    for (let i = 0; i < menus.length; i++) if (menuNombre(menus[i]) === nombre) return i;
    return -1;
  }

  function hojaHistorico(snapshots, menusUnicos, fmtFecha) {
    const m = { name: 'Histórico', cols: [{ min: 1, max: 12, width: 12 }], merges: [], filas: [] };
    m.filas.push({ r: 1, ht: 30, cells: [{ c: 0, s: 2, v: 'HISTÓRICO DE ENTREGAS' }] });
    m.filas.push({ r: 2, ht: 30, cells: [{ c: 0, s: 4, v: 'Fecha' }, { c: 1, s: 4, v: 'Hora' }, { c: 2, s: 4, v: 'Área' }] });
    menusUnicos.forEach((mn, i) => m.filas[1].cells.push({ c: 3 + i, s: 4, v: mn }));
    m.filas[1].cells.push({ c: 3 + menusUnicos.length, s: 4, v: 'Total Área' });
    m.filas[1].cells.push({ c: 4 + menusUnicos.length, s: 4, v: 'Personal' });

    let r = 3;
    (snapshots || []).forEach(s => {
      s.areas.forEach(a => {
        const cells = [
          { c: 0, v: fmtFecha(s.fecha) },
          { c: 1, v: new Date(s.ts).toLocaleTimeString('es-VE') },
          { c: 2, v: a.nombre },
        ];
        menusUnicos.forEach((mn, i) => {
          const j = menuIndex(a.menus || [], mn);
          cells.push({ c: 3 + i, s: 5, v: (j >= 0 && a.valores[j] != null) ? Number(a.valores[j]) || 0 : 0 });
        });
        cells.push({ c: 3 + menusUnicos.length, s: 5, v: (a.valores || []).reduce((s2, v) => s2 + (Number(v) || 0), 0) });
        cells.push({ c: 4 + menusUnicos.length, v: (a.personal || []).map(p => p.nombre).join(', ') });
        m.filas.push({ r: r++, ht: 30, cells });
      });
    });
    return m;
  }

  global.XLSXNATIVO = {
    workbook, descargar, hojaSolicitud, hojaArea, hojaHistorico,
    fechaLarga, fechaCorta, TITULO, SUBTITULO, col, esc, zip, crc32,
  };
})(window);