/* WWSG proposal generator: builds a .docx from the sample's template parts.
   Works in the browser (window.WWSGGen) and in Node (module.exports). */
(function (root) {
  const ENT = { nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", rsquo: '’', lsquo: '‘',
    rdquo: '”', ldquo: '“', mdash: '—', ndash: '–', hellip: '…', eacute: 'é', egrave: 'è', aacute: 'á',
    oacute: 'ó', iacute: 'í', uacute: 'ú', ntilde: 'ñ', ccedil: 'ç', uuml: 'ü', ouml: 'ö', auml: 'ä',
    reg: '®', trade: '™', copy: '©', bull: '•', middot: '·', rsaquo: '›', lsaquo: '‹' };
  function decode(s) {
    return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e) => {
      if (e[0] === '#') {
        const n = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return isNaN(n) ? m : String.fromCodePoint(n);
      }
      const k = ENT[e.toLowerCase()];
      return k === undefined ? m : k;
    });
  }
  const xesc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  const BLOCK = new Set(['p', 'div', 'li', 'ul', 'ol', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'section', 'article', 'table', 'tr']);

  /* HTML (WordPress "Content") -> [[{text,b,i}], ...] paragraphs */
  function bioToParas(html) {
    if (!html) return [];
    html = String(html).replace(/\r/g, '').replace(/<!--[\s\S]*?-->/g, '')
      .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, '')
      .replace(/\[\/?[a-z_]+[^\]]*\]/gi, ''); // strip shortcodes
    const paras = [];
    let cur = [];
    let b = 0, i = 0;
    const stack = [];
    const brk = () => { if (cur.length) paras.push(cur); cur = []; };
    const re = /<\/?([a-zA-Z0-9]+)([^>]*)>|[^<]+/g;
    let m;
    while ((m = re.exec(html))) {
      if (m[1]) {
        const tag = m[1].toLowerCase();
        const closing = m[0][1] === '/';
        if (tag === 'br') { brk(); continue; }
        if (BLOCK.has(tag)) { brk(); continue; }
        let kind = null;
        if (tag === 'em' || tag === 'i') kind = 'i';
        else if (tag === 'strong' || tag === 'b') kind = 'b';
        else if (tag === 'span' && !closing && /font-weight:\s*(bold|[6-9]00)/i.test(m[2])) kind = 'b';
        else if (tag === 'span' && !closing && /font-style:\s*italic/i.test(m[2])) kind = 'i';
        if (!closing) {
          stack.push({ tag, kind });
          if (kind === 'b') b++; if (kind === 'i') i++;
        } else {
          for (let k = stack.length - 1; k >= 0; k--) {
            if (stack[k].tag === tag) {
              const kd = stack[k].kind;
              if (kd === 'b') b--; if (kd === 'i') i--;
              stack.splice(k, 1);
              break;
            }
          }
        }
      } else {
        const chunks = decode(m[0]).split(/\n[ \t ]*\n/);
        chunks.forEach((c, idx) => {
          if (idx > 0) brk();
          const t = c.replace(/[\s ]+/g, ' ');
          if (t) cur.push({ text: t, b: b > 0, i: i > 0 });
        });
      }
    }
    brk();
    // tidy: trim ends, merge same-format runs, drop empties
    return paras.map(runs => {
      const out = [];
      runs.forEach(r => {
        const last = out[out.length - 1];
        if (last && last.b === r.b && last.i === r.i) last.text += r.text; else out.push({ ...r });
      });
      if (out.length) { out[0].text = out[0].text.replace(/^\s+/, ''); out[out.length - 1].text = out[out.length - 1].text.replace(/\s+$/, ''); }
      return out.filter(r => r.text).map(r => ({ ...r, text: r.text.replace(/ {2,}/g, ' ') }));
    }).filter(p => p.length && p.some(r => r.text.trim()));
  }

  function parasToText(paras) { return paras.map(p => p.map(r => r.text).join('')).join('\n\n'); }
  function textToParas(t) {
    return String(t || '').replace(/\r/g, '').split(/\n\s*\n/).map(s => s.replace(/\s*\n\s*/g, ' ').trim()).filter(Boolean).map(s => [{ text: s, b: false, i: false }]);
  }

  const F = '<w:rFonts w:ascii="Calibri" w:eastAsia="Times New Roman" w:hAnsi="Calibri" w:cs="Calibri"/>';
  const NOLIG = '<w:kern w:val="0"/>';
  const run = (text, { b = false, i = false, color = null, u = false, sz = 22 } = {}) =>
    `<w:r><w:rPr>${F}${b ? '<w:b/><w:bCs/>' : ''}${i ? '<w:i/><w:iCs/>' : ''}${color ? `<w:color w:val="${color}"/>` : ''}${NOLIG}${sz ? `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` : ''}${u ? '<w:u w:val="single"/>' : ''}</w:rPr><w:t xml:space="preserve">${xesc(text)}</w:t></w:r>`;

  const ORANGE = 'ED7D31';

  function drawing(n, rid) {
    const E = 1828800; // 2 inches, as in the sample
    return `<w:r><w:rPr>${F}<w:b/><w:bCs/><w:noProof/><w:color w:val="${ORANGE}"/>${NOLIG}</w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251659264 + n * 2048}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>59690</wp:posOffset></wp:positionV><wp:extent cx="${E}" cy="${E}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${1000 + n}" name="Picture ${n}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${1000 + n}" name="Picture ${n}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="${E}" cy="${E}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  }

  /* One speaker block, mirroring the sample: photo + orange heading, bio, italic fee line with webpage link */
  function speakerXml(s, n, rels) {
    let x = '';
    const heading = (s.heading || '').toUpperCase();
    let img = '';
    if (s.image) {
      const rid = `rIdImg${n}`;
      rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/speaker${n}.jpeg"/>`);
      img = drawing(n, rid);
    }
    x += `<w:p><w:pPr><w:keepNext/><w:rPr>${F}<w:b/><w:bCs/><w:color w:val="${ORANGE}"/>${NOLIG}</w:rPr></w:pPr>${img}${run(heading, { b: true, color: ORANGE, sz: null })}</w:p>`;
    (s.bio || []).forEach(p => {
      x += `<w:p>${p.map(r => run(r.text, { b: r.b, i: r.i })).join('')}</w:p>`;
    });
    let fee = run(`${s.feeLine}`, { i: true });
    if (s.url) {
      const rid = `rIdLnk${n}`;
      rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xesc(s.url)}" TargetMode="External"/>`);
      fee += run(' For additional information, please visit: ', { i: true });
      fee += `<w:hyperlink r:id="${rid}" w:history="1">${run(`${s.linkName}’s Webpage.`, { b: true, color: ORANGE, u: true })}</w:hyperlink>`;
    }
    x += `<w:p>${fee}</w:p>`;
    // clear any floating photo before the next speaker, then a blank line
    x += `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:br w:type="textWrapping" w:clear="all"/></w:r></w:p>`;
    x += `<w:p/>`;
    return x;
  }

  function feeLine(s) {
    const ref = [s.honorific, s.lastName].filter(Boolean).join(' ');
    return `${ref}’s honorarium is ${s.fee || '$xx'} plus ${s.travel || 'First-Class Travel and Expenses for one (1)'}.`;
  }

  /* JSZip: the JSZip constructor. templateB64: base64 template docx. speakers: [{heading,bio,feeLine,url,linkName,imageBytes(Uint8Array)}] */
  async function buildDocx(JSZip, templateB64, speakers, meta = {}) {
    const zip = await JSZip.loadAsync(templateB64, typeof templateB64 === 'string' ? { base64: true } : {});
    const rels = [];
    let body = '';
    speakers.forEach((s, idx) => {
      const n = idx + 1;
      body += speakerXml({ ...s, image: !!s.imageBytes }, n, rels);
      if (s.imageBytes) zip.file(`word/media/speaker${n}.jpeg`, s.imageBytes);
    });
    const doc = await zip.file('word/document.xml').async('string');
    zip.file('word/document.xml', doc.replace('{{BODY}}', () => body));
    const relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
    zip.file('word/_rels/document.xml.rels', relsXml.replace('</Relationships>', () => rels.join('') + '</Relationships>'));
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    let core = await zip.file('docProps/core.xml').async('string');
    core = core.replace(/<dc:title>.*?<\/dc:title>/, `<dc:title>${xesc(meta.title || '')}</dc:title>`)
      .replace(/<dc:creator>.*?<\/dc:creator>/, `<dc:creator>${xesc(meta.author || 'WWSG')}</dc:creator>`)
      .replace(/<cp:lastModifiedBy>.*?<\/cp:lastModifiedBy>/, `<cp:lastModifiedBy>${xesc(meta.author || 'WWSG')}</cp:lastModifiedBy>`)
      .replace(/(<dcterms:created[^>]*>).*?(<\/dcterms:created>)/, `$1${now}$2`)
      .replace(/(<dcterms:modified[^>]*>).*?(<\/dcterms:modified>)/, `$1${now}$2`);
    zip.file('docProps/core.xml', core);
    return zip;
  }

  function guessHonorific(html) {
    const t = ' ' + String(html || '').replace(/<[^>]+>/g, ' ').toLowerCase() + ' ';
    const he = (t.match(/\b(he|his|him|himself)\b/g) || []).length;
    const she = (t.match(/\b(she|her|hers|herself)\b/g) || []).length;
    if (/^\s*dr\.|\bdr\.\s/.test(t.slice(0, 200))) return 'Dr.';
    if (he > she) return 'Mr.';
    if (she > he) return 'Ms.';
    return '';
  }

  const api = { bioToParas, parasToText, textToParas, buildDocx, feeLine, guessHonorific, decode };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.WWSGGen = api;
})(typeof window !== 'undefined' ? window : globalThis);
