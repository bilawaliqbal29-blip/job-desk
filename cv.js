/* Builds Bilawal's CV as an ATS-friendly PDF (buildCvPdf) or Word document (buildCv).
   Works in Node (module.exports) and in the browser (window.buildCv), given the docx library.
   master: the base CV. tailor (optional): {headline, summary, competencies, bullets:{roleKey:[...]}, order:[roleKeys]} */
(function (root) {
  const dash = x => String(x || "").replace(/[\u2013\u2014]/g, "-");
  function buildCv(docx, master, tailor) {
    const t = tailor || {};
    const { Document, Paragraph, TextRun, AlignmentType, BorderStyle, LevelFormat, TabStopType } = docx;
    const ACCENT = "1F4E5F", MUTED = "5B6A75", FONT = "Calibri";
    const W = 11906, MARGIN = 1000, RIGHT = W - MARGIN * 2; // A4, ~0.7in margins

    const run = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 21, bold: !!o.bold, italics: !!o.italics, color: o.color });
    const p = (children, o = {}) => new Paragraph({ children, spacing: { before: o.before || 0, after: o.after == null ? 60 : o.after, line: 264 }, alignment: o.align, tabStops: o.tabs, border: o.border, keepNext: o.keepNext });
    const section = title => p([run(title.toUpperCase(), { bold: true, size: 21, color: ACCENT })], {
      before: 200, after: 80, keepNext: true,
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 2 } }
    });
    const bullet = text => new Paragraph({ children: [run(text)], numbering: { reference: "dots", level: 0 }, spacing: { after: 40, line: 264 } });
    const rightTab = [{ type: TabStopType.RIGHT, position: RIGHT }];

    const kids = [];
    kids.push(p([run(master.name, { bold: true, size: 40, color: "16212A" })], { after: 20 }));
    kids.push(p([run(t.headline || master.headline, { size: 22, color: ACCENT, bold: true })], { after: 40 }));
    kids.push(p([run(master.contact, { size: 18, color: MUTED })], { after: 120 }));

    kids.push(section("Professional Summary"));
    kids.push(p([run(t.summary || master.summary)], { after: 60 }));

    kids.push(section("Skills"));
    kids.push(p([run((t.competencies && t.competencies.length ? t.competencies : master.competencies).join(" | "))], { after: 60 }));

    kids.push(section("Work Experience"));
    const order = t.order && t.order.length ? t.order : master.roles.map(r => r.key);
    order.forEach((key, i) => {
      const r = master.roles.find(x => x.key === key); if (!r) return;
      kids.push(p([run(r.title + "  |  " + r.company, { bold: true })], { before: i ? 140 : 40, after: 10, keepNext: true }));
      kids.push(p([run(r.place + "  |  " + dash(r.dates), { italics: true, color: MUTED, size: 19 })], { after: 40, keepNext: true }));
      const bl = (t.bullets && t.bullets[key] && t.bullets[key].length) ? t.bullets[key] : r.bullets;
      bl.forEach(b => kids.push(bullet(b)));
    });

    kids.push(section("Education"));
    master.education.forEach(e => {
      kids.push(p([run(e.degree, { bold: true })], { after: 0, keepNext: true }));
      kids.push(p([run(e.school + "  |  " + dash(e.dates), { color: MUTED, size: 19 })], { after: 80 }));
    });

    kids.push(section("Certifications"));
    master.certs.forEach(c => kids.push(p([run(c)], { after: 40 })));

    return new Document({
      creator: master.name, title: master.name + " CV",
      styles: { default: { document: { run: { font: FONT, size: 21 } } } },
      numbering: { config: [{ reference: "dots", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
      sections: [{ properties: { page: { size: { width: W, height: 16838 }, margin: { top: 900, bottom: 900, left: MARGIN, right: MARGIN } } }, children: kids }]
    });
  }

  /* ATS-friendly PDF of the same CV: real text, one column, standard headings. Needs jsPDF (window.jspdf.jsPDF). */
  function buildCvPdf(jsPDF, master, tailor) {
    const t = tailor || {};
    const clean = x => dash(x).replace(/[‘’]/g, "'").replace(/[“”]/g, '"').replace(/\s+/g, " ").trim();
    const doc = new jsPDF({ unit: "pt", format: "a4" });
    doc.setProperties({ title: master.name + " - Resume", author: master.name, subject: "Resume" });
    const W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight(), M = 48, CW = W - 2 * M;
    const ACC = [31, 78, 95], INK = [30, 30, 30], MUT = [90, 90, 90];
    let y = M;
    const ensure = h => { if (y + h > H - M) { doc.addPage(); y = M; } };
    const write = (txt, o = {}) => {
      const size = o.size || 10, lh = o.lh || 1.3;
      doc.setFont("helvetica", o.style || "normal"); doc.setFontSize(size); doc.setTextColor(...(o.color || INK));
      doc.splitTextToSize(clean(txt), CW).forEach(ln => { ensure(size * lh); doc.text(ln, M, y + size); y += size * lh; });
      y += o.gap == null ? 2 : o.gap;
    };
    const section = title => { ensure(32); y += 8; write(title, { size: 11, style: "bold", color: ACC, gap: 0 }); doc.setDrawColor(...ACC); doc.setLineWidth(0.7); doc.line(M, y + 1, W - M, y + 1); y += 7; };
    const bullet = txt => {
      doc.setFont("helvetica", "normal"); doc.setFontSize(10); doc.setTextColor(...INK);
      const lines = doc.splitTextToSize(clean(txt), CW - 14);
      ensure(13 * Math.min(lines.length, 2));
      doc.text("•", M + 3, y + 10);
      lines.forEach(ln => { ensure(13); doc.text(ln, M + 14, y + 10); y += 13; });
      y += 1.5;
    };
    write(master.name, { size: 20, style: "bold", lh: 1.15 });
    write(t.headline || master.headline, { size: 11, style: "bold", color: ACC, gap: 1 });
    write(master.contact, { size: 9, color: MUT });
    section("PROFESSIONAL SUMMARY"); write(t.summary || master.summary, { gap: 0 });
    section("SKILLS"); write((t.competencies && t.competencies.length ? t.competencies : master.competencies).map(clean).join(" | "), { gap: 0 });
    section("WORK EXPERIENCE");
    const order = t.order && t.order.length ? t.order : master.roles.map(r => r.key);
    order.forEach(key => {
      const r = master.roles.find(x => x.key === key); if (!r) return;
      ensure(44);
      write(r.title + " | " + r.company, { size: 10.5, style: "bold", gap: 0 });
      write(r.place + " | " + r.dates, { size: 9.5, style: "italic", color: MUT });
      ((t.bullets && t.bullets[key] && t.bullets[key].length) ? t.bullets[key] : r.bullets).forEach(bullet);
      y += 4;
    });
    section("EDUCATION");
    master.education.forEach(e => { write(e.degree, { size: 10.5, style: "bold", gap: 0 }); write(e.school + " | " + e.dates, { size: 9.5, color: MUT, gap: 4 }); });
    section("CERTIFICATIONS");
    master.certs.forEach(c => write(c));
    return doc.output("blob");
  }
  if (typeof module !== "undefined" && module.exports) module.exports = { buildCv, buildCvPdf }; else { root.buildCv = buildCv; root.buildCvPdf = buildCvPdf; }
})(typeof window !== "undefined" ? window : this);
