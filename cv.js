/* Builds Bilawal's CV as a Word document.
   Works in Node (module.exports) and in the browser (window.buildCv), given the docx library.
   master: the base CV. tailor (optional): {headline, summary, competencies, bullets:{roleKey:[...]}, order:[roleKeys]} */
(function (root) {
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

    kids.push(section("Core Competencies"));
    kids.push(p([run((t.competencies && t.competencies.length ? t.competencies : master.competencies).join("  ·  "))], { after: 60 }));

    kids.push(section("Experience"));
    const order = t.order && t.order.length ? t.order : master.roles.map(r => r.key);
    order.forEach((key, i) => {
      const r = master.roles.find(x => x.key === key); if (!r) return;
      kids.push(p([run(r.title + "  |  " + r.company, { bold: true }), run("\t" + r.dates, { color: MUTED, size: 19 })], { before: i ? 140 : 40, after: 10, tabs: rightTab, keepNext: true }));
      kids.push(p([run(r.place, { italics: true, color: MUTED, size: 19 })], { after: 40, keepNext: true }));
      const bl = (t.bullets && t.bullets[key] && t.bullets[key].length) ? t.bullets[key] : r.bullets;
      bl.forEach(b => kids.push(bullet(b)));
    });

    kids.push(section("Education"));
    master.education.forEach(e => {
      kids.push(p([run(e.degree, { bold: true }), run("\t" + e.dates, { color: MUTED, size: 19 })], { after: 0, tabs: rightTab, keepNext: true }));
      kids.push(p([run(e.school, { color: MUTED, size: 19 })], { after: 80 }));
    });

    kids.push(section("Certification"));
    master.certs.forEach(c => kids.push(p([run(c)], { after: 40 })));

    return new Document({
      creator: master.name, title: master.name + " CV",
      styles: { default: { document: { run: { font: FONT, size: 21 } } } },
      numbering: { config: [{ reference: "dots", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
      sections: [{ properties: { page: { size: { width: W, height: 16838 }, margin: { top: 900, bottom: 900, left: MARGIN, right: MARGIN } } }, children: kids }]
    });
  }
  if (typeof module !== "undefined" && module.exports) module.exports = buildCv; else root.buildCv = buildCv;
})(typeof window !== "undefined" ? window : this);
