/**
 * create_pptx_simple.js
 * Minimal: erzeugt "Imperialismus_Praesentation.pptx" mit ~13 Folien + Sprechernotizen.
 * Benötigt: Node.js und pptxgenjs
 *
 * Anwendung:
 * 1) Öffne PowerShell und navigiere zu einem Arbeitsordner (z. B. Desktop).
 * 2) mkdir imperialismus-pptx
 * 3) cd imperialismus-pptx
 * 4) npm init -y
 * 5) npm install pptxgenjs
 * 6) Lege diese Datei (create_pptx_simple.js) in den Ordner.
 * 7) node create_pptx_simple.js
 *
 * Die erzeugte Datei heißt: Imperialismus_Praesentation.pptx
 * Titelfolie: kein Name wird eingetragen (Platzhalter bleibt). Datum auf Titelfolie: 30.06.2026
 */

const pptxgen = require("pptxgenjs");

const NAVY = "16213E";
const NAVY_DEEP = "0F1830";
const RED = "C8102E";
const WHITE = "FFFFFF";
const OFFWHITE = "F4F6FA";
const SLATE = "5B6478";
const LIGHTLINE = "DCE1EC";

function addKicker(slide, text, color = RED) {
  slide.addText(text.toUpperCase(), {
    x: 0.6, y: 0.32, w: 8, h: 0.3, fontSize: 12, color, bold: true,
    fontFace: "Calibri", charSpacing: 2, margin: 0
  });
}

function addPageNum(slide, n, dark = false) {
  slide.addText(String(n).padStart(2, "0"), {
    x: 9.35, y: 5.28, w: 0.5, h: 0.3, fontSize: 9, color: dark ? "8B93A8" : "A7AEC2",
    align: "right", fontFace: "Calibri", margin: 0
  });
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Präsentation Generator";
  pres.title = "Imperialismus heute vs. früher";

  // Slide 1 — Title
  {
    const s = pres.addSlide();
    s.background = { color: NAVY_DEEP };
    s.addText("IMPERIALISMUS HEUTE VS. FRÜHER", { x: 0.7, y: 2.55, w: 8.6, h: 1.0, fontSize: 36, bold: true, color: WHITE, fontFace: "Cambria" });
    s.addText("Hat sich Imperialismus nur verändert?", { x: 0.7, y: 3.45, w: 8.6, h: 0.55, fontSize: 18, italic: true, color: "B9C2D9", fontFace: "Calibri" });
    s.addShape("rect", { x: 0.7, y: 4.35, w: 0.5, h: 0.02, fill: { color: RED } });
    s.addText("Name: ___________________", { x: 0.7, y: 4.55, w: 4.3, h: 0.3, fontSize: 12, color: "8B93A8", fontFace: "Calibri" });
    s.addText("30.06.2026", { x: 0.7, y: 4.85, w: 4.3, h: 0.3, fontSize: 12, color: "8B93A8", fontFace: "Calibri" });
    addPageNum(s, 1);
    s.addNotes("Begrüßen Sie das Publikum und stellen Sie die Leitfrage vor: 'Existiert Imperialismus heute noch oder hat er lediglich seine Form verändert?' Kurzankündigung der Agenda.");
  }

  // Slide 2 — Agenda
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addKicker(s, "Überblick");
    s.addText("Agenda", { x: 0.6, y: 0.6, w: 8, h: 0.6, fontSize: 32, bold: true, color: NAVY, fontFace: "Cambria" });
    const items = [
      "Definition Imperialismus",
      "Kolonialismus",
      "Imperialismus früher",
      "Imperialismus heute",
      "Politisches System der USA",
      "Einfluss der US-Tech-Giganten",
      "Motive",
      "Medien und Journalismus",
      "Fazit"
    ];
    let y = 1.6;
    items.forEach((it, idx) => {
      s.addText(`${String(idx + 1).padStart(2, "0")}  ${it}`, { x: 0.8, y: y, w: 8.6, h: 0.4, fontSize: 14, color: NAVY, fontFace: "Calibri" });
      y += 0.45;
    });
    addPageNum(s, 2);
    s.addNotes("Kurzer Überblick über Aufbau und Ablauf der Präsentation; Hinweis auf Vergleich früher vs. heute.");
  }

  // Slide 3 — Definition
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addKicker(s, "01 · Grundlagen");
    s.addText("Was ist Imperialismus?", { x: 0.6, y: 0.6, w: 8.5, h: 0.6, fontSize: 30, bold: true, color: NAVY, fontFace: "Cambria" });
    s.addShape("rect", { x: 0.6, y: 1.5, w: 4.5, h: 2.1, fill: { color: NAVY } });
    s.addText("Imperialismus bezeichnet das Bestreben eines Staates, seinen politischen, wirtschaftlichen oder militärischen Einfluss auf andere Länder auszudehnen.", { x: 0.95, y: 1.75, w: 3.85, h: 1.6, fontSize: 14, color: WHITE, fontFace: "Calibri" });
    s.addText("Direkte Herrschaft vs. indirekte Einflussnahme", { x: 0.6, y: 3.85, w: 4.5, h: 0.3, fontSize: 12, bold: true, color: RED, fontFace: "Calibri" });
    s.addText("Direkte Herrschaft = formelle Kontrolle (Kolonien). Indirekter Einfluss = wirtschaftlich, kulturell oder technologisch ohne formelle Eroberung.", { x: 0.6, y: 4.15, w: 4.5, h: 0.9, fontSize: 11.5, color: SLATE, fontFace: "Calibri" });
    s.addText("Beispiele: Großbritannien, Frankreich, Deutschland (um 1900)", { x: 5.6, y: 1.6, w: 3.6, h: 0.6, fontSize: 12, color: NAVY, fontFace: "Calibri" });
    addPageNum(s, 3);
    s.addNotes("Definition vorlesen und die Unterscheidung zwischen direkter vs. indirekter Einflussnahme erklären; historische Mächte nennen.");
  }

  // Slide 4 — Kolonialismus
  {
    const s = pres.addSlide();
    s.background = { color: OFFWHITE };
    addKicker(s, "02 · Grundlagen");
    s.addText("Kolonialismus", { x: 0.6, y: 0.6, w: 8.5, h: 0.6, fontSize: 30, bold: true, color: NAVY, fontFace: "Cambria" });
    s.addText("Kolonialismus ist die praktische Umsetzung imperialistischer Ziele: direkte Eroberung, Verwaltung und Ausbeutung fremder Gebiete.", { x: 0.6, y: 1.3, w: 8.8, h: 0.6, fontSize: 13.5, color: SLATE, fontFace: "Calibri" });
    s.addText("Typische Merkmale:", { x: 0.6, y: 2.1, w: 3.0, h: 0.3, fontSize: 13, bold: true, color: RED, fontFace: "Calibri" });
    const traits = ["Eroberung (militärisch)", "Ausbeutung (Rohstoffe)", "Unterdrückung der Bevölkerung"];
    let ty = 2.45;
    traits.forEach(t => { s.addText("— " + t, { x: 0.8, y: ty, w: 6.0, h: 0.3, fontSize: 12, color: NAVY, fontFace: "Calibri" }); ty += 0.38; });
    s.addText("Historische Beispiele: Indien (Großbritannien), Belgisch-Kongo, Deutsch-Südwestafrika", { x: 0.6, y: 4.0, w: 8.8, h: 0.5, fontSize: 11.5, color: SLATE, fontFace: "Calibri" });
    addPageNum(s, 4);
    s.addNotes("Kolonialismus als Umsetzung des Imperialismus: Eroberung, Rohstoffextraktion, Unterdrückung; Beispiele mit kurzen Kontext.");
  }

  // Slide 5 — Vergleich Früher vs Heute
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addKicker(s, "03–04 · Vergleich");
    s.addText("Imperialismus früher vs. heute", { x: 0.6, y: 0.6, w: 8.8, h: 0.6, fontSize: 28, bold: true, color: NAVY, fontFace: "Cambria" });

    s.addText("FRÜHER", { x: 0.8, y: 1.2, w: 3.6, h: 0.3, fontSize: 14, bold: true, color: NAVY });
    const left = ["Militärische Eroberung", "Kolonien", "Rohstoffgewinnung", "Direkte Kontrolle"];
    let ly = 1.5;
    left.forEach(t => { s.addText("— " + t, { x: 0.9, y: ly, w: 3.5, h: 0.3, fontSize: 12, color: NAVY, fontFace: "Calibri" }); ly += 0.45; });

    s.addText("HEUTE", { x: 5.0, y: 1.2, w: 3.6, h: 0.3, fontSize: 14, bold: true, color: RED });
    const right = ["Wirtschaftlicher Einfluss", "Technologie", "Finanzmärkte", "Internationale Unternehmen", "Digitale Plattformen"];
    let ry = 1.5;
    right.forEach(t => { s.addText("— " + t, { x: 5.1, y: ry, w: 3.5, h: 0.3, fontSize: 12, color: NAVY, fontFace: "Calibri" }); ry += 0.45; });

    addPageNum(s, 5);
    s.addNotes("Zentraler Vergleich: Mittel haben sich verändert (sichtbar → subtil), Zweck bleibt (Macht/Resourcen). Heute oft wirtschaftlich und technologisch.");
  }

  // Slide 6 — Politisches System der USA
  {
    const s = pres.addSlide();
    s.background = { color: OFFWHITE };
    addKicker(s, "05 · Hintergrund");
    s.addText("Politisches System der USA", { x: 0.6, y: 0.6, w: 8.8, h: 0.6, fontSize: 28, bold: true, color: NAVY, fontFace: "Cambria" });
    s.addText("Die USA sind eine Demokratie mit Gewaltenteilung in drei Gewalten:", { x: 0.6, y: 1.25, w: 8.8, h: 0.3, fontSize: 13, color: SLATE, fontFace: "Calibri" });

    s.addText("Exekutive – Präsident: führt Regierung, Oberbefehlshaber", { x: 0.8, y: 1.9, w: 8.0, h: 0.3, fontSize: 12, color: NAVY, fontFace: "Calibri" });
    s.addText("Legislative – Kongress (Senat & Repräsentantenhaus): Gesetze, Haushalt", { x: 0.8, y: 2.35, w: 8.0, h: 0.3, fontSize: 12, color: NAVY, fontFace: "Calibri" });
    s.addText("Judikative – Supreme Court: prüft Verfassungskonformität", { x: 0.8, y: 2.8, w: 8.0, h: 0.3, fontSize: 12, color: NAVY, fontFace: "Calibri" });

    s.addText("Warum global wichtig? Institutionelle Stabilität, wirtschaftliche Macht und militärische Kapazitäten erhöhen den Einfluss.", { x: 0.6, y: 3.5, w: 8.8, h: 0.6, fontSize: 11.5, italic: true, color: SLATE, fontFace: "Calibri" });
    addPageNum(s, 6);
    s.addNotes("Vereinfachte Darstellung der drei Gewalten und Erklärung, warum das politische System der USA den internationalen Einfluss fördert; Überleitung zu Tech-Konzernen.");
  }

  // Slide 7 — US-Tech-Giganten
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addKicker(s, "06 · Heutiger Einfluss");
    s.addText("US-Tech-Giganten als moderner Imperialismus?", { x: 0.6, y: 0.55, w: 8.9, h: 0.7, fontSize: 25, bold: true, color: NAVY, fontFace: "Cambria" });

    s.addText("Beispiele: Apple, Google, Microsoft, Amazon, Meta, NVIDIA", { x: 0.8, y: 1.5, w: 8.0, h: 0.4, fontSize: 12.5, color: NAVY, fontFace: "Calibri" });
    s.addText("Bereiche der Kontrolle:", { x: 0.8, y: 2.1, w: 8.0, h: 0.3, fontSize: 13, bold: true, color: RED, fontFace: "Calibri" });
    const areas = ["Kontrolle über Daten", "Digitale Infrastruktur & Cloud", "Künstliche Intelligenz", "Betriebssysteme & App-Stores"];
    let ay = 2.45;
    areas.forEach(a => { s.addText("— " + a, { x: 0.9, y: ay, w: 7.8, h: 0.3, fontSize: 12, color: NAVY, fontFace: "Calibri" }); ay += 0.45; });

    s.addText("Diskussionsfrage: Ersetzen digitale Unternehmen heute klassische Kolonialmächte?", { x: 0.8, y: 4.5, w: 8.0, h: 0.6, fontSize: 12, italic: true, color: SLATE, fontFace: "Calibri" });
    addPageNum(s, 7);
    s.addNotes("Vorstellung der großen Tech-Konzerne und der vier zentralen Einflussbereiche; Diskussionsfrage in den Raum stellen.");
  }

  // Slide 8 — Motive
  {
    const s = pres.addSlide();
    s.background = { color: OFFWHITE };
    addKicker(s, "07 · Hintergrund");
    s.addText("Motive des Imperialismus", { x: 0.6, y: 0.6, w: 8.8, h: 0.6, fontSize: 28, bold: true, color: NAVY, fontFace: "Cambria" });

    s.addText("Früher:", { x: 0.8, y: 1.4, w: 3.5, h: 0.3, fontSize: 13, bold: true, color: SLATE });
    const oldMotives = ["Rohstoffe", "Macht", "Nationalismus", "Militär"];
    let oy = 1.75;
    oldMotives.forEach(m => { s.addText("— " + m, { x: 0.9, y: oy, w: 3.5, h: 0.3, fontSize: 12, color: NAVY, fontFace: "Calibri" }); oy += 0.45; });

    s.addText("Heute:", { x: 5.0, y: 1.4, w: 3.5, h: 0.3, fontSize: 13, bold: true, color: RED });
    const newMotives = ["Daten", "Künstliche Intelligenz", "Wirtschaftswachstum", "Technologische Vorherrschaft", "Geopolitischer Einfluss"];
    let ny = 1.75;
    newMotives.forEach(m => { s.addText("— " + m, { x: 5.1, y: ny, w: 3.5, h: 0.3, fontSize: 12, color: NAVY, fontFace: "Calibri" }); ny += 0.45; });

    addPageNum(s, 8);
    s.addNotes("Vergleich der Motive: Früher Rohstoffe/Macht; heute Daten/KI und wirtschaftliche/technologische Vormachtstellung.");
  }

  // Slide 9 — Medien und Journalismus
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addKicker(s, "08 · Heutiger Einfluss");
    s.addText("Medien und Journalismus", { x: 0.6, y: 0.6, w: 8.8, h: 0.6, fontSize: 28, bold: true, color: NAVY, fontFace: "Cambria" });
    s.addText("Wie beeinflussen Medien Meinungen?", { x: 0.6, y: 1.25, w: 8.8, h: 0.35, fontSize: 14, italic: true, color: SLATE, fontFace: "Calibri" });

    s.addText("Klassischer Journalismus:", { x: 0.8, y: 1.8, w: 8.0, h: 0.3, fontSize: 13, bold: true, color: NAVY });
    const classic = ["Agenda Setting (Themen setzen)", "Framing (Deutung beeinflussen)", "Redaktionelle Standards & Verifikation"];
    let cY = 2.15;
    classic.forEach(p => { s.addText("— " + p, { x: 0.9, y: cY, w: 7.8, h: 0.3, fontSize: 11.5, color: SLATE, fontFace: "Calibri" }); cY += 0.38; });

    s.addText("Soziale Medien:", { x: 5.0, y: 1.8, w: 3.5, h: 0.3, fontSize: 13, bold: true, color: RED });
    const social = ["Algorithmen entscheiden Reichweite", "Schnelligkeit vor Verifikation", "Nutzer als Sender (Desinformation-Risiko)"];
    let sY = 2.15;
    social.forEach(p => { s.addText("— " + p, { x: 5.1, y: sY, w: 3.6, h: 0.3, fontSize: 11.5, color: "5B6478", fontFace: "Calibri" }); sY += 0.38; });

    addPageNum(s, 9);
    s.addNotes("Erklärung von Agenda Setting und Framing; Vergleich klassischer Medien vs. soziale Medien mit Chancen und Risiken.");
  }

  // Slide 10 — Fallbeispiele
  {
    const s = pres.addSlide();
    s.background = { color: OFFWHITE };
    addKicker(s, "Fallbeispiele");
    s.addText("Fallbeispiele aktueller Einflussnahme", { x: 0.6, y: 0.6, w: 8.8, h: 0.6, fontSize: 26, bold: true, color: NAVY, fontFace: "Cambria" });

    const cases = [
      ["TikTok-Debatte (USA vs. China)", "Streit um Datenkontrolle und nationale Sicherheit."],
      ["Google & Android", "Dominanz des mobilen Betriebssystems beeinflusst Milliarden Geräte."],
      ["Microsoft Cloud", "Regierungen und Firmen sind abhängig von US-Cloud-Infrastruktur."],
      ["Amazon Web Services", "AWS betreibt kritische Teile der digitalen Infrastruktur."]
    ];
    let cy = 1.6;
    for (let i = 0; i < cases.length; i++) {
      s.addText(`• ${cases[i][0]}`, { x: 0.8, y: cy, w: 8.6, h: 0.25, fontSize: 12.5, bold: true, color: NAVY, fontFace: "Calibri" });
      s.addText(cases[i][1], { x: 1.1, y: cy + 0.28, w: 8.2, h: 0.4, fontSize: 11, color: SLATE, fontFace: "Calibri" });
      cy += 0.9;
    }

    addPageNum(s, 10);
    s.addNotes("Vier kurze Fallbeispiele: TikTok (Geopolitik & Daten), Android (Plattformmacht), Microsoft Cloud und AWS (digitale Abhängigkeit).");
  }

  // Slide 11 — Diskussion
  {
    const s = pres.addSlide();
    s.background = { color: NAVY_DEEP };
    addKicker(s, "Diskussion", "E4536A");
    s.addText("Diskussion", { x: 0.6, y: 0.65, w: 8.8, h: 0.6, fontSize: 30, bold: true, color: WHITE, fontFace: "Cambria" });
    s.addText("Fragen an das Publikum", { x: 0.6, y: 1.3, w: 8.8, h: 0.35, fontSize: 14, italic: true, color: "B9C2D9", fontFace: "Calibri" });

    const q = [
      "Ist wirtschaftlicher Einfluss eine neue Form des Imperialismus?",
      "Haben Technologieunternehmen mehr Macht als manche Staaten?",
      "Sind soziale Medien ein politisches Instrument?"
    ];
    let qy = 2.05;
    q.forEach(t => {
      s.addShape("rect", { x: 0.6, y: qy, w: 8.8, h: 0.85, fill: { color: "1F2D4D" } });
      s.addText(t, { x: 1.0, y: qy + 0.18, w: 7.6, h: 0.45, fontSize: 14, color: WHITE, fontFace: "Calibri" });
      qy += 1.05;
    });

    addPageNum(s, 11, true);
    s.addNotes("Stelle die drei Fragen, gib dem Publikum Zeit zu antworten und sammle kurze Positionen.");
  }

  // Slide 12 — Fazit
  {
    const s = pres.addSlide();
    s.background = { color: NAVY_DEEP };
    addKicker(s, "Fazit", "E4536A");
    s.addText("Fazit", { x: 0.6, y: 0.65, w: 8.8, h: 0.6, fontSize: 32, bold: true, color: WHITE, fontFace: "Cambria" });

    const points = [
      "Imperialismus existiert weiterhin.",
      "Die Methoden haben sich verändert.",
      "Heute dominieren Wirtschaft, Technologie und Informationen.",
      "Staaten und Konzerne konkurrieren um globalen Einfluss."
    ];
    let py = 1.55;
    points.forEach(p => {
      s.addText("— " + p, { x: 0.8, y: py, w: 8.0, h: 0.4, fontSize: 15, color: WHITE, fontFace: "Calibri" });
      py += 0.6;
    });

    s.addText("„Ist der digitale Imperialismus gefährlicher als der klassische Imperialismus?“", { x: 0.95, y: 4.15, w: 8.1, h: 1.1, fontSize: 15, italic: true, bold: true, color: WHITE, fontFace: "Calibri" });

    addPageNum(s, 12, true);
    s.addNotes("Zusammenfassen: Imperialismus verändert, ist aber nicht verschwunden. Abschlussfrage offen lassen.");
  }

  // Slide 13 — Quellen
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addKicker(s, "Quellen");
    s.addText("Quellen", { x: 0.6, y: 0.6, w: 8.8, h: 0.6, fontSize: 30, bold: true, color: NAVY, fontFace: "Cambria" });

    const sources = [
      "Vereinte Nationen (UN)", "Weltbank", "OECD", "Encyclopaedia Britannica",
      "Statista", "Pew Research Center", "Harvard Business Review", "Fachliteratur und wissenschaftliche Artikel"
    ];
    let sx = 0.6, sy = 1.6;
    for (let i = 0; i < sources.length; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = sx + col * 4.6;
      const y = sy + row * 0.7;
      s.addText(sources[i], { x: x, y: y, w: 4.4, h: 0.4, fontSize: 12, color: NAVY, fontFace: "Calibri" });
    }

    s.addText("Vielen Dank für Ihre Aufmerksamkeit.", { x: 0.6, y: 5.0, w: 8.8, h: 0.35, fontSize: 12, italic: true, color: SLATE, fontFace: "Calibri" });
    addPageNum(s, 13);
    s.addNotes("Quellenarten nennen; Abschluss und Einladung zu Fragen.")
  }

  // Write file
  const outFile = "Imperialismus_Praesentation.pptx";
  try {
    await pres.writeFile({ fileName: outFile });
    console.log("PPTX erfolgreich erstellt:", outFile);
  } catch (err) {
    try {
      await pres.writeFile(outFile);
      console.log("PPTX erfolgreich erstellt (alternative writeFile):", outFile);
    } catch (err2) {
      console.error("Fehler beim Speichern der PPTX-Datei:", err2);
    }
  }
})();
