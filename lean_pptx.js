const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Lean Thinking & Amélioration Continue";

// ── Palette
const C = {
  navy:   "0F2D4A",
  blue:   "1E6FA8",
  lblue:  "DBEAFE",
  teal:   "0B7A6A",
  lteal:  "CCFBF1",
  green:  "166534",
  lgreen: "DCFCE7",
  orange: "C45200",
  lorange:"FEF3C7",
  red:    "991B1B",
  lred:   "FEE2E2",
  amber:  "92400E",
  lamber: "FEF3C7",
  purple: "5B21B6",
  lpurp:  "EDE9FE",
  gray:   "374151",
  lgray:  "F3F4F6",
  dgray:  "6B7280",
  white:  "FFFFFF",
  black:  "111827",
};

const ms = () => ({ type:"outer", blur:5, offset:2, angle:135, color:"000000", opacity:0.08 });

// helper: add a section card
function card(slide, x, y, w, h, opts={}) {
  const { fill=C.lgray, borderColor=null, borderW=0, radius=0.08 } = opts;
  const shape = borderColor
    ? { fill:{color:fill}, line:{color:borderColor, width:borderW||1.5} }
    : { fill:{color:fill} };
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: radius, ...shape, shadow: ms() });
}

function accentCard(slide, x, y, w, h, accentColor, fill=C.white) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill:{color:fill}, line:{color:"E5E7EB", width:0.5}, shadow:ms() });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w:0.06, h, fill:{color:accentColor} });
}

function badge(slide, x, y, text, bg, fg) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:1.1, h:0.22, rectRadius:0.04, fill:{color:bg} });
  slide.addText(text, { x, y, w:1.1, h:0.22, fontSize:8, bold:true, color:fg, align:"center", valign:"middle" });
}

function sectionTag(slide, text) {
  slide.addText(text, { x:0.4, y:0.15, w:9.2, h:0.2, fontSize:8, bold:true, color:C.dgray, charSpacing:3 });
}

function slideTitle(slide, text, y=0.38) {
  slide.addText(text, { x:0.4, y, w:9.2, h:0.55, fontSize:26, bold:true, color:C.black, fontFace:"Calibri" });
}

function sub(slide, text, y=0.95) {
  slide.addText(text, { x:0.4, y, w:9.2, h:0.25, fontSize:11, color:C.dgray, italic:true });
}

function divider(slide, y=1.22) {
  slide.addShape(pres.shapes.LINE, { x:0.4, y, w:9.2, h:0, line:{color:"E5E7EB", width:1} });
}

function sm(slide, x, y, w, h, text, opts={}) {
  slide.addText(text, { x, y, w, h, fontSize:11, color:C.gray, wrap:true, valign:"top", ...opts });
}

function highlightBox(slide, x, y, w, h, text, borderColor, bgColor, textColor) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill:{color:bgColor}, line:{color:bgColor, width:0} });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w:0.06, h, fill:{color:borderColor} });
  slide.addText(text, { x:x+0.12, y:y+0.06, w:w-0.2, h:h-0.12, fontSize:11, color:textColor, wrap:true, valign:"top" });
}

// ─────────────────────────────────────────────
// SLIDE 0 — COUVERTURE
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.navy };
  // Title
  s.addText("Lean Thinking &", { x:0.5, y:0.6, w:9, h:0.7, fontSize:38, bold:true, color:C.white, fontFace:"Calibri" });
  s.addText("Amélioration Continue", { x:0.5, y:1.25, w:9, h:0.7, fontSize:38, bold:true, color:"60A5FA", fontFace:"Calibri" });
  s.addText("Penser et agir KAIZEN au quotidien", { x:0.5, y:2.05, w:9, h:0.3, fontSize:14, color:"93C5FD", italic:true });
  s.addShape(pres.shapes.LINE, { x:0.5, y:2.45, w:9, h:0, line:{color:"1E4D8C", width:1} });
  // Stats
  const stats = [
    ["-40%", "Problèmes récurrents"],
    ["+2h/j", "De valeur ajoutée"],
    ["0 Dh", "D'investissement"],
  ];
  stats.forEach(([num, lbl], i) => {
    const x = 0.5 + i * 3.17;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:2.75, w:2.8, h:1.5, rectRadius:0.1, fill:{color:"162B42"} });
    s.addText(num, { x, y:2.9, w:2.8, h:0.7, fontSize:28, bold:true, color:"60A5FA", align:"center" });
    s.addText(lbl, { x, y:3.55, w:2.8, h:0.3, fontSize:10, color:"93C5FD", align:"center" });
  });
  s.addText("Formation · 20 minutes · Ingénieurs & Responsables", { x:0.5, y:4.6, w:9, h:0.25, fontSize:9, color:"4B6B8A", align:"center" });
}

// ─────────────────────────────────────────────
// SLIDE 1 — POURQUOI CETTE FORMATION ?
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "INTRODUCTION");
  slideTitle(s, "Pourquoi cette formation ?");
  sub(s, "Le travail non structuré cache un coût élevé invisible");
  divider(s);

  // Iceberg concept — left
  card(s, 0.4, 1.35, 4.4, 2.7, { fill:"EFF6FF" });
  s.addText("LE COÛT CACHÉ DU TRAVAIL", { x:0.5, y:1.45, w:4.2, h:0.25, fontSize:9, bold:true, color:C.blue, align:"center", charSpacing:2 });
  // Visible bar
  s.addShape(pres.shapes.RECTANGLE, { x:1.6, y:1.82, w:1.8, h:0.45, fill:{color:"3B82F6"} });
  s.addText("Visible — 20%", { x:1.6, y:1.82, w:1.8, h:0.45, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle" });
  s.addShape(pres.shapes.LINE, { x:0.5, y:2.27, w:4.2, h:0, line:{color:"93C5FD", width:1.5, dashType:"dash"} });
  s.addText("── Ligne de flottaison ──", { x:0.5, y:2.27, w:4.2, h:0.18, fontSize:8, color:"3B82F6", align:"center" });
  // Hidden bar (trapeze-like wider)
  s.addShape(pres.shapes.RECTANGLE, { x:0.9, y:2.45, w:3.3, h:1.1, fill:{color:"EF4444"} });
  s.addText("Caché — 80%\nErreurs · Refaites · Attentes\nDéplacements inutiles", { x:0.9, y:2.45, w:3.3, h:1.1, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle" });

  // Right side — problem list
  card(s, 5.1, 1.35, 4.5, 2.7, { fill:C.lgray });
  s.addText("Ce qu'on vit au quotidien sans Lean", { x:5.2, y:1.42, w:4.3, h:0.28, fontSize:10, bold:true, color:C.black });
  const problems = [
    "Problème résolu → revient 2 semaines après",
    "Technicien attend une pièce introuvable",
    "Réunion sans ordre du jour ni décision",
    "Même erreur commise 3 fois par des équipes différentes",
    "Rapport écrit mais jamais lu ni suivi",
  ];
  problems.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:5.15, y:1.78 + i*0.45, w:0.06, h:0.3, fill:{color:C.red} });
    sm(s, 5.3, 1.75 + i*0.45, 4.1, 0.35, p);
  });

  // Bottom
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.2, w:9.2, h:0.6, fill:{color:"FEF2F2"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.2, w:0.06, h:0.6, fill:{color:C.red} });
  s.addText("Le coût caché représente 20 à 40% du temps de travail — et ça peut changer dès cette semaine.", { x:0.55, y:4.24, w:9.0, h:0.5, fontSize:12, bold:true, color:"991B1B", valign:"middle", wrap:true });
}

// ─────────────────────────────────────────────
// SLIDE 2 — LE CERCLE VICIEUX
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "INTRODUCTION · LA PROBLÉMATIQUE");
  slideTitle(s, "Sans amélioration structurée : le cercle vicieux");
  divider(s, 1.0);

  const nodes1 = [
    ["Problème","FEE2E2","991B1B"],
    ["Pansement rapide","FEF3C7","92400E"],
    ["Le problème revient","DBEAFE","1E40AF"],
    ["Stress équipe","F3E8FF","6B21A8"],
  ];
  nodes1.forEach(([txt, bg, fg], i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.4+i*2.3, y:1.15, w:2.0, h:0.55, rectRadius:0.06, fill:{color:bg} });
    s.addText(txt, { x:0.4+i*2.3, y:1.15, w:2.0, h:0.55, fontSize:11, bold:true, color:fg, align:"center", valign:"middle" });
    if (i < 3) s.addText("→", { x:2.35+i*2.3, y:1.3, w:0.35, h:0.25, fontSize:18, color:"9CA3AF", align:"center" });
  });
  s.addText("↓", { x:4.6, y:1.72, w:0.8, h:0.4, fontSize:22, color:"EF4444", align:"center" });
  const nodes2 = [
    ["Fatigue / Erreur","FECACA","7F1D1D"],
    ["Nouveau problème","FDE68A","78350F"],
    ["Morale dégradée","BBF7D0","166534"],
    ["Perte de compétences","FCA5A5","991B1B"],
  ];
  nodes2.forEach(([txt, bg, fg], i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.4+i*2.3, y:2.15, w:2.0, h:0.55, rectRadius:0.06, fill:{color:bg} });
    s.addText(txt, { x:0.4+i*2.3, y:2.15, w:2.0, h:0.55, fontSize:11, bold:true, color:fg, align:"center", valign:"middle" });
    if (i < 3) s.addText("→", { x:2.35+i*2.3, y:2.3, w:0.35, h:0.25, fontSize:18, color:"9CA3AF", align:"center" });
  });

  // Key insight
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:2.95, w:9.2, h:0.7, fill:{color:"FEF2F2"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:2.95, w:0.06, h:0.7, fill:{color:C.red} });
  s.addText("Faire de l'amélioration continue réduit considérablement les problèmes qui se répètent.\nChaque problème non traité à la racine revient en moyenne 3 fois — soit 3× le temps perdu.", { x:0.56, y:2.99, w:9.0, h:0.6, fontSize:11, color:"991B1B", wrap:true, valign:"middle" });

  // With Lean box
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.4, y:3.82, w:9.2, h:0.95, rectRadius:0.08, fill:{color:"F0FDF4"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:3.82, w:0.06, h:0.95, fill:{color:C.teal} });
  s.addText("✓  Avec KAIZEN : On identifie → On analyse → On standardise → Le problème ne revient plus.", { x:0.56, y:3.92, w:9.0, h:0.4, fontSize:12, bold:true, color:C.teal, valign:"middle" });
  s.addText("Résultat mesuré : -40% de problèmes récurrents en 6 mois avec PDCA + 5 Pourquoi systématique.", { x:0.56, y:4.28, w:9.0, h:0.35, fontSize:10, color:"166534", valign:"middle" });
}

// ─────────────────────────────────────────────
// SLIDE 3 — QU'EST-CE QUE LE LEAN ?
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 1 · LEAN THINKING");
  slideTitle(s, "Qu'est-ce que le Lean Management ?");
  sub(s, "Lean = Maigre · Né chez Toyota TPS · Womack & Jones 1990");
  divider(s);

  // Definition
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.4, y:1.35, w:9.2, h:0.75, rectRadius:0.08, fill:{color:"F0FDF4"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.35, w:0.06, h:0.75, fill:{color:C.teal} });
  s.addText("Définition : Produire exactement ce que le client veut, en éliminant tout ce qui ne crée pas de valeur.", { x:0.56, y:1.42, w:9.0, h:0.6, fontSize:13, bold:true, color:C.teal, valign:"middle", wrap:true });

  // 5 Principles flow
  s.addText("LES 5 PRINCIPES", { x:0.4, y:2.25, w:9.2, h:0.22, fontSize:8, bold:true, color:C.dgray, charSpacing:3 });
  const principles = [["Valeur","DBEAFE","1E40AF"],["Flux valeur","DCFCE7","166534"],["Flux continu","FEF3C7","92400E"],["Flux tiré","EDE9FE","5B21B6"],["Perfection","FCE7F3","9F1239"]];
  principles.forEach(([txt, bg, fg], i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.4+i*1.88, y:2.5, w:1.6, h:0.5, rectRadius:0.06, fill:{color:bg} });
    s.addText(txt, { x:0.4+i*1.88, y:2.5, w:1.6, h:0.5, fontSize:11, bold:true, color:fg, align:"center", valign:"middle" });
    if (i<4) s.addText("→", { x:2.0+i*1.88, y:2.6, w:0.3, h:0.3, fontSize:14, color:"9CA3AF", align:"center" });
  });

  // Stats
  const stats = [["-40%","Temps de cycle","3B82F6"],["-30%","Coûts opérationnels","22C55E"],["-50%","Défauts qualité","EF4444"],["+25%","Productivité terrain","F59E0B"]];
  stats.forEach(([num, lbl, col], i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.4+i*2.32, y:3.25, w:2.1, h:1.0, rectRadius:0.08, fill:{color:C.lgray} });
    s.addText(num, { x:0.4+i*2.32, y:3.32, w:2.1, h:0.5, fontSize:26, bold:true, color:col, align:"center" });
    s.addText(lbl, { x:0.4+i*2.32, y:3.78, w:2.1, h:0.35, fontSize:10, color:C.dgray, align:"center" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.42, w:9.2, h:0.55, fill:{color:"EFF6FF"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.42, w:0.06, h:0.55, fill:{color:C.blue} });
  s.addText("Toyota a réduit ses coûts de production de 50% en 10 ans grâce au TPS — sans investissement majeur, juste en éliminant les gaspillages.", { x:0.56, y:4.46, w:9.0, h:0.45, fontSize:10, color:"1E40AF", valign:"middle", wrap:true });
}

// ─────────────────────────────────────────────
// SLIDE 4 — KAIZEN vs KAIKAKU
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 1 · PRIORITÉ KAIZEN");
  slideTitle(s, "KAIZEN — Notre priorité absolue");
  sub(s, "改善 · Amélioration continue · Masaaki Imai · Toyota Production System 1986");
  divider(s);

  // KAIZEN card — dominant, left large
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.35, w:5.6, h:2.8, fill:{color:"F0FDF4"}, line:{color:"22C55E", width:2} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.35, w:0.08, h:2.8, fill:{color:"22C55E"} });
  s.addText("KAIZEN — NOTRE THÈME", { x:0.6, y:1.42, w:5.2, h:0.28, fontSize:11, bold:true, color:"166534", charSpacing:2 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.65, y:1.72, w:1.3, h:0.25, rectRadius:0.04, fill:{color:"DCFCE7"} });
  s.addText("RECOMMANDÉ", { x:0.65, y:1.72, w:1.3, h:0.25, fontSize:8, bold:true, color:"166534", align:"center", valign:"middle" });
  s.addText("Petites améliorations quotidiennes impliquant TOUS les niveaux sans investissement.", { x:0.6, y:2.05, w:5.1, h:0.4, fontSize:11, color:C.gray, wrap:true });
  const ksteps = ["Observer : noter 3 gaspillages/jour (5G)","Analyser : 5 Pourquoi en 30 min","Améliorer : tester une solution cette semaine","Standardiser : partager si ça marche"];
  ksteps.forEach((st, i) => {
    s.addShape(pres.shapes.OVAL, { x:0.6, y:2.52+i*0.38, w:0.22, h:0.22, fill:{color:"22C55E"} });
    s.addText(String(i+1), { x:0.6, y:2.52+i*0.38, w:0.22, h:0.22, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle" });
    sm(s, 0.9, 2.51+i*0.38, 4.9, 0.32, st);
  });

  // KAIKAKU card — right, smaller, muted
  s.addShape(pres.shapes.RECTANGLE, { x:6.2, y:1.35, w:3.4, h:2.8, fill:{color:"FFF7ED"}, line:{color:"FCA5A5", width:1} });
  s.addShape(pres.shapes.RECTANGLE, { x:6.2, y:1.35, w:0.06, h:2.8, fill:{color:"F87171"} });
  s.addText("KAIKAKU", { x:6.35, y:1.42, w:3.1, h:0.28, fontSize:11, bold:true, color:"C45200" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:6.35, y:1.72, w:1.5, h:0.25, rectRadius:0.04, fill:{color:"FEE2E2"} });
  s.addText("INVESTISSEMENT REQUIS", { x:6.35, y:1.72, w:1.5, h:0.25, fontSize:7, bold:true, color:"991B1B", align:"center", valign:"middle" });
  s.addText("Transformation radicale, ponctuelle, décidée au niveau direction.", { x:6.35, y:2.08, w:3.0, h:0.45, fontSize:10, color:C.dgray, wrap:true });
  const kk = ["Coût élevé","Résultat incertain","Décision direction","Mois / années"];
  kk.forEach((t, i) => {
    s.addText("–  " + t, { x:6.35, y:2.6+i*0.34, w:3.0, h:0.28, fontSize:10, color:"9CA3AF" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.3, w:9.2, h:0.7, fill:{color:"F0FDF4"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.3, w:0.06, h:0.7, fill:{color:C.teal} });
  s.addText("Message clé : Vous n'avez pas besoin d'un grand projet. Une petite amélioration chaque jour par chaque équipe = des centaines de milliers de Dh économisés par an, sans budget.", { x:0.56, y:4.34, w:9.0, h:0.6, fontSize:11, bold:true, color:C.teal, valign:"middle", wrap:true });
}

// ─────────────────────────────────────────────
// SLIDE 5 — VA / NVA
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 1 · VALEUR AJOUTÉE");
  slideTitle(s, "Ce que le client paie — ce qu'il ne paie pas");
  sub(s, "Cas concret : Production industrielle · Relation Client–Fournisseur de service");
  divider(s);

  // Bar
  s.addText("RÉPARTITION RÉELLE D'UNE JOURNÉE D'INTERVENTION (EXEMPLE ATELIER)", { x:0.4, y:1.3, w:9.2, h:0.2, fontSize:8, bold:true, color:C.dgray, charSpacing:2 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.55, w:2.76, h:0.35, fill:{color:"166534"} });
  s.addText("VA · 30%", { x:0.4, y:1.55, w:2.76, h:0.35, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle" });
  s.addShape(pres.shapes.RECTANGLE, { x:3.16, y:1.55, w:2.30, h:0.35, fill:{color:"D97706"} });
  s.addText("NVA non supp · 25%", { x:3.16, y:1.55, w:2.30, h:0.35, fontSize:8, bold:true, color:C.white, align:"center", valign:"middle" });
  s.addShape(pres.shapes.RECTANGLE, { x:5.46, y:1.55, w:4.14, h:0.35, fill:{color:"DC2626"} });
  s.addText("NVA supprimable · 45%", { x:5.46, y:1.55, w:4.14, h:0.35, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle" });

  // 3 columns
  const cats = [
    { label:"VA — Valeur Ajoutée", badge:"LE CLIENT PAIE", bg:"DCFCE7", bc:"22C55E", fc:"166534", body:"Transforme le produit/service vers l'objectif final.", example:"Réaction chimique · Filtration · Contrôle qualité · Transport effectif · Manutention productive", action:"OPTIMISER\nAugmenter le temps passé ici", aBg:"DCFCE7", aFg:"166534" },
    { label:"NVA non supprimable", badge:"NÉCESSAIRE MAIS NON PAYÉ", bg:"FEF3C7", bc:"F59E0B", fc:"92400E", body:"Contraintes réglementaires ou techniques inévitables.", example:"Arrêts réglementaires · EPI obligatoires · Consignation machine · Briefing sécurité · Permis de travail", action:"RÉDUIRE\nOptimiser sans supprimer", aBg:"FEF3C7", aFg:"92400E" },
    { label:"NVA supprimable", badge:"PUR GASPILLAGE", bg:"FEE2E2", bc:"EF4444", fc:"991B1B", body:"N'ajoute aucune valeur. À éliminer en priorité absolue.", example:"Attente pièce introuvable · Déplacements inutiles · Réunions sans agenda · Retraitement hors specs", action:"SUPPRIMER\nPriorité absolue → cible cette semaine", aBg:"FEE2E2", aFg:"991B1B" },
  ];
  cats.forEach((c, i) => {
    const x = 0.4 + i * 3.12;
    s.addShape(pres.shapes.RECTANGLE, { x, y:2.05, w:2.95, h:2.95, fill:{color:C.white}, line:{color:c.bc, width:1} });
    s.addShape(pres.shapes.RECTANGLE, { x, y:2.05, w:2.95, h:0.25, fill:{color:c.bc} });
    s.addText(c.badge, { x, y:2.05, w:2.95, h:0.25, fontSize:8, bold:true, color:C.white, align:"center", valign:"middle" });
    s.addText(c.label, { x:x+0.08, y:2.34, w:2.8, h:0.28, fontSize:10, bold:true, color:c.fc });
    sm(s, x+0.08, 2.64, 2.8, 0.4, c.body, { fontSize:10 });
    s.addText("Exemples :", { x:x+0.08, y:3.06, w:2.8, h:0.18, fontSize:9, bold:true, color:C.dgray });
    sm(s, x+0.08, 3.22, 2.8, 0.6, c.example, { fontSize:9 });
    s.addShape(pres.shapes.RECTANGLE, { x, y:3.85, w:2.95, h:0.55, fill:{color:c.aBg} });
    s.addText(c.action, { x:x+0.06, y:3.87, w:2.85, h:0.5, fontSize:9, bold:true, color:c.aFg, align:"center", valign:"middle" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.12, w:9.2, h:0.38, fill:{color:"FFFBEB"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.12, w:0.06, h:0.38, fill:{color:C.amber} });
  s.addText("Si on réduit les NVA-S de 45% → 20% : gain de ~2h/personne/jour = 10h/semaine par équipe de 5 — sans recruter.", { x:0.56, y:5.15, w:9.0, h:0.32, fontSize:10, bold:true, color:C.amber, valign:"middle", wrap:true });
}

// ─────────────────────────────────────────────
// SLIDE 6 — LES 7 MUDA
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 1 · LES 3 FORMES DE GASPILLAGE");
  slideTitle(s, "Les 3 formes + les 7 MUDA à éliminer");
  sub(s, "MUDA = 無駄 · Taiichi Ohno · Toyota · Identifier, nommer, éliminer");
  divider(s);

  // 3 types top
  const types3 = [
    ["MUDA","Gaspillage","Activités sans valeur à supprimer","FEE2E2","DC2626"],
    ["MURI","Surcharge","Pousser au-delà des limites raisonnables","FEF3C7","D97706"],
    ["MURA","Variabilité","Irrégularité dans les processus","EDE9FE","7C3AED"],
  ];
  types3.forEach(([t, n, d, bg, fg], i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.4+i*3.12, y:1.32, w:2.9, h:0.75, rectRadius:0.06, fill:{color:bg} });
    s.addText(t, { x:0.4+i*3.12, y:1.35, w:2.9, h:0.28, fontSize:14, bold:true, color:fg, align:"center" });
    s.addText(n+" — "+d, { x:0.4+i*3.12, y:1.6, w:2.9, h:0.45, fontSize:9, color:fg, align:"center", wrap:true });
  });

  // 7+1 MUDA grid — 4 per row
  s.addText("LES 7 MUDA — Identifier 3 par jour par personne", { x:0.4, y:2.2, w:9.2, h:0.22, fontSize:8, bold:true, color:C.dgray, charSpacing:2 });
  const mudas = [
    ["📦","Surproduction","Produire plus que la demande"],
    ["⏳","Attente","Technicien attend pièce/autorisation"],
    ["🚚","Transport inutile","Déplacer sans valeur ajoutée"],
    ["🔄","Traitement excessif","Plus de soin que nécessaire"],
    ["📊","Stocks excessifs","Pièces qui dorment en stock"],
    ["🚶","Mouvements inutiles","Chercher outil, faire des allers-retours"],
    ["❌","Défauts/Retouches","Refaire une tâche = double coût"],
    ["🎯","Sous-utilisation talent","Idées ignorées, compétences non exploitées"],
  ];
  mudas.forEach(([icon, name, desc], i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const x = 0.4 + col * 2.35;
    const y = 2.5 + row * 0.95;
    const isLast = i === 7;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:2.2, h:0.85, rectRadius:0.06, fill:{color: isLast ? "F0FDF4" : C.white}, line:{color: isLast ? "22C55E" : "FCA5A5", width:1.5} });
    s.addText(icon, { x, y:y+0.02, w:0.5, h:0.5, fontSize:18, align:"center" });
    s.addText(name, { x:x+0.48, y:y+0.04, w:1.65, h:0.25, fontSize:10, bold:true, color: isLast ? C.teal : C.red });
    s.addText(desc, { x:x+0.48, y:y+0.28, w:1.65, h:0.3, fontSize:9, color:C.dgray, wrap:true });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.12, w:9.2, h:0.38, fill:{color:"F0FDF4"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.12, w:0.06, h:0.38, fill:{color:C.teal} });
  s.addText("Plan concret : Identifier 3 gaspillages/jour × 5 personnes = 15 MUDA/semaine → 600 gaspillages/an éliminés → gain de plusieurs semaines de travail.", { x:0.56, y:5.15, w:9.0, h:0.32, fontSize:10, bold:true, color:C.teal, valign:"middle", wrap:true });
}

// ─────────────────────────────────────────────
// SLIDE 7 — PDCA
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 2 · MÉTHODES");
  slideTitle(s, "PDCA — La roue de l'amélioration continue");
  sub(s, "W. Edwards Deming 1950 · Cycle Shewhart · Base du Kaizen Toyota · GE, Boeing");
  divider(s);

  const pdca = [
    ["P","Plan","1D4ED8","DBEAFE","Identifier pb · Fixer objectif\nChercher causes · Définir solution"],
    ["D","Do","166534","DCFCE7","Exécuter le plan\nTester sur périmètre limité"],
    ["C","Check","92400E","FEF3C7","Évaluer les résultats\nMesurer écart objectif/réel"],
    ["A","Act","991B1B","FEE2E2","Capitaliser si succès\nStandardiser · Recommencer si échec"],
  ];
  pdca.forEach(([ltr, name, fg, bg, body], i) => {
    const x = 0.4 + i * 2.32;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.35, w:2.1, h:2.7, rectRadius:0.08, fill:{color:bg}, line:{color:fg, width:1} });
    s.addShape(pres.shapes.OVAL, { x:x+0.7, y:1.45, w:0.7, h:0.7, fill:{color:fg} });
    s.addText(ltr, { x:x+0.7, y:1.45, w:0.7, h:0.7, fontSize:22, bold:true, color:C.white, align:"center", valign:"middle" });
    s.addText(name, { x, y:2.2, w:2.1, h:0.3, fontSize:13, bold:true, color:fg, align:"center" });
    s.addText(body, { x:x+0.1, y:2.55, w:1.9, h:0.8, fontSize:10, color:fg, align:"center", wrap:true });
  });

  // Case concret
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.22, w:9.2, h:1.1, fill:{color:"EFF6FF"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.22, w:0.06, h:1.1, fill:{color:C.blue} });
  s.addText("Application concrète — Intervention pompe :", { x:0.56, y:4.28, w:9.0, h:0.25, fontSize:10, bold:true, color:"1E40AF" });
  s.addText("P : Temps = 4h, Objectif = 2h30 · Causes : outillage éparpillé, pas de procédure\nD : Chariot d'outillage dédié + check-list → test sur 2 interventions\nC : Résultat = 2h45 (proche objectif)  ·  A : Standardiser pour toutes les équipes", { x:0.56, y:4.55, w:9.0, h:0.72, fontSize:10, color:C.blue, wrap:true, lineSpacingMultiple:1.2 });
}

// ─────────────────────────────────────────────
// SLIDE 8 — DMAIC
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 2 · MÉTHODES");
  slideTitle(s, "DMAIC — Pour les problèmes complexes");
  sub(s, "Bill Smith · Motorola 1986 · Six Sigma · Adopté par GE, Boeing, Honeywell");
  divider(s);

  const phases = [
    ["D","Définir","1D4ED8","DBEAFE","Problème clairement défini\nQQOQCCP · CTQ · Charte"],
    ["M","Mesurer","166534","DCFCE7","Collecter données actuelles\nBaseline · Feuille relevé"],
    ["A","Analyser","92400E","FEF3C7","Causes racines\n5 Pourquoi · Ishikawa · Pareto"],
    ["I","Améliorer","5B21B6","EDE9FE","Générer solutions\nBrainstorming · SMED"],
    ["C","Contrôler","991B1B","FEE2E2","Pérenniser les gains\nMgmt visuel · Standards"],
  ];
  phases.forEach(([ltr, name, fg, bg, body], i) => {
    const x = 0.35 + i*1.88;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.35, w:1.72, h:2.2, rectRadius:0.07, fill:{color:bg} });
    s.addShape(pres.shapes.OVAL, { x:x+0.5, y:1.42, w:0.7, h:0.7, fill:{color:fg} });
    s.addText(ltr, { x:x+0.5, y:1.42, w:0.7, h:0.7, fontSize:20, bold:true, color:C.white, align:"center", valign:"middle" });
    s.addText(name, { x, y:2.15, w:1.72, h:0.28, fontSize:11, bold:true, color:fg, align:"center" });
    s.addText(body, { x:x+0.08, y:2.45, w:1.57, h:0.85, fontSize:9, color:fg, wrap:true, align:"center" });
  });

  // Comparison
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:3.7, w:4.4, h:0.95, fill:{color:C.lgray}, line:{color:"E5E7EB", width:0.5} });
  s.addText("PDCA", { x:0.5, y:3.75, w:4.2, h:0.25, fontSize:10, bold:true, color:C.blue });
  s.addText("Quotidien · Cycle court 1 semaine · Données simples", { x:0.5, y:4.0, w:4.2, h:0.5, fontSize:10, color:C.gray, wrap:true });
  s.addShape(pres.shapes.RECTANGLE, { x:5.0, y:3.7, w:4.6, h:0.95, fill:{color:C.lgray}, line:{color:"E5E7EB", width:0.5} });
  s.addText("DMAIC", { x:5.1, y:3.75, w:4.4, h:0.25, fontSize:10, bold:true, color:C.purple });
  s.addText("Projet structuré · 1–3 mois · Données rigoureuses · Chef de projet", { x:5.1, y:4.0, w:4.4, h:0.5, fontSize:10, color:C.gray, wrap:true });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.82, w:9.2, h:0.45, fill:{color:"EFF6FF"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.82, w:0.06, h:0.45, fill:{color:C.blue} });
  s.addText("Commencez par PDCA → passez à DMAIC si le problème résiste après 2 cycles.", { x:0.56, y:4.87, w:9.0, h:0.35, fontSize:11, bold:true, color:"1E40AF", valign:"middle" });
}

// ─────────────────────────────────────────────
// SLIDE 9 — 8D
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 2 · MÉTHODES");
  slideTitle(s, "8D — Résolution en équipe pour incidents graves");
  sub(s, "Ford Motor Company 1987 · Team Oriented Problem Solving · Industrie automobile mondiale");
  divider(s);

  const d8 = [
    ["D1","Constituer l'équipe","Groupe multidisciplinaire · chef projet","3B82F6","DBEAFE"],
    ["D2","Définir le problème","Description précise : Quoi, Quand, Où, Combien","3B82F6","DBEAFE"],
    ["D3","Action provisoire","Stopper l'hémorragie immédiatement","22C55E","DCFCE7"],
    ["D4","Causes racines","5 Pourquoi · Ishikawa · Valider par les faits","22C55E","DCFCE7"],
    ["D5","Qualifier solutions","Impact vs faisabilité","F59E0B","FEF3C7"],
    ["D6","Implémenter actions","Plan · Responsables · Délais · Suivi","F59E0B","FEF3C7"],
    ["D7","Actions préventives","Empêcher récurrence · Standards mis à jour","EF4444","FEE2E2"],
    ["D8","Féliciter l'équipe","Capitaliser · Partager leçons · Reconnaître","8B5CF6","EDE9FE"],
  ];
  d8.forEach(([code, name, desc, fg, bg], i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const x = 0.4 + col * 2.35;
    const y = 1.35 + row * 1.2;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:2.2, h:1.05, fill:{color:bg}, line:{color:fg, width:0.8} });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.05, h:1.05, fill:{color:fg} });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:x+0.12, y:y+0.06, w:0.42, h:0.24, rectRadius:0.04, fill:{color:fg} });
    s.addText(code, { x:x+0.12, y:y+0.06, w:0.42, h:0.24, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle" });
    s.addText(name, { x:x+0.6, y:y+0.06, w:1.52, h:0.28, fontSize:10, bold:true, color:fg });
    s.addText(desc, { x:x+0.12, y:y+0.38, w:2.0, h:0.55, fontSize:9, color:C.gray, wrap:true });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.82, w:9.2, h:0.45, fill:{color:"EFF6FF"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.82, w:0.06, h:0.45, fill:{color:C.blue} });
  s.addText("Quand utiliser : Incident grave · récurrent · fort impact · Durée : 3 jours à 2 semaines", { x:0.56, y:4.87, w:9.0, h:0.35, fontSize:11, bold:true, color:"1E40AF", valign:"middle" });
}

// ─────────────────────────────────────────────
// SLIDE 10 — QQOQCCP
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 3 · OUTILS TERRAIN");
  slideTitle(s, "QQOQCCP — Comprendre avant d'agir");
  sub(s, "Méthode journalistique · Adaptée par le Lean · Phase Définir (D) · 30 minutes max");
  divider(s);

  const qs = [
    ["Q","Qui ?","Technicien maintenance","1D4ED8","DBEAFE"],
    ["Q","Quoi ?","Arrêt pompe ligne 2","1D4ED8","DBEAFE"],
    ["O","Où ?","Atelier, ligne 2","1D4ED8","DBEAFE"],
    ["Q","Quand ?","Lundis matin au démarrage","1D4ED8","DBEAFE"],
    ["C","Comment ?","Vibrations + arrêt auto","22C55E","DCFCE7"],
    ["C","Combien ?","3h d'arrêt · 12 fois/mois","22C55E","DCFCE7"],
    ["P","Pourquoi ?","Perte production ~2t/arrêt","EF4444","FEE2E2"],
  ];
  qs.forEach(([ltr, q, ex, fg, bg], i) => {
    const col = i < 4 ? i : i - 4;
    const row = i < 4 ? 0 : 1;
    const w = i < 4 ? 2.22 : 2.95;
    const total = i < 4 ? 4 : 3;
    const startX = i < 4 ? 0.4 : 0.4 + col * 3.1;
    const x = i < 4 ? 0.4 + i * 2.37 : startX;
    const y = 1.35 + row * 1.45;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h:1.3, rectRadius:0.07, fill:{color:bg} });
    s.addShape(pres.shapes.OVAL, { x:x+0.08, y:y+0.08, w:0.4, h:0.4, fill:{color:fg} });
    s.addText(ltr, { x:x+0.08, y:y+0.08, w:0.4, h:0.4, fontSize:14, bold:true, color:C.white, align:"center", valign:"middle" });
    s.addText(q, { x:x+0.55, y:y+0.1, w:w-0.65, h:0.3, fontSize:12, bold:true, color:fg });
    sm(s, x+0.1, y+0.55, w-0.18, 0.6, ex, { fontSize:11 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.82, w:9.2, h:0.45, fill:{color:"EFF6FF"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.82, w:0.06, h:0.45, fill:{color:C.blue} });
  s.addText("Résultat : Description complète du problème en 30 minutes · Sans réunion supplémentaire", { x:0.56, y:4.87, w:9.0, h:0.35, fontSize:11, bold:true, color:"1E40AF", valign:"middle" });
}

// ─────────────────────────────────────────────
// SLIDE 11 — 5 POURQUOI
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 3 · OUTILS TERRAIN");
  slideTitle(s, "5 Pourquoi — Trouver la cause racine");
  sub(s, "Taiichi Ohno · Toyota 1950 · Adopté par Ford, Boeing, NASA");
  divider(s);

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.32, w:9.2, h:0.5, fill:{color:"FFFBEB"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.32, w:0.06, h:0.5, fill:{color:C.amber} });
  s.addText("Cas : Intervention prolongée non prévue — opérateur en zone risque 2h de trop", { x:0.56, y:1.38, w:9.0, h:0.38, fontSize:11, bold:true, color:C.amber, valign:"middle" });

  const whys = [
    ["1","Pourquoi l'intervention a-t-elle duré 2h de plus ?","L'intervention a duré 2h de plus que prévu"],
    ["2","Pourquoi 2h de plus ?","La pièce de rechange n'était pas disponible en stock"],
    ["3","Pourquoi la pièce n'était pas disponible ?","Pas de seuil d'alerte stock minimum défini"],
    ["4","Pourquoi pas de seuil d'alerte ?","Pas de responsable stock dédié pour pièces critiques"],
    ["5","CAUSE RACINE","Pas de politique de gestion des pièces de rechange critiques"],
  ];
  whys.forEach(([n, q, ans], i) => {
    const isLast = i === 4;
    const fg = isLast ? C.amber : "92400E";
    const bg = isLast ? "FEF3C7" : "FFFBEB";
    s.addShape(pres.shapes.OVAL, { x:0.4, y:2.0+i*0.57, w:0.32, h:0.32, fill:{color: isLast ? C.amber : "FEF3C7"}, line:{color:C.amber, width:isLast?0:1} });
    s.addText(n, { x:0.4, y:2.0+i*0.57, w:0.32, h:0.32, fontSize:10, bold:true, color: isLast ? C.white : C.amber, align:"center", valign:"middle" });
    if (i < 4) s.addText("↓", { x:0.51, y:2.32+i*0.57, w:0.1, h:0.22, fontSize:11, color:"FCA5A5", align:"center" });
    s.addShape(pres.shapes.RECTANGLE, { x:0.85, y:2.0+i*0.57, w:8.75, h:0.45, fill:{color:bg} });
    s.addText(q, { x:0.95, y:2.03+i*0.57, w:3.8, h:0.4, fontSize:10, bold:true, color:fg, valign:"middle" });
    s.addText("→  "+ans, { x:4.8, y:2.03+i*0.57, w:4.5, h:0.4, fontSize:10, color:C.gray, valign:"middle" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.92, w:9.2, h:0.38, fill:{color:"DCFCE7"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.92, w:0.06, h:0.38, fill:{color:C.green} });
  s.addText("Solution : Définir stock minimum pour pièces critiques → Implémentation : 1 jour · Impact : -50% temps d'attente inutile", { x:0.56, y:4.95, w:9.0, h:0.32, fontSize:10, bold:true, color:C.green, valign:"middle", wrap:true });
}

// ─────────────────────────────────────────────
// SLIDE 12 — SMED
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 3 · OUTILS TERRAIN");
  slideTitle(s, "SMED — Réduire les temps de changement");
  sub(s, "Single Minute Exchange of Die · Shigeo Shingo · Toyota 1950 · Résultat moyen : -40 à -75%");
  divider(s);

  // Left principle
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.35, w:5.3, h:3.1, fill:{color:C.lgray}, line:{color:"E5E7EB", width:0.5} });
  s.addText("LE PRINCIPE SMED", { x:0.5, y:1.42, w:5.1, h:0.25, fontSize:9, bold:true, color:C.dgray, charSpacing:2 });

  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.75, w:5.1, h:0.85, fill:{color:"FEE2E2"}, line:{color:"EF4444", width:1.5} });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.6, y:1.82, w:1.0, h:0.25, rectRadius:0.04, fill:{color:"EF4444"} });
  s.addText("INTERNE", { x:0.6, y:1.82, w:1.0, h:0.25, fontSize:8, bold:true, color:C.white, align:"center", valign:"middle" });
  s.addText("Machine ARRÊTÉE obligatoirement", { x:1.7, y:1.81, w:3.7, h:0.28, fontSize:11, bold:true, color:"991B1B" });
  s.addText("Démonter · Nettoyer · Installer", { x:1.7, y:2.1, w:3.7, h:0.42, fontSize:10, color:C.gray });

  s.addText("↓  Transformer en EXTERNE", { x:0.5, y:2.65, w:5.1, h:0.3, fontSize:11, bold:true, color:C.teal, align:"center" });

  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:3.0, w:5.1, h:0.85, fill:{color:"DCFCE7"}, line:{color:"22C55E", width:1.5} });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.6, y:3.07, w:1.0, h:0.25, rectRadius:0.04, fill:{color:"22C55E"} });
  s.addText("EXTERNE", { x:0.6, y:3.07, w:1.0, h:0.25, fontSize:8, bold:true, color:C.white, align:"center", valign:"middle" });
  s.addText("Machine EN MARCHE — préparer avant", { x:1.7, y:3.06, w:3.7, h:0.28, fontSize:11, bold:true, color:"166534" });
  s.addText("Préparer · Préchauffer · Outillage à portée", { x:1.7, y:3.35, w:3.7, h:0.42, fontSize:10, color:C.gray });

  // Right — result
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.9, y:1.35, w:3.7, h:1.6, rectRadius:0.08, fill:{color:"F0FDF4"}, line:{color:"22C55E", width:1} });
  s.addText("Cas concret", { x:6.0, y:1.42, w:3.5, h:0.25, fontSize:9, bold:true, color:C.dgray });
  s.addText("Changement de filtre", { x:6.0, y:1.68, w:3.5, h:0.28, fontSize:12, color:C.gray });
  s.addText("3h  →  1h20", { x:6.0, y:1.98, w:3.5, h:0.45, fontSize:22, bold:true, color:C.teal, align:"center" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:6.9, y:2.5, w:1.5, h:0.3, rectRadius:0.05, fill:{color:"DCFCE7"} });
  s.addText("-55%", { x:6.9, y:2.5, w:1.5, h:0.3, fontSize:11, bold:true, color:"166534", align:"center", valign:"middle" });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.9, y:3.1, w:3.7, h:1.35, rectRadius:0.08, fill:{color:"EFF6FF"}, line:{color:C.blue, width:1} });
  s.addText("Ce que vous gagnez :", { x:6.0, y:3.17, w:3.5, h:0.25, fontSize:9, bold:true, color:C.blue });
  const smedGains = ["Moins d'arrêt = moins d'attente équipe","Temps libéré = autres tâches VA","Moins d'exposition zone risque"];
  smedGains.forEach((g, i) => {
    s.addShape(pres.shapes.OVAL, { x:6.0, y:3.48+i*0.32, w:0.18, h:0.18, fill:{color:C.blue} });
    s.addText(g, { x:6.25, y:3.47+i*0.32, w:3.2, h:0.25, fontSize:10, color:C.blue });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.65, w:9.2, h:0.45, fill:{color:"FFFBEB"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.65, w:0.06, h:0.45, fill:{color:C.amber} });
  s.addText("Action cette semaine : Filmer une intervention · Lister les opérations · Reclasser interne/externe · Potentiel identifié en 1 jour", { x:0.56, y:4.69, w:9.0, h:0.37, fontSize:10, bold:true, color:C.amber, valign:"middle", wrap:true });
}

// ─────────────────────────────────────────────
// SLIDE 13 — BRAINSTORMING
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 3 · OUTILS TERRAIN");
  slideTitle(s, "Brainstorming — Générer des solutions ensemble");
  sub(s, "Alex Osborn 1941 · BBDO · Brain Writing : Bernd Rohrbach 1969 · 4 formats adaptés terrain");
  divider(s);

  const formats = [
    { icon:"👥", name:"CLASSIQUE", badge:"STANDARD", bg:"EDE9FE", bc:"5B21B6", fg:"5B21B6", when:"6–10 personnes · 20–40 min", how:"Groupe · Animateur · Secrétaire · Idées libres sans jugement", use:"Réunion équipe planifiée" },
    { icon:"🧍", name:"STAND-UP", badge:"URGENCE", bg:"DBEAFE", bc:"1D4ED8", fg:"1D4ED8", when:"5–8 personnes · 10–15 min", how:"Debout devant tableau blanc · Rapide · Focus sur 1 problème", use:"Début de shift ou problème du jour" },
    { icon:"✍️", name:"BRAIN WRITING", badge:"HIÉRARCHIE PRÉSENTE", bg:"DCFCE7", bc:"166534", fg:"166534", when:"6 personnes · 30 min", how:"Chaque personne écrit 3 idées · Passe la feuille · 6×3×5 = 100 idées", use:"Quand le chef est dans la salle" },
    { icon:"👤", name:"SOLO", badge:"ÉQUIPE INDISPONIBLE", bg:"FEF3C7", bc:"D97706", fg:"92400E", when:"1 personne · 15 min", how:"Seul · Lister TOUS les gaspillages vus dans la semaine · Pas de filtre", use:"Indisponibilité de toute l'équipe au même moment" },
  ];
  formats.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.7;
    const y = 1.35 + row * 1.65;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:1.5, fill:{color:f.bg}, line:{color:f.bc, width:1} });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:0.28, fill:{color:f.bc} });
    s.addText(f.name+" — "+f.badge, { x:x+0.06, y, w:4.4, h:0.28, fontSize:8, bold:true, color:C.white, valign:"middle" });
    s.addText(f.icon, { x:x+0.06, y:y+0.34, w:0.5, h:0.5, fontSize:22 });
    s.addText(f.when, { x:x+0.55, y:y+0.35, w:3.85, h:0.25, fontSize:10, bold:true, color:f.fg });
    s.addText(f.how, { x:x+0.55, y:y+0.6, w:3.85, h:0.38, fontSize:9, color:C.gray, wrap:true });
    s.addText("→ "+f.use, { x:x+0.06, y:y+1.15, w:4.3, h:0.28, fontSize:9, bold:true, color:f.fg, italic:true });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.82, w:9.2, h:0.45, fill:{color:"EFF6FF"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.82, w:0.06, h:0.45, fill:{color:C.blue} });
  s.addText("4 règles : Pas de jugement · Quantité > qualité · Toutes les idées comptent · Construire sur les idées des autres", { x:0.56, y:4.87, w:9.0, h:0.35, fontSize:10, bold:true, color:"1E40AF", valign:"middle" });
}

// ─────────────────────────────────────────────
// SLIDE 14 — CTQ + 5G + MANAGEMENT VISUEL
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PARTIE 3 · OUTILS 2ÈME GÉNÉRATION");
  slideTitle(s, "CTQ · 5G · Management Visuel");
  sub(s, "Outils avancés pour ancrer le Lean dans la réalité quotidienne");
  divider(s);

  // CTQ
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.35, w:2.9, h:3.3, fill:{color:"FFFBEB"}, line:{color:"F59E0B", width:1} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.35, w:2.9, h:0.28, fill:{color:"D97706"} });
  s.addText("CTQ — Critical To Quality", { x:0.46, y:1.35, w:2.8, h:0.28, fontSize:9, bold:true, color:C.white, valign:"middle" });
  s.addText("Bill Smith · Motorola 1986", { x:0.5, y:1.7, w:2.75, h:0.22, fontSize:9, color:C.dgray, italic:true });
  s.addText("Transformer les attentes floues en exigences mesurables.", { x:0.5, y:1.95, w:2.75, h:0.45, fontSize:10, color:C.gray, wrap:true });
  s.addText('Client dit :\n"Qualité bonne"', { x:0.5, y:2.48, w:2.75, h:0.5, fontSize:10, color:C.gray, italic:true });
  s.addText("↓  CTQ →", { x:0.5, y:3.0, w:2.75, h:0.25, fontSize:10, bold:true, color:C.amber, align:"center" });
  s.addText("Spec précise · Délai ±4h\nTaux incident = 0\nExpo ≤ 30min/zone", { x:0.5, y:3.28, w:2.75, h:0.75, fontSize:10, bold:true, color:C.amber, wrap:true });

  // 5G
  s.addShape(pres.shapes.RECTANGLE, { x:3.5, y:1.35, w:3.0, h:3.3, fill:{color:"F0FDF4"}, line:{color:"22C55E", width:1} });
  s.addShape(pres.shapes.RECTANGLE, { x:3.5, y:1.35, w:3.0, h:0.28, fill:{color:"166534"} });
  s.addText("5G — Les 5 Gemba · Toyota", { x:3.56, y:1.35, w:2.9, h:0.28, fontSize:9, bold:true, color:C.white, valign:"middle" });
  const gs = [["Gemba","Aller sur le terrain"],["Gembutsu","Observer l'objet réel"],["Genjitsu","Faits réels · données"],["Genri","Théorie · principes"],["Genba Walk","15 min/jour · 3 gaspillages"]];
  gs.forEach(([g, d], i) => {
    s.addShape(pres.shapes.OVAL, { x:3.55, y:1.74+i*0.47, w:0.22, h:0.22, fill:{color:i<4?"0B7A6A":"22C55E"} });
    s.addText(g, { x:3.84, y:1.72+i*0.47, w:2.5, h:0.24, fontSize:10, bold:true, color:"166534" });
    s.addText(d, { x:3.84, y:1.93+i*0.47, w:2.5, h:0.22, fontSize:9, color:C.dgray });
  });

  // Management Visuel
  s.addShape(pres.shapes.RECTANGLE, { x:6.7, y:1.35, w:2.9, h:3.3, fill:{color:"EFF6FF"}, line:{color:"3B82F6", width:1} });
  s.addShape(pres.shapes.RECTANGLE, { x:6.7, y:1.35, w:2.9, h:0.28, fill:{color:"1D4ED8"} });
  s.addText("Management Visuel · Andon 1960", { x:6.76, y:1.35, w:2.8, h:0.28, fontSize:9, bold:true, color:C.white, valign:"middle" });
  s.addText("Rendre l'info visible, immédiate, compréhensible par tous.", { x:6.76, y:1.7, w:2.75, h:0.45, fontSize:10, color:C.gray, wrap:true });
  const andon = [["22C55E","VERT","Objectif atteint · OK"],["F59E0B","ORANGE","Surveillance · Attention"],["EF4444","ROUGE","Problème · Agir maintenant"]];
  andon.forEach(([col, label, desc], i) => {
    s.addShape(pres.shapes.OVAL, { x:6.76, y:2.28+i*0.7, w:0.38, h:0.38, fill:{color:col} });
    s.addText(label, { x:7.22, y:2.28+i*0.7, w:2.2, h:0.22, fontSize:10, bold:true, color:col });
    s.addText(desc, { x:7.22, y:2.5+i*0.7, w:2.2, h:0.22, fontSize:9, color:C.dgray });
  });
  s.addText("Tableau de bord visible par tous → Déclenche les bons comportements sans réunion.", { x:6.76, y:4.08, w:2.75, h:0.45, fontSize:9, bold:true, color:"1D4ED8", wrap:true });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.82, w:9.2, h:0.45, fill:{color:"F0FDF4"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.82, w:0.06, h:0.45, fill:{color:C.teal} });
  s.addText("5G + Management Visuel = identifier 150 gaspillages/an · ROI infini · 0 investissement", { x:0.56, y:4.87, w:9.0, h:0.35, fontSize:11, bold:true, color:C.teal, valign:"middle" });
}

// ─────────────────────────────────────────────
// SLIDE 15 — MESSAGE AUX RESPONSABLES
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "VOTRE RÔLE DE RESPONSABLE");
  slideTitle(s, "Ce qu'on attend de vous — concretement");
  sub(s, "Votre rôle : exiger, accompagner et standardiser — pas tout faire seul");
  divider(s);

  // Ce que vous faites
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.35, w:4.5, h:3.1, fill:{color:"EFF6FF"}, line:{color:C.blue, width:1.5} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.35, w:4.5, h:0.35, fill:{color:"1D4ED8"} });
  s.addText("CE QUE VOUS FAITES — chaque semaine", { x:0.5, y:1.35, w:4.3, h:0.35, fontSize:9, bold:true, color:C.white, valign:"middle" });
  const doList = [
    "Gemba Walk 15 min/jour → noter 3 gaspillages",
    "Exiger QQOQCCP + 5 Pourquoi pour chaque problème",
    "Valider 1 amélioration KAIZEN par semaine par équipe",
    "Afficher le tableau de bord (Mgmt Visuel) chaque lundi",
    "Appliquer PDCA sur chaque problème récurrent",
    "Refuser d'accepter 'c'est comme ça' comme réponse",
  ];
  doList.forEach((t, i) => {
    s.addShape(pres.shapes.OVAL, { x:0.55, y:1.85+i*0.38, w:0.22, h:0.22, fill:{color:"1D4ED8"} });
    s.addText(String(i+1), { x:0.55, y:1.85+i*0.38, w:0.22, h:0.22, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle" });
    sm(s, 0.85, 1.83+i*0.38, 3.9, 0.34, t, { fontSize:10 });
  });

  // Ce que vous exigez
  s.addShape(pres.shapes.RECTANGLE, { x:5.1, y:1.35, w:4.5, h:3.1, fill:{color:C.lgray}, line:{color:"E5E7EB", width:0.5} });
  s.addShape(pres.shapes.RECTANGLE, { x:5.1, y:1.35, w:4.5, h:0.35, fill:{color:C.gray} });
  s.addText("CE QUE VOUS EXIGEZ DE VOS ÉQUIPES", { x:5.2, y:1.35, w:4.3, h:0.35, fontSize:9, bold:true, color:C.white, valign:"middle" });

  const exigences = [["✅","Identifier 3 gaspillages/jour (MUDA)","lgreen","166534"],["✅","Proposer UNE solution par problème","lgreen","166534"],["✅","Suivre les standards définis","lgreen","166534"],["❌","Remettre à plus tard sans raison","lred","991B1B"],["❌","Accepter un problème récurrent sans PDCA","lred","991B1B"]];
  exigences.forEach(([ico, txt, bgK, fgK], i) => {
    const bg = bgK === "lgreen" ? "F0FDF4" : "FEF2F2";
    const fg = bgK === "lgreen" ? "166534" : "991B1B";
    s.addShape(pres.shapes.RECTANGLE, { x:5.2, y:1.85+i*0.45, w:4.25, h:0.38, fill:{color:bg} });
    s.addText(ico+" "+txt, { x:5.3, y:1.88+i*0.45, w:4.1, h:0.3, fontSize:10, bold: bgK==="lgreen", color:fg, valign:"middle" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.62, w:9.2, h:0.75, fill:{color:"FFFBEB"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.62, w:0.06, h:0.75, fill:{color:C.amber} });
  s.addText("Ce que vous gagnez : Votre équipe résout ses propres problèmes · Vous passez de pompier à leader · Votre indicateur d'arrêts non planifiés baisse de 40% en 6 mois · Vous avez enfin des données pour défendre vos demandes.", { x:0.56, y:4.66, w:9.0, h:0.67, fontSize:10, bold:true, color:C.amber, valign:"middle", wrap:true });
}

// ─────────────────────────────────────────────
// SLIDE 16 — PLAN CONCRET 15 GASPILLAGES/SEMAINE
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "PLAN D'ACTION TERRAIN");
  slideTitle(s, "Comment éliminer 15 gaspillages par semaine");
  sub(s, "Plan concret · Outils utilisés · Gains réels à chaque étape");
  divider(s);

  s.addText("LES 5 ÉTAPES — DÈS LUNDI MATIN", { x:0.4, y:1.3, w:9.2, h:0.22, fontSize:8, bold:true, color:C.dgray, charSpacing:2 });

  const steps = [
    { day:"Lun", tool:"5G · Gemba Walk", action:"15 min sur le terrain · noter 3 gaspillages observés", gain:"3 MUDA identifiés", gainBg:"DCFCE7", gainFg:"166534" },
    { day:"Mar", tool:"QQOQCCP + 5 Pourquoi", action:"30 min · Prendre le gaspillage le plus critique · identifier cause racine", gain:"1 cause racine résolue", gainBg:"DBEAFE", gainFg:"1E40AF" },
    { day:"Mer", tool:"Brainstorming Stand-up", action:"10 min debout · équipe propose solutions · choisir la plus simple", gain:"1 solution testée", gainBg:"EDE9FE", gainFg:"5B21B6" },
    { day:"Jeu", tool:"PDCA — Phase D + Do", action:"Tester la solution · mesurer le résultat avant/après", gain:"Résultat mesuré", gainBg:"FEF3C7", gainFg:"92400E" },
    { day:"Ven", tool:"Mgmt Visuel + Act", action:"Mettre à jour le tableau · standardiser si succès · partager équipe", gain:"Standard établi", gainBg:"DCFCE7", gainFg:"166534" },
  ];
  steps.forEach((st, i) => {
    const x = 0.4 + i * 1.88;
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.58, w:1.75, h:2.8, fill:{color:C.white}, line:{color:"E5E7EB", width:0.8} });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.58, w:1.75, h:0.38, fill:{color:C.navy} });
    s.addText(st.day, { x, y:1.58, w:1.75, h:0.38, fontSize:13, bold:true, color:C.white, align:"center", valign:"middle" });
    s.addText(st.tool, { x:x+0.06, y:2.02, w:1.63, h:0.4, fontSize:9, bold:true, color:C.black, wrap:true });
    sm(s, x+0.06, 2.45, 1.63, 1.15, st.action, { fontSize:9 });
    s.addShape(pres.shapes.RECTANGLE, { x, y:3.75, w:1.75, h:0.55, fill:{color:st.gainBg} });
    s.addText(st.gain, { x:x+0.04, y:3.78, w:1.67, h:0.48, fontSize:9, bold:true, color:st.gainFg, align:"center", valign:"middle" });
  });

  // Cumul
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.45, w:9.2, h:0.45, fill:{color:C.navy} });
  s.addText("📊  1 semaine = 15 gaspillages × 50 semaines = 750 MUDA/an éliminés → équivalent de 3 à 5 semaines de travail récupérées", { x:0.5, y:4.49, w:9.0, h:0.37, fontSize:10, bold:true, color:C.white, valign:"middle", wrap:true });

  // Gains par rôle
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.02, w:4.5, h:0.45, fill:{color:"F0FDF4"} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.02, w:0.06, h:0.45, fill:{color:C.teal} });
  s.addText("Responsable : 40% moins de problèmes récurrents · Équipe autonome", { x:0.56, y:5.06, w:4.2, h:0.37, fontSize:9, bold:true, color:C.teal, valign:"middle" });
  s.addShape(pres.shapes.RECTANGLE, { x:5.1, y:5.02, w:4.5, h:0.45, fill:{color:"EFF6FF"} });
  s.addShape(pres.shapes.RECTANGLE, { x:5.1, y:5.02, w:0.06, h:0.45, fill:{color:C.blue} });
  s.addText("Opérateur : Moins de retouches · Travail mieux organisé · Moins de stress", { x:5.26, y:5.06, w:4.2, h:0.37, fontSize:9, bold:true, color:C.blue, valign:"middle" });
}

// ─────────────────────────────────────────────
// SLIDE 17 — CE QUE VOUS GAGNEZ RÉELLEMENT
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  sectionTag(s, "CONCLUSION");
  slideTitle(s, "Ce que vous gagnez réellement");
  sub(s, "Résultats mesurés sur des équipes industrielles similaires");
  divider(s);

  const gains = [
    ["-40%","Problèmes récurrents","En 6 mois avec PDCA + 5 Pourquoi systématique","3B82F6"],
    ["+2h/j","Valeur ajoutée","Par personne · Sans recruter · Juste en supprimant NVA","22C55E"],
    ["-55%","Temps de changement","Avec SMED · Filmer une intervention = potentiel identifié en 1 jour","0B7A6A"],
    ["0 Dh","D'investissement","Pour commencer · Juste votre regard et 15 min/jour","F59E0B"],
  ];
  gains.forEach(([num, lbl, desc, col], i) => {
    const col2 = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col2 * 4.7;
    const y = 1.32 + row * 1.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:4.5, h:1.35, rectRadius:0.08, fill:{color:C.lgray} });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.06, h:1.35, fill:{color:col} });
    s.addText(num, { x:x+0.2, y:y+0.1, w:1.5, h:0.7, fontSize:32, bold:true, color:col });
    s.addText(lbl, { x:x+0.2, y:y+0.75, w:1.8, h:0.3, fontSize:11, bold:true, color:C.black });
    sm(s, x+1.9, y+0.12, 2.5, 1.1, desc, { fontSize:10 });
  });

  // Final CTA
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.4, y:4.55, w:9.2, h:0.75, rectRadius:0.08, fill:{color:C.navy} });
  s.addText('"Commencez petit, commencez maintenant, capitalisez toujours."', { x:0.5, y:4.6, w:9.0, h:0.38, fontSize:13, bold:true, color:"93C5FD", align:"center" });
  s.addText("Votre équipe vous attend. Les outils sont prêts. La première amélioration, c'est cette semaine.", { x:0.5, y:4.95, w:9.0, h:0.28, fontSize:10, color:"BFDBFE", align:"center" });
}

// ─────────────────────────────────────────────
// WRITE FILE
// ─────────────────────────────────────────────
pres.writeFile({ fileName: "/home/claude/Formation_Lean_Management_KAIZEN.pptx" })
  .then(() => console.log("✅ PPTX created"))
  .catch(e => console.error("❌", e));
