const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign
} = require('docx');
const fs = require('fs');

const BLUE = "1F4E79";
const LIGHT_BLUE = "D6E4F0";
const ACCENT = "2E75B6";
const LIGHT_YELLOW = "FFF9E6";
const LIGHT_GREEN = "E8F5E9";
const LIGHT_RED = "FDECEA";
const LIGHT_GREY = "F5F5F5";
const BORDER_COLOR = "BBBBBB";

const border = { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, bold: true, size: 32, color: BLUE, font: "Arial" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 140 },
    children: [new TextRun({ text, bold: true, size: 26, color: ACCENT, font: "Arial" })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, bold: true, size: 24, color: "333333", font: "Arial" })]
  });
}

function body(text, options = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, font: "Arial", ...options })]
  });
}

function italic(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, font: "Arial", italics: true, color: "555555" })]
  });
}

function blank(size = 80) {
  return new Paragraph({
    spacing: { before: size, after: 0 },
    children: [new TextRun("")]
  });
}

function bullet(text, indent = 720) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function numbered(text) {
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function colorBox(paragraphs, fillColor) {
  return new Table({
    width: { size: 9200, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: 9200, type: WidthType.DXA },
        shading: { fill: fillColor, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 200, right: 200 },
        children: paragraphs
      })]
    })]
  });
}

function twoCol(left, right, leftWidth = 4500, rightWidth = 4700) {
  return new Table({
    width: { size: 9200, type: WidthType.DXA },
    columnWidths: [leftWidth, rightWidth],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders,
          width: { size: leftWidth, type: WidthType.DXA },
          shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 160, right: 160 },
          children: left
        }),
        new TableCell({
          borders,
          width: { size: rightWidth, type: WidthType.DXA },
          shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 160, right: 160 },
          children: right
        })
      ]
    })]
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } },
    children: [new TextRun("")]
  });
}

function exerciseBox(title, instructions, paragraphs, fillColor = LIGHT_YELLOW) {
  return new Table({
    width: { size: 9200, type: WidthType.DXA },
    rows: [
      new TableRow({
        children: [new TableCell({
          borders,
          width: { size: 9200, type: WidthType.DXA },
          shading: { fill: ACCENT, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 200, right: 200 },
          children: [new Paragraph({ children: [new TextRun({ text: title, bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })]
        })]
      }),
      new TableRow({
        children: [new TableCell({
          borders,
          width: { size: 9200, type: WidthType.DXA },
          shading: { fill: fillColor, type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 200, right: 200 },
          children: [
            new Paragraph({ spacing: { before: 40, after: 80 }, children: [new TextRun({ text: instructions, size: 21, font: "Arial", italics: true, color: "444444" })] }),
            ...paragraphs
          ]
        })]
      })
    ]
  });
}

function writeLine(label = "") {
  return new Paragraph({
    spacing: { before: 60, after: 0 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA", space: 1 } },
    children: [new TextRun({ text: label, size: 21, font: "Arial", color: "888888" })]
  });
}

// ─── DOCUMENT ────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } } }]
      },
      {
        reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } } }]
      },
      {
        reference: "alpha",
        levels: [{ level: 0, format: LevelFormat.LOWER_LETTER, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } } }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: BLUE },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: ACCENT },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [

      // ── TITRE ──────────────────────────────────────────────────────────────
      new Table({
        width: { size: 9740, type: WidthType.DXA },
        rows: [new TableRow({
          children: [new TableCell({
            borders,
            width: { size: 9740, type: WidthType.DXA },
            shading: { fill: BLUE, type: ShadingType.CLEAR },
            margins: { top: 200, bottom: 200, left: 300, right: 300 },
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DELF B2 — Production Orale", bold: true, size: 40, color: "FFFFFF", font: "Arial" })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Phiếu bài tập : Structures, Connecteurs & Problématique", size: 24, color: "CCE0F5", font: "Arial" })] })
            ]
          })]
        })]
      }),

      blank(160),

      // ════════════════════════════════════════════════════════════════════════
      // PARTIE 1 — PROBLÉMATIQUE
      // ════════════════════════════════════════════════════════════════════════
      h1("PARTIE 1 — Construire une Problématique"),
      divider(),

      colorBox([
        new Paragraph({ spacing: { before: 60, after: 80 }, children: [new TextRun({ text: "Pourquoi une bonne problématique ?", bold: true, size: 22, font: "Arial", color: BLUE })] }),
        body("Une problématique réussie au DELF B2 est une question qui révèle le paradoxe ou la tension centrale du sujet. Elle structure toute votre prise de parole."),
        blank(60),
        body("Deux types principaux :", { bold: true }),
        bullet("Sujets paradoxaux (tension entre deux réalités) → structure en opposition sóng đôi"),
        bullet("Sujets clairement négatifs (discrimination, danger, exploitation) → structure en « comment lutter contre »"),
      ], LIGHT_BLUE),

      blank(120),
      h2("1.1 — Structure « Paradoxe » (sóng đôi)"),
      body("Formule : Alors que / Bien que [Réalité A], [Réalité B] — Comment / Dans quelle mesure [question centrale] ?", { bold: true }),
      blank(80),

      new Table({
        width: { size: 9740, type: WidthType.DXA },
        columnWidths: [3200, 6540],
        rows: [
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3200, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Sujet", bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 6540, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Exemple de problématique", bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3200, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Égalité salariale", size: 22, font: "Arial", bold: true })] })] }),
              new TableCell({ borders, width: { size: 6540, type: WidthType.DXA }, shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Alors que des lois garantissent l'égalité de rémunération, un écart salarial persistant pénalise encore les femmes — dans quelle mesure les entreprises doivent-elles s'adapter pour rendre cette égalité effective ?", size: 21, font: "Arial", italics: true })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3200, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Télétravail", size: 22, font: "Arial", bold: true })] })] }),
              new TableCell({ borders, width: { size: 6540, type: WidthType.DXA }, shading: { fill: LIGHT_GREY, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Si le télétravail libère les salariés des contraintes des déplacements, il risque à la fois de les isoler et de rendre leurs frontières professionnelles invisibles — comment trouver le juste équilibre entre flexibilité et cohésion d'équipe ?", size: 21, font: "Arial", italics: true })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3200, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Semaine de 4 jours", size: 22, font: "Arial", bold: true })] })] }),
              new TableCell({ borders, width: { size: 6540, type: WidthType.DXA }, shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Bien que la semaine de quatre jours soit présentée comme un gain de productivité et de bien-être, sa généralisation soulève la question de l'équité entre secteurs — dans quelle mesure est-elle applicable à l'ensemble du monde du travail ?", size: 21, font: "Arial", italics: true })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3200, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Réseaux sociaux / Écrans", size: 22, font: "Arial", bold: true })] })] }),
              new TableCell({ borders, width: { size: 6540, type: WidthType.DXA }, shading: { fill: LIGHT_GREY, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Alors que les écrans et les réseaux sociaux sont devenus des outils de communication incontournables, leur usage excessif menace le développement cognitif et social — comment préserver leurs bénéfices tout en en limitant les effets néfastes ?", size: 21, font: "Arial", italics: true })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3200, type: WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Vivre plus longtemps", size: 22, font: "Arial", bold: true })] })] }),
              new TableCell({ borders, width: { size: 6540, type: WidthType.DXA }, shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Si l'espérance de vie augmente grâce aux progrès médicaux, vivre plus longtemps ne signifie pas toujours vivre mieux — dans quelle mesure nos sociétés sont-elles prêtes à garantir l'autonomie et la dignité des personnes âgées ?", size: 21, font: "Arial", italics: true })] })] })
            ]
          }),
        ]
      }),

      blank(120),
      h2("1.2 — Structure « Comment lutter contre » (sujet négatif)"),
      body("Formule : Face à [problème avéré], comment [société / entreprises / individus] peuvent-ils [lutter / agir / limiter] ?", { bold: true }),
      blank(80),

      new Table({
        width: { size: 9740, type: WidthType.DXA },
        columnWidths: [3200, 6540],
        rows: [
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3200, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Sujet", bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 6540, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Exemple de problématique", bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3200, type: WidthType.DXA }, shading: { fill: LIGHT_RED, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Instituts de beauté pour enfants", size: 22, font: "Arial", bold: true })] })] }),
              new TableCell({ borders, width: { size: 6540, type: WidthType.DXA }, shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Face à la marchandisation croissante de l'enfance et à la normalisation précoce des codes esthétiques adultes, comment protéger les enfants d'une culture de l'apparence qui fragilise leur développement psychologique ?", size: 21, font: "Arial", italics: true })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3200, type: WidthType.DXA }, shading: { fill: LIGHT_RED, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Obsolescence programmée", size: 22, font: "Arial", bold: true })] })] }),
              new TableCell({ borders, width: { size: 6540, type: WidthType.DXA }, shading: { fill: LIGHT_GREY, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Face à des pratiques industrielles qui raccourcissent délibérément la durée de vie des produits au détriment des consommateurs et de l'environnement, comment légiférer efficacement pour protéger les citoyens et accélérer la transition vers une économie circulaire ?", size: 21, font: "Arial", italics: true })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3200, type: WidthType.DXA }, shading: { fill: LIGHT_RED, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Fake news", size: 22, font: "Arial", bold: true })] })] }),
              new TableCell({ borders, width: { size: 6540, type: WidthType.DXA }, shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Face à la prolifération des fausses informations sur les réseaux sociaux qui menacent la démocratie, comment développer l'esprit critique des citoyens tout en évitant une censure incompatible avec la liberté d'expression ?", size: 21, font: "Arial", italics: true })] })] })
            ]
          }),
        ]
      }),

      blank(120),

      exerciseBox(
        "✏️  EXERCICE 1 — Rédigez une problématique",
        "Pour chaque sujet ci-dessous, construisez une problématique en choisissant le bon type (paradoxe ou comment lutter contre). Utilisez les structures modèles.",
        [
          blank(40),
          body("1.  La sieste au travail", { bold: true }),
          writeLine(), writeLine(), blank(60),
          body("2.  Le droit à l'oubli sur Internet", { bold: true }),
          writeLine(), writeLine(), blank(60),
          body("3.  La chirurgie esthétique", { bold: true }),
          writeLine(), writeLine(), blank(60),
          body("4.  Le redoublement à l'école", { bold: true }),
          writeLine(), writeLine(), blank(60),
        ],
        LIGHT_YELLOW
      ),

      blank(180),

      // ════════════════════════════════════════════════════════════════════════
      // PARTIE 2 — PHRASES D'AMORCE
      // ════════════════════════════════════════════════════════════════════════
      h1("PARTIE 2 — Phrases d'amorce & Introduction"),
      divider(),

      body("L'amorce contextualise le sujet AVANT d'annoncer votre problématique. Elle doit être factuelle, précise et naturelle. Évitez « Ce sujet est très intéressant... »"),

      blank(80),
      h2("2.1 — Types d'amorces"),

      new Table({
        width: { size: 9740, type: WidthType.DXA },
        columnWidths: [2800, 6940],
        rows: [
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2800, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Type d'amorce", bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 6940, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Formule / Exemple", bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })] })
            ]
          }),
          ...[
            ["Chiffre / statistique", "D'après une étude récente, [X%] des [population] ... Ce constat soulève la question de..."],
            ["Phénomène de société", "Depuis plusieurs années, on observe une tendance croissante à [phénomène]. Dans ce contexte, ..."],
            ["Paradoxe constaté", "Alors que [Réalité A], force est de constater que [Réalité B]. Cette contradiction nous amène à nous demander..."],
            ["Citation / Formule", "Comme le souligne [expert / rapport], « [citation courte] ». Ce constat illustre bien..."],
            ["Fait d'actualité", "Le document qui nous est proposé aborde [thème], une question au cœur du débat public depuis..."],
          ].map(([type, ex], i) => new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2800, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? LIGHT_BLUE : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: type, size: 21, font: "Arial", bold: true })] })] }),
              new TableCell({ borders, width: { size: 6940, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? LIGHT_GREY : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: ex, size: 21, font: "Arial", italics: true })] })] })
            ]
          }))
        ]
      }),

      blank(120),
      h2("2.2 — Structure complète d'une introduction (modèle)"),

      colorBox([
        body("① AMORCE", { bold: true, color: BLUE }),
        italic("Le document proposé aborde la question de l'égalité salariale, un sujet au cœur du débat public. Pour 88 % des femmes françaises, concilier vie professionnelle et maternité reste une difficulté majeure."),
        blank(60),
        body("② PRÉSENTATION DU DOCUMENT (si applicable)", { bold: true, color: BLUE }),
        italic("D'un côté, les sociétés encouragent la natalité ; de l'autre, les femmes sont encore pénalisées dans leur carrière lorsqu'elles ont des enfants."),
        blank(60),
        body("③ PROBLÉMATIQUE", { bold: true, color: BLUE }),
        italic("Ce paradoxe m'amène à me poser la question suivante : dans quelle mesure les entreprises doivent-elles s'adapter pour rendre l'égalité professionnelle réellement effective ?"),
        blank(60),
        body("④ ANNONCE DU PLAN", { bold: true, color: BLUE }),
        italic("Pour répondre à cette question, je défendrai une position favorable à un meilleur équilibre vie pro/vie familiale. Je présenterai d'abord les raisons qui soutiennent cette position, avant d'envisager les mesures concrètes que les entreprises pourraient adopter."),
      ], LIGHT_YELLOW),

      blank(120),

      exerciseBox(
        "✏️  EXERCICE 2 — Construire une introduction complète",
        "Choisissez l'un des sujets suivants et rédigez une introduction complète (amorce + présentation + problématique + annonce du plan) en vous appuyant sur les structures ci-dessus.",
        [
          blank(40),
          body("Sujets au choix : La semaine de 4 jours  /  Les seniors dans les start-ups  /  Le droit à la déconnexion  /  Manger bio et budget"),
          blank(80),
          writeLine("Amorce :"),
          writeLine(), blank(40),
          writeLine("Présentation :"),
          writeLine(), blank(40),
          writeLine("Problématique :"),
          writeLine(), blank(40),
          writeLine("Annonce du plan :"),
          writeLine(), blank(40),
        ],
        LIGHT_GREEN
      ),

      blank(180),

      // ════════════════════════════════════════════════════════════════════════
      // PARTIE 3 — CONNECTEURS
      // ════════════════════════════════════════════════════════════════════════
      h1("PARTIE 3 — Connecteurs pour l'oral"),
      divider(),

      body("Les connecteurs structurent votre discours et vous évitent les répétitions. À l'oral, ils doivent sonner naturels — pas scolaires."),

      blank(100),
      h2("3.1 — Tableau des connecteurs par fonction"),

      new Table({
        width: { size: 9740, type: WidthType.DXA },
        columnWidths: [2400, 3700, 3640],
        rows: [
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2400, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Fonction", bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 3700, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Connecteurs courants", bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 3640, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Variantes stylistiques (oral B2)", bold: true, size: 22, color: "FFFFFF", font: "Arial" })] })] })
            ]
          }),
          ...[
            ["Introduire", "Tout d'abord / En premier lieu / Pour commencer", "Il convient de souligner que... / Il est important de noter que..."],
            ["Ajouter", "De plus / En outre / Par ailleurs / Ensuite", "À cela s'ajoute le fait que... / On peut également mentionner..."],
            ["Illustrer", "Par exemple / Ainsi / C'est le cas de", "Pour illustrer ce point... / Prenons l'exemple de..."],
            ["Concéder (CONTRE)", "Certes / Il est vrai que / On peut admettre que", "Je reconnais volontiers que... / Il serait malhonnête de nier que..."],
            ["Réfuter / Nuancer", "Cependant / Néanmoins / Toutefois / Mais", "Or, force est de constater que... / Il n'en demeure pas moins que..."],
            ["Causer", "En effet / Car / Étant donné que / Puisque", "La raison en est que... / Ce phénomène s'explique par..."],
            ["Conséquence", "Donc / Ainsi / Par conséquent / C'est pourquoi", "Il en résulte que... / Cela implique nécessairement que..."],
            ["Opposer", "Alors que / Tandis que / En revanche / Au contraire", "D'un côté... de l'autre... / Si X est vrai, Y l'est tout autant."],
            ["Conclure", "En conclusion / Pour conclure / En définitive", "En guise de conclusion... / Pour résumer ma position..."],
          ].map(([f, c, v], i) => new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2400, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? LIGHT_BLUE : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: f, size: 21, font: "Arial", bold: true })] })] }),
              new TableCell({ borders, width: { size: 3700, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? LIGHT_GREY : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: c, size: 21, font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 3640, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? LIGHT_YELLOW : "FAFAFA", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: v, size: 21, font: "Arial", italics: true })] })] }),
            ]
          }))
        ]
      }),

      blank(120),
      h2("3.2 — Formules pour la concession (CONTRE → mais POUR)"),
      colorBox([
        body("La concession est la marque d'un B2 solide. Elle montre que vous nuancez, pas que vous capitullez."),
        blank(60),
        bullet("Certes, [argument contre], mais il n'en reste pas moins que [argument pour]."),
        bullet("On peut certes reconnaître que [contre], cependant [pour] l'emporte largement."),
        bullet("Il est vrai que [contre]. Néanmoins, cet argument ne saurait effacer le fait que [pour]."),
        bullet("Je reconnais volontiers que [contre]. Or, à y regarder de plus près, [pour]."),
        bullet("Si [contre] est une réalité qu'il serait malhonnête de nier, [pour] me semble bien plus déterminant."),
      ], LIGHT_GREEN),

      blank(120),

      exerciseBox(
        "✏️  EXERCICE 3 — Compléter avec le bon connecteur",
        "Choisissez dans la liste : [Certes / Par conséquent / En outre / Toutefois / Il convient de souligner que / Or / C'est pourquoi / En revanche]",
        [
          blank(40),
          body("1. Le télétravail permet d'économiser du temps. ________, certains salariés se sentent isolés."),
          writeLine(), blank(60),
          body("2. ________ les femmes sont encore sous-représentées aux postes de direction, la situation évolue lentement."),
          writeLine(), blank(60),
          body("3. La semaine de 4 jours améliore le bien-être. ________, elle réduit l'empreinte carbone."),
          writeLine(), blank(60),
          body("4. ________ l'e-sport exige un entraînement intensif, il n'en constitue pas moins un sport à part entière."),
          writeLine(), blank(60),
          body("5. Les écrans perturbent le sommeil des enfants. ________, les médecins recommandent d'en limiter l'usage."),
          writeLine(), blank(60),
          body("6. Le document traite de l'égalité salariale. ________, malgré les lois, un écart de 16 % persiste en France."),
          writeLine(), blank(60),
        ],
        LIGHT_YELLOW
      ),

      blank(180),

      // ════════════════════════════════════════════════════════════════════════
      // PARTIE 4 — STRUCTURES CLÉS
      // ════════════════════════════════════════════════════════════════════════
      h1("PARTIE 4 — Structures & Formules à mémoriser"),
      divider(),

      h2("4.1 — Présenter et défendre une position"),

      colorBox([
        body("Annoncer sa position :", { bold: true }),
        bullet("Pour ma part, je défendrai une position [favorable à / critique envers] ..."),
        bullet("Je me montrerai plutôt [en faveur de / réservé quant à] ..."),
        bullet("Ma position est la suivante : [reformulation claire]."),
        blank(60),
        body("Présenter un argument :", { bold: true }),
        bullet("Il convient de souligner que [argument] — en effet, [explication + exemple]."),
        bullet("Le premier argument en faveur de [position] réside dans [fait]. Ainsi, [illustration]."),
        bullet("D'un point de vue [économique / social / environnemental], [argument]."),
        blank(60),
        body("Appuyer par un exemple :", { bold: true }),
        bullet("Pour illustrer ce point, prenons l'exemple de [pays / entreprise / étude] qui [résultat concret]."),
        bullet("C'est d'ailleurs ce que révèle [étude / statistique] : [chiffre ou fait]."),
        bullet("On peut citer le cas de [X], qui [fait concret], preuve que [argument]."),
      ], LIGHT_BLUE),

      blank(120),
      h2("4.2 — Structures de nuance et de reformulation"),

      colorBox([
        body("Nuancer sans céder :", { bold: true }),
        bullet("Si l'on peut comprendre [contre], cela ne remet pas fondamentalement en cause [pour]."),
        bullet("Cette objection mérite d'être prise au sérieux, mais elle ne constitue pas une raison suffisante pour..."),
        blank(60),
        body("Reformuler avec élégance :", { bold: true }),
        bullet("Autrement dit, [reformulation plus claire de l'argument]."),
        bullet("En d'autres termes, il s'agit de [reformulation]."),
        bullet("Ce qui revient à dire que [reformulation]."),
        blank(60),
        body("Exprimer la causalité :", { bold: true }),
        bullet("Ce phénomène s'explique en grande partie par [cause]."),
        bullet("La raison principale en est que [cause] — ce qui conduit inévitablement à [conséquence]."),
        bullet("L'origine de ce problème tient à [cause structurelle]."),
      ], LIGHT_GREEN),

      blank(120),
      h2("4.3 — Structures pour la conclusion"),

      colorBox([
        bullet("En définitive, il ressort de cette réflexion que [synthèse de la position]."),
        bullet("Pour conclure, je maintiens que [position], même si [nuance / limite reconnue]."),
        bullet("En guise de conclusion, on peut affirmer que [synthèse]. Reste ouverte la question de [ouverture]."),
        bullet("Cette question nous invite, plus largement, à nous demander [ouverture vers un sujet plus vaste]."),
        bullet("Et vous, partagez-vous ce constat ? / Qu'en pensez-vous de votre côté ?"),
      ], LIGHT_YELLOW),

      blank(120),

      exerciseBox(
        "✏️  EXERCICE 4 — Développer un argument complet",
        "En vous appuyant sur les structures ci-dessus, développez UN argument POUR sur le sujet suivant, en incluant : position + argument + exemple + nuance (concession).",
        [
          blank(40),
          body("Sujet : La suppression des notes à l'école / L'école inversée / Le coaching en ligne", { bold: true }),
          blank(100),
          writeLine("Position :"),
          writeLine(), blank(40),
          writeLine("Argument principal :"),
          writeLine(), writeLine(), blank(40),
          writeLine("Exemple / illustration :"),
          writeLine(), writeLine(), blank(40),
          writeLine("Concession (certes... mais...) :"),
          writeLine(), writeLine(), blank(40),
        ],
        LIGHT_GREEN
      ),

      blank(180),

      // ════════════════════════════════════════════════════════════════════════
      // PARTIE 5 — VOCABULAIRE
      // ════════════════════════════════════════════════════════════════════════
      h1("PARTIE 5 — Exercices de vocabulaire thématique"),
      divider(),

      exerciseBox(
        "✏️  EXERCICE 5A — Associez le terme à sa définition",
        "Reliez chaque mot à sa définition. Tous les termes viennent des thèmes vus en Anki.",
        [
          blank(40),
          new Table({
            width: { size: 8800, type: WidthType.DXA },
            columnWidths: [4200, 4600],
            rows: [
              new TableRow({ children: [
                new TableCell({ borders, width: { size: 4200, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Terme", bold: true, size: 21, color: "FFFFFF", font: "Arial" })] })] }),
                new TableCell({ borders, width: { size: 4600, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Définition", bold: true, size: 21, color: "FFFFFF", font: "Arial" })] })] }),
              ]}),
              ...[
                ["le présentéisme", "a. Fait de se connecter aux outils professionnels hors des heures de travail"],
                ["le plafond de verre", "b. Être physiquement présent au bureau sans être réellement productif"],
                ["le burn-out", "c. Obstacle invisible empêchant les femmes d'accéder aux postes dirigeants"],
                ["l'hyperconnexion", "d. Épuisement professionnel total dû à un stress chronique excessif"],
                ["le droit à la déconnexion", "e. Droit légal des salariés de ne pas répondre aux sollicitations pro hors horaires"],
                ["l'obsolescence programmée", "f. Stratégie industrielle consistant à réduire délibérément la durée de vie des produits"],
                ["le terroir", "g. Lien intime entre un produit agricole et son territoire, son sol et son savoir-faire"],
                ["le lookisme", "h. Discrimination fondée sur l'apparence physique"],
              ].map(([t, d], i) => new TableRow({ children: [
                new TableCell({ borders, width: { size: 4200, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? LIGHT_GREY : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: t, size: 21, font: "Arial" })] })] }),
                new TableCell({ borders, width: { size: 4600, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? LIGHT_YELLOW : "FAFAFA", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: d, size: 21, font: "Arial" })] })] }),
              ]}))
            ]
          }),
          blank(80),
          body("Réponses : _______________________________________"),
        ],
        LIGHT_YELLOW
      ),

      blank(120),

      exerciseBox(
        "✏️  EXERCICE 5B — Complétez avec le bon terme",
        "Choisissez dans la liste : [le brain drain / la médiation culturelle / le volontourisme / l'entomophagie / la sobriété numérique / le report modal / la désertification médicale / l'empowerment]",
        [
          blank(40),
          numbered("La ________ désigne le phénomène par lequel les jeunes diplômés quittent leur pays pour chercher de meilleures opportunités à l'étranger."),
          blank(40),
          numbered("Supprimer ses emails inutiles et limiter les vidéos en streaming contribue à ________, qui représente 2 % des émissions mondiales de CO2."),
          blank(40),
          numbered("Dans les zones rurales françaises, ________ pose un problème de santé publique majeur : certaines régions comptent moins de deux médecins pour 1 000 habitants."),
          blank(40),
          numbered("Passer de la voiture individuelle aux transports en commun ou au vélo s'appelle le ________ — un objectif central des politiques de mobilité durable."),
          blank(40),
          numbered("La consommation d'insectes, ou ________, est présentée comme une solution protéinée durable pour nourrir une population mondiale croissante."),
          blank(40),
          numbered("Certaines associations de ________ font de l'art contemporain accessible à tous en organisant des visites guidées dans les quartiers populaires."),
          blank(40),
        ],
        LIGHT_GREEN
      ),

      blank(120),

      exerciseBox(
        "✏️  EXERCICE 5C — Reformulez sans répéter",
        "Réécrivez chaque phrase en remplaçant le mot souligné par un synonyme ou une tournure équivalente. Plusieurs réponses possibles.",
        [
          blank(40),
          body("1. Le télétravail est AVANTAGEUX pour les salariés.  →"),
          writeLine(), blank(60),
          body("2. Cette mesure est MAUVAISE pour les travailleurs modestes.  →"),
          writeLine(), blank(60),
          body("3. Les femmes SUBISSENT une discrimination à l'embauche.  →"),
          writeLine(), blank(60),
          body("4. Le sport féminin MANQUE de visibilité médiatique.  →"),
          writeLine(), blank(60),
          body("5. La sieste AMÉLIORE les performances des salariés.  →"),
          writeLine(), blank(60),
        ],
        LIGHT_YELLOW
      ),

      blank(180),

      // ════════════════════════════════════════════════════════════════════════
      // PARTIE 6 — MINI-SIMULATION
      // ════════════════════════════════════════════════════════════════════════
      h1("PARTIE 6 — Mini-simulation : plan de prise de parole"),
      divider(),

      body("Utilisez cette fiche pour préparer une prise de parole complète de 3 minutes sur un sujet DELF B2."),
      blank(80),

      new Table({
        width: { size: 9740, type: WidthType.DXA },
        rows: [
          ...[
            ["Sujet choisi", LIGHT_BLUE],
            ["Amorce (1-2 phrases)", "FFFFFF"],
            ["Problématique (question paradoxale ou comment lutter contre)", LIGHT_YELLOW],
            ["Ma position", "FFFFFF"],
            ["Argument 1 POUR + exemple", LIGHT_GREEN],
            ["Argument 2 POUR + exemple", "FFFFFF"],
            ["Argument 3 POUR + exemple", LIGHT_GREEN],
            ["Concession (1 CONTRE + réfutation)", LIGHT_RED],
            ["Conclusion + ouverture", LIGHT_YELLOW],
          ].map(([label, fill]) => new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2800, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: label, size: 21, bold: true, color: "FFFFFF", font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 6940, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "", size: 22, font: "Arial" })] }), new Paragraph({ children: [new TextRun({ text: "", size: 22, font: "Arial" })] })] })
            ]
          }))
        ]
      }),

      blank(200),

      // ── FOOTER ──
      new Table({
        width: { size: 9740, type: WidthType.DXA },
        rows: [new TableRow({ children: [new TableCell({ borders, width: { size: 9740, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Bon courage pour votre DELF B2 ! ", bold: true, size: 24, color: "FFFFFF", font: "Arial" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Révisez avec vos cartes Anki, pratiquez vos amorces à voix haute, et structurez chaque réponse avec une problématique claire.", size: 20, color: "CCE0F5", font: "Arial" })] }),
        ] })] })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/delf_b2_production_orale.docx", buffer);
  console.log("Done!");
});
