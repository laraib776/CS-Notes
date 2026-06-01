const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
  TableOfContents
} = require('docx');
const fs = require('fs');
const path = require('path');

// ─── Color palette ───────────────────────────────────────────────────────────
const C = {
  primary:   "1A56DB",  // rich blue
  secondary: "7E3AF2",  // purple
  accent:    "0E9F6E",  // green
  warning:   "C27804",  // amber
  danger:    "E02424",  // red
  dark:      "111928",  // near-black
  mid:       "374151",  // dark grey
  light:     "6B7280",  // medium grey
  bg1:       "EBF5FF",  // light blue bg
  bg2:       "F5F3FF",  // light purple bg
  bg3:       "ECFDF6",  // light green bg
  bgWarn:    "FFFBEB",  // light amber
  bgCode:    "1E293B",  // dark code bg
  codeText:  "E2E8F0",  // code text
  border:    "D1D5DB",  // grey border
  headerBg:  "1A56DB",  // table header bg
  white:     "FFFFFF",
};

const border = (color = C.border) => ({ style: BorderStyle.SINGLE, size: 1, color });
const borders = (color = C.border) => ({ top: border(color), bottom: border(color), left: border(color), right: border(color) });
const noBorder = () => ({ style: BorderStyle.NONE, size: 0, color: "FFFFFF" });
const noBorders = () => ({ top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder() });

// ─── Helper builders ─────────────────────────────────────────────────────────

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", size: 36, bold: true, color: C.primary })],
    spacing: { before: 400, after: 120 },
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color: C.secondary })],
    spacing: { before: 320, after: 100 },
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: C.mid })],
    spacing: { before: 240, after: 80 },
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Arial", size: 22, color: C.mid, ...opts })],
    spacing: { before: 60, after: 60 },
  });
}

function bodyMixed(runs) {
  return new Paragraph({
    children: runs.map(r => new TextRun({ font: "Arial", size: 22, color: C.mid, ...r })),
    spacing: { before: 60, after: 60 },
  });
}

function bullet(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: C.mid, ...opts })],
    spacing: { before: 40, after: 40 },
  });
}

function numberedItem(text, ref = "numbers") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: C.mid })],
    spacing: { before: 40, after: 40 },
  });
}

function spacer(size = 120) {
  return new Paragraph({ children: [new TextRun("")], spacing: { before: size, after: 0 } });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function divider() {
  return new Paragraph({
    children: [new TextRun("")],
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.border } },
    spacing: { before: 160, after: 160 },
  });
}

// Colored info box (shaded paragraph)
function infoBox(lines, bgColor, accentColor, label = "") {
  const rows = [];
  if (label) {
    rows.push(new TableRow({
      children: [new TableCell({
        borders: noBorders(),
        shading: { fill: accentColor, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 60, left: 200, right: 200 },
        width: { size: 9360, type: WidthType.DXA },
        children: [new Paragraph({
          children: [new TextRun({ text: label, font: "Arial", size: 20, bold: true, color: C.white })],
        })],
      })],
    }));
  }
  rows.push(new TableRow({
    children: [new TableCell({
      borders: { top: noBorder(), bottom: border(accentColor), left: { style: BorderStyle.SINGLE, size: 12, color: accentColor }, right: noBorder() },
      shading: { fill: bgColor, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 200, right: 200 },
      width: { size: 9360, type: WidthType.DXA },
      children: lines.map(l => new Paragraph({
        children: [new TextRun({ text: l, font: "Arial", size: 22, color: C.mid })],
        spacing: { before: 40, after: 40 },
      })),
    })],
  }));
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows,
  });
}

// Code block
function codeBlock(lines) {
  const codeRows = lines.map(line =>
    new Paragraph({
      children: [new TextRun({ text: line || " ", font: "Courier New", size: 18, color: C.codeText })],
      spacing: { before: 20, after: 20 },
    })
  );
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: noBorders(),
        shading: { fill: C.bgCode, type: ShadingType.CLEAR },
        margins: { top: 140, bottom: 140, left: 240, right: 240 },
        width: { size: 9360, type: WidthType.DXA },
        children: codeRows,
      })],
    })],
  });
}

// Two-column comparison table
function compTable(headers, rows, colWidths = [4680, 4680]) {
  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: borders(C.primary),
      shading: { fill: C.primary, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      width: { size: colWidths[i], type: WidthType.DXA },
      children: [new Paragraph({
        children: [new TextRun({ text: h, font: "Arial", size: 22, bold: true, color: C.white })],
      })],
    })),
  });
  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders: borders(C.border),
      shading: { fill: ri % 2 === 0 ? "F9FAFB" : C.white, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 160, right: 160 },
      width: { size: colWidths[ci], type: WidthType.DXA },
      children: [new Paragraph({
        children: [new TextRun({ text: cell, font: "Arial", size: 21, color: C.mid })],
      })],
    })),
  }));
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
  });
}

// Status code badge table
function statusTable(groups) {
  const allRows = [];
  groups.forEach(g => {
    // Group header
    allRows.push(new TableRow({
      children: [
        new TableCell({
          borders: borders(g.color),
          shading: { fill: g.color, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          columnSpan: 3,
          children: [new Paragraph({
            children: [new TextRun({ text: g.group, font: "Arial", size: 22, bold: true, color: C.white })],
          })],
        }),
      ],
    }));
    g.codes.forEach((c, ri) => {
      allRows.push(new TableRow({
        children: [
          new TableCell({
            borders: borders(C.border),
            shading: { fill: ri % 2 === 0 ? "F9FAFB" : C.white, type: ShadingType.CLEAR },
            margins: { top: 70, bottom: 70, left: 160, right: 160 },
            width: { size: 1500, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: c[0], font: "Courier New", size: 21, bold: true, color: g.color })],
            })],
          }),
          new TableCell({
            borders: borders(C.border),
            shading: { fill: ri % 2 === 0 ? "F9FAFB" : C.white, type: ShadingType.CLEAR },
            margins: { top: 70, bottom: 70, left: 160, right: 160 },
            width: { size: 2500, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: c[1], font: "Arial", size: 21, bold: true, color: C.dark })],
            })],
          }),
          new TableCell({
            borders: borders(C.border),
            shading: { fill: ri % 2 === 0 ? "F9FAFB" : C.white, type: ShadingType.CLEAR },
            margins: { top: 70, bottom: 70, left: 160, right: 160 },
            width: { size: 5360, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: c[2], font: "Arial", size: 21, color: C.light })],
            })],
          }),
        ],
      }));
    });
  });
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1500, 2500, 5360],
    rows: allRows,
  });
}

// Section title banner
function sectionBanner(moduleNum, title) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1200, 8160],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: noBorders(),
          shading: { fill: C.primary, type: ShadingType.CLEAR },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 120, bottom: 120, left: 160, right: 160 },
          width: { size: 1200, type: WidthType.DXA },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: moduleNum, font: "Arial", size: 32, bold: true, color: C.white })],
          })],
        }),
        new TableCell({
          borders: noBorders(),
          shading: { fill: "EBF5FF", type: ShadingType.CLEAR },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 120, bottom: 120, left: 240, right: 160 },
          width: { size: 8160, type: WidthType.DXA },
          children: [new Paragraph({
            children: [new TextRun({ text: title, font: "Arial", size: 30, bold: true, color: C.primary })],
          })],
        }),
      ],
    })],
  });
}

// ─── Document ────────────────────────────────────────────────────────────────

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
      {
        reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
      {
        reference: "numbers2",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
      {
        reference: "check",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2713", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
    ],
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: C.primary },
        paragraph: { spacing: { before: 400, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: C.secondary },
        paragraph: { spacing: { before: 320, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: C.mid },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 2 } },
    ],
  },
  sections: [
    // ═══════════════════════════════════════════════════════════════════════
    // COVER PAGE
    // ═══════════════════════════════════════════════════════════════════════
    {
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
      },
      children: [
        spacer(1800),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "API MASTERY", font: "Arial", size: 72, bold: true, color: C.primary })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Complete Course Notes", font: "Arial", size: 40, color: C.secondary })],
          spacing: { before: 120, after: 240 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.primary } },
          children: [new TextRun("")],
          spacing: { before: 0, after: 400 },
        }),
        spacer(200),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "From Zero to Production-Ready APIs", font: "Arial", size: 28, italic: true, color: C.light })],
          spacing: { before: 0, after: 600 },
        }),
        spacer(400),
        // Module list on cover
        ...["Module 01 — What is an API?", "Module 02 — API Architecture & Types", "Module 03 — REST API Deep Dive",
            "Module 04 — HTTP Methods & Status Codes", "Module 05 — Request & Response Structure",
            "Module 06 — Authentication & Security", "Module 07 — Other API Paradigms (GraphQL, gRPC, WebSockets)",
            "Module 08 — API Design Best Practices", "Module 09 — Building APIs with Code",
            "Module 10 — Testing & Documentation", "Module 11 — API Lifecycle, Versioning & Rate Limiting",
            "Module 12 — Advanced Topics & Interview Prep"].map(m =>
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: m, font: "Arial", size: 22, color: C.mid })],
            spacing: { before: 60, after: 60 },
          })
        ),
        spacer(600),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "2025 Edition", font: "Arial", size: 20, color: C.light, italic: true })],
        }),
      ],
    },

    // ═══════════════════════════════════════════════════════════════════════
    // MAIN CONTENT
    // ═══════════════════════════════════════════════════════════════════════
    {
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [
              new TextRun({ text: "API MASTERY — Complete Course Notes", font: "Arial", size: 18, color: C.light }),
              new TextRun({ text: "\t", font: "Arial", size: 18 }),
              new TextRun({ text: "Page ", font: "Arial", size: 18, color: C.light }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: C.light }),
            ],
            tabStops: [{ type: "right", position: 9360 }],
            border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.border } },
          })],
        }),
      },
      children: [

        // TABLE OF CONTENTS
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun({ text: "Table of Contents", font: "Arial", size: 36, bold: true, color: C.primary })],
          spacing: { before: 200, after: 120 },
        }),
        new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
        pageBreak(),

        // ─── MODULE 01 ───────────────────────────────────────────────────
        sectionBanner("01", "What is an API?"),
        spacer(80),

        h1("Module 01 — What is an API?"),

        h2("1.1  Definition"),
        body("API stands for Application Programming Interface. It is a set of rules, protocols, and tools that allows different software applications to communicate and share data with each other."),
        body("In plain language: an API acts as a contract between two systems — it defines what requests can be made, how to make them, and what responses to expect."),
        spacer(80),

        infoBox([
          "Restaurant Analogy (Classic)",
          "",
          "You (Client) sit at a table and look at the menu.",
          "The Waiter (API) takes your order to the kitchen.",
          "The Kitchen (Server/Database) prepares the food.",
          "The Waiter brings back your meal (Response).",
          "",
          "You never enter the kitchen. You don't need to know HOW the food is made.",
          "The API is the defined interface between you and the kitchen."
        ], C.bg1, C.primary, "The Restaurant Analogy"),
        spacer(100),

        h2("1.2  Why APIs Exist"),
        body("Before APIs were widespread, every application had to build everything itself — its own user database, payment processing, maps, etc. APIs changed that by enabling reuse and specialisation."),
        spacer(40),
        compTable(
          ["Without APIs", "With APIs"],
          [
            ["Build your own maps from scratch", "Use Google Maps API"],
            ["Build your own payment processor", "Use Stripe API"],
            ["Build your own SMS gateway", "Use Twilio API"],
            ["Build your own weather data system", "Use OpenWeather API"],
            ["Build your own login system", "Use GitHub/Google OAuth API"],
          ]
        ),
        spacer(120),

        h2("1.3  Real-World Examples"),
        bullet("Weather apps request live data from OpenWeather API"),
        bullet("Ride-sharing apps display routes via Google Maps API"),
        bullet("E-commerce sites process payments via Stripe or PayPal API"),
        bullet("Social login buttons use Google, Facebook, or GitHub OAuth APIs"),
        bullet("Mobile apps send push notifications via Firebase API"),
        bullet("Travel sites aggregate flights from airline APIs"),
        spacer(120),

        h2("1.4  Key Benefits of APIs"),
        compTable(
          ["Benefit", "Explanation"],
          [
            ["Reusability", "One API can serve many clients (web, mobile, desktop, IoT)"],
            ["Modularity", "Change backend without affecting frontend"],
            ["Scalability", "APIs can be scaled independently"],
            ["Interoperability", "Connect systems built in different languages/frameworks"],
            ["Speed of Development", "Use third-party APIs instead of building from scratch"],
            ["Monetisation", "Companies can sell API access as a product"],
          ]
        ),
        spacer(100),

        pageBreak(),

        // ─── MODULE 02 ───────────────────────────────────────────────────
        sectionBanner("02", "API Architecture & Types"),
        spacer(80),

        h1("Module 02 — API Architecture & Types"),

        h2("2.1  How APIs Fit in the Stack"),
        body("A typical modern application has multiple layers. The API layer sits between the frontend (what users see) and the backend (the business logic and database)."),
        spacer(80),
        codeBlock([
          "  Browser / Mobile App",
          "         |",
          "    API Gateway          <-- Rate limiting, Auth, Routing",
          "    /    |    \\",
          " REST  GraphQL  gRPC     <-- API style",
          "   |",
          " Backend Service         <-- Business Logic (Python, Node, Java...)",
          "   |",
          " Database                <-- MySQL, PostgreSQL, MongoDB...",
        ]),
        spacer(120),

        h2("2.2  Types of APIs by Accessibility"),
        spacer(60),

        h3("2.2.1  Open / Public APIs"),
        body("Available to any developer with little or no authentication required. They enable the developer ecosystem and integrations."),
        bullet("OpenWeather API — free weather data"),
        bullet("NASA APIs — space imagery and data"),
        bullet("The Movie Database (TMDB) API"),
        bullet("JSONPlaceholder — fake REST API for testing"),
        spacer(80),

        h3("2.2.2  Partner APIs"),
        body("Shared with specific business partners under contractual agreements. Access is controlled and monitored."),
        bullet("Visa / Mastercard APIs for banking partners"),
        bullet("Shipping APIs shared with select merchants"),
        bullet("Healthcare data APIs shared with certified providers"),
        spacer(80),

        h3("2.2.3  Internal / Private APIs"),
        body("Used only within an organisation. Expose internal systems and data to internal teams, micro-services, or mobile apps owned by the company."),
        codeBlock([
          "  HR System  <--> Internal API <-->  Payroll System",
          "  CRM System <--> Internal API <-->  Billing System",
        ]),
        spacer(80),

        h3("2.2.4  Composite APIs"),
        body("Combine multiple API calls into a single request. Used to reduce round trips and improve performance in complex microservices systems."),
        codeBlock([
          "  // Instead of 3 separate calls:",
          "  GET /users/1",
          "  GET /orders?userId=1",
          "  GET /payments?userId=1",
          "",
          "  // One composite call returns all:",
          "  GET /dashboard/1",
          "  --> { user: {...}, orders: [...], payments: [...] }",
        ]),
        spacer(120),

        h2("2.3  API Communication Styles"),
        compTable(
          ["Style", "Protocol", "Format", "Best For"],
          [
            ["REST", "HTTP/HTTPS", "JSON / XML", "Web & Mobile apps (most common)"],
            ["GraphQL", "HTTP/HTTPS", "JSON", "Flexible data fetching, frontends"],
            ["gRPC", "HTTP/2", "Protocol Buffers", "High-performance microservices"],
            ["SOAP", "HTTP/HTTPS", "XML only", "Enterprise legacy systems"],
            ["WebSocket", "WS / WSS", "JSON / Binary", "Real-time apps (chat, gaming)"],
          ],
          [2000, 1800, 1800, 3760]
        ),
        spacer(100),

        pageBreak(),

        // ─── MODULE 03 ───────────────────────────────────────────────────
        sectionBanner("03", "REST API Deep Dive"),
        spacer(80),

        h1("Module 03 — REST API Deep Dive"),

        h2("3.1  What is REST?"),
        body("REST stands for Representational State Transfer. It is an architectural style, not a protocol. It defines 6 constraints that an API should follow. When an API follows these constraints it is called a RESTful API."),
        spacer(80),
        infoBox([
          "The 6 REST Constraints:",
          "",
          "1.  Client-Server  — frontend and backend are separated",
          "2.  Stateless      — each request contains all necessary information; server stores no session",
          "3.  Cacheable      — responses can be cached to improve performance",
          "4.  Uniform Interface — consistent resource-based URLs and HTTP verbs",
          "5.  Layered System — client doesn't know how many layers exist (load balancer, cache, etc.)",
          "6.  Code on Demand (optional) — server can send executable code (JavaScript)",
        ], C.bg3, C.accent, "REST Constraints"),
        spacer(120),

        h2("3.2  Resources & URLs"),
        body("In REST, everything is a resource. Resources are represented by nouns in the URL. The action is determined by the HTTP method, NOT the URL."),
        spacer(60),
        compTable(
          ["Bad URL (non-RESTful)", "Good URL (RESTful)"],
          [
            ["/getUsers", "/users"],
            ["/createUser", "/users  (POST)"],
            ["/deleteUser/1", "/users/1  (DELETE)"],
            ["/getUserPosts/1", "/users/1/posts"],
            ["/updateUserEmail", "/users/1  (PATCH)"],
          ]
        ),
        spacer(100),

        h2("3.3  Resource Naming Conventions"),
        bullet("Use plural nouns: /users not /user"),
        bullet("Use lowercase with hyphens: /blog-posts not /blogPosts"),
        bullet("Nest for relationships: /users/1/orders/5"),
        bullet("Keep URLs short and readable"),
        bullet("Never use verbs in URLs: /users not /getUsers"),
        spacer(120),

        h2("3.4  HATEOAS"),
        body("HATEOAS (Hypermedia As The Engine Of Application State) is an advanced REST concept where API responses include links to related actions, making the API self-discoverable."),
        codeBlock([
          '{',
          '  "id": 1,',
          '  "name": "Ahmed",',
          '  "links": [',
          '    { "rel": "self",   "href": "/users/1",        "method": "GET" },',
          '    { "rel": "orders", "href": "/users/1/orders", "method": "GET" },',
          '    { "rel": "delete", "href": "/users/1",        "method": "DELETE" }',
          '  ]',
          '}',
        ]),
        spacer(100),

        pageBreak(),

        // ─── MODULE 04 ───────────────────────────────────────────────────
        sectionBanner("04", "HTTP Methods & Status Codes"),
        spacer(80),

        h1("Module 04 — HTTP Methods & Status Codes"),

        h2("4.1  HTTP Methods (CRUD Mapping)"),
        compTable(
          ["HTTP Method", "CRUD Operation", "URL Example", "Description"],
          [
            ["GET", "Read", "/users  or  /users/1", "Retrieve resource(s). No body. Safe & Idempotent."],
            ["POST", "Create", "/users", "Create a new resource. Has request body."],
            ["PUT", "Update (Full)", "/users/1", "Replace entire resource. Idempotent."],
            ["PATCH", "Update (Partial)", "/users/1", "Update specific fields. Non-idempotent."],
            ["DELETE", "Delete", "/users/1", "Remove a resource."],
            ["HEAD", "Read (Meta)", "/users", "Like GET but returns headers only, no body."],
            ["OPTIONS", "Inspect", "/users", "Returns allowed methods. Used in CORS."],
          ],
          [1400, 1600, 2200, 4160]
        ),
        spacer(120),

        h2("4.2  PUT vs PATCH — Key Difference"),
        spacer(60),
        infoBox([
          "PUT — Replace the entire record",
          "",
          'PUT /users/1  with body: { "name": "Ali", "email": "ali@example.com", "age": 25 }',
          "  --> Replaces ALL fields. If you omit 'age', it gets deleted/nulled.",
          "",
          "PATCH — Update only provided fields",
          "",
          'PATCH /users/1  with body: { "email": "newemail@example.com" }',
          "  --> Only email is updated. All other fields remain unchanged.",
        ], C.bgWarn, C.warning, "PUT vs PATCH"),
        spacer(120),

        h2("4.3  HTTP Status Codes — Complete Reference"),
        spacer(80),
        statusTable([
          {
            group: "2xx — Success",
            color: C.accent,
            codes: [
              ["200", "OK", "Standard success response for GET, PUT, PATCH, DELETE"],
              ["201", "Created", "Resource successfully created — used with POST"],
              ["202", "Accepted", "Request accepted but processing not complete (async)"],
              ["204", "No Content", "Success but nothing to return — common for DELETE"],
            ],
          },
          {
            group: "3xx — Redirection",
            color: C.warning,
            codes: [
              ["301", "Moved Permanently", "URL has permanently changed — update your links"],
              ["304", "Not Modified", "Cached version is still valid — no new data sent"],
            ],
          },
          {
            group: "4xx — Client Errors",
            color: C.danger,
            codes: [
              ["400", "Bad Request", "Malformed syntax or invalid request parameters"],
              ["401", "Unauthorized", "No credentials or invalid credentials provided"],
              ["403", "Forbidden", "Authenticated but does not have permission"],
              ["404", "Not Found", "Resource does not exist"],
              ["405", "Method Not Allowed", "HTTP method not supported on this endpoint"],
              ["409", "Conflict", "Resource conflict (e.g. duplicate entry)"],
              ["422", "Unprocessable Entity", "Validation failed on input data"],
              ["429", "Too Many Requests", "Rate limit exceeded"],
            ],
          },
          {
            group: "5xx — Server Errors",
            color: C.secondary,
            codes: [
              ["500", "Internal Server Error", "Generic server crash — check logs"],
              ["502", "Bad Gateway", "Upstream server returned invalid response"],
              ["503", "Service Unavailable", "Server overloaded or down for maintenance"],
              ["504", "Gateway Timeout", "Upstream server did not respond in time"],
            ],
          },
        ]),
        spacer(100),

        pageBreak(),

        // ─── MODULE 05 ───────────────────────────────────────────────────
        sectionBanner("05", "Request & Response Structure"),
        spacer(80),

        h1("Module 05 — Request & Response Structure"),

        h2("5.1  Anatomy of an HTTP Request"),
        codeBlock([
          "  POST /api/v1/users HTTP/1.1",
          "  Host: api.example.com",
          "  Authorization: Bearer eyJhbGciOiJIUzI1NiIsInR5c...",
          "  Content-Type: application/json",
          "  Accept: application/json",
          "  User-Agent: MyApp/2.0",
          "  X-Request-ID: 550e8400-e29b-41d4-a716-446655440000",
          "",
          "  {",
          '    "name": "Laraib",',
          '    "email": "laraib@example.com",',
          '    "role": "admin"',
          "  }",
        ]),
        spacer(80),
        compTable(
          ["Part", "Description"],
          [
            ["Request Line", "Method + URL + HTTP version"],
            ["Host", "Domain of the server"],
            ["Authorization", "Auth token (Bearer, API Key, Basic)"],
            ["Content-Type", "Format of the request body (application/json)"],
            ["Accept", "Format the client wants in the response"],
            ["Body", "Data sent with POST / PUT / PATCH requests"],
            ["Query Params", "Filters/sorting appended to URL: ?page=2&limit=10"],
            ["Path Params", "Variables in the URL: /users/{id}"],
          ]
        ),
        spacer(120),

        h2("5.2  Query Parameters vs Path Parameters"),
        spacer(60),
        infoBox([
          "Path Parameters — identify a specific resource:",
          "  GET /users/42          --> fetch user with id 42",
          "  GET /posts/7/comments  --> fetch comments of post 7",
          "",
          "Query Parameters — filter, sort, or paginate:",
          "  GET /users?role=admin&page=2&limit=20",
          "  GET /products?category=electronics&sort=price_asc",
          "  GET /search?q=API+tutorial&lang=en",
        ], C.bg2, C.secondary, "Path Params vs Query Params"),
        spacer(120),

        h2("5.3  Anatomy of an HTTP Response"),
        codeBlock([
          "  HTTP/1.1 201 Created",
          "  Content-Type: application/json",
          "  X-Request-ID: 550e8400-e29b-41d4-a716-446655440000",
          "  Date: Tue, 10 Jun 2025 10:30:00 GMT",
          "",
          "  {",
          '    "status": "success",',
          '    "data": {',
          '      "id": 101,',
          '      "name": "Laraib",',
          '      "email": "laraib@example.com",',
          '      "created_at": "2025-06-10T10:30:00Z"',
          "    },",
          '    "message": "User created successfully"',
          "  }",
        ]),
        spacer(120),

        h2("5.4  Pagination Patterns"),
        h3("Offset-based Pagination"),
        codeBlock([
          "  GET /posts?page=3&limit=10",
          "  --> Returns posts 21-30",
        ]),
        h3("Cursor-based Pagination (preferred for large datasets)"),
        codeBlock([
          "  GET /posts?cursor=eyJpZCI6MjB9&limit=10",
          "  Response includes: { next_cursor: 'eyJpZCI6MzB9', has_more: true }",
        ]),
        spacer(100),

        pageBreak(),

        // ─── MODULE 06 ───────────────────────────────────────────────────
        sectionBanner("06", "Authentication & Security"),
        spacer(80),

        h1("Module 06 — Authentication & Security"),

        h2("6.1  Authentication vs Authorisation"),
        compTable(
          ["Term", "Question it Answers", "Example"],
          [
            ["Authentication (AuthN)", "Who are you?", "Username/password login"],
            ["Authorisation (AuthZ)", "What are you allowed to do?", "Admin can delete; user cannot"],
          ],
          [2000, 3000, 4360]
        ),
        spacer(120),

        h2("6.2  API Key Authentication"),
        body("A unique secret string issued to a client. Simple but limited — no expiry, no user context, difficult to rotate."),
        codeBlock([
          "  // In query parameter (less secure):",
          "  GET /data?api_key=sk-abc123xyz",
          "",
          "  // In header (preferred):",
          "  GET /data",
          "  x-api-key: sk-abc123xyz",
        ]),
        spacer(80),
        infoBox([
          "API Key Best Practices:",
          "  - Never hardcode in source code. Use environment variables.",
          "  - Rotate keys regularly.",
          "  - Scope keys to specific permissions.",
          "  - Log and monitor key usage.",
          "  - Use HTTPS always.",
        ], C.bgWarn, C.warning),
        spacer(120),

        h2("6.3  Basic Authentication"),
        body("Encodes username:password in Base64 and sends in the Authorization header. Simple but insecure unless over HTTPS. Avoid for production APIs."),
        codeBlock([
          "  // username:password --> Base64 encode",
          '  "admin:secret123" --> "YWRtaW46c2VjcmV0MTIz"',
          "",
          "  Authorization: Basic YWRtaW46c2VjcmV0MTIz",
        ]),
        spacer(120),

        h2("6.4  JWT Authentication (JSON Web Tokens)"),
        body("The most widely used modern authentication method. A JWT is a self-contained token with three parts separated by dots."),
        codeBlock([
          "  eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9   <-- Header (algorithm + type)",
          "  .eyJ1c2VySWQiOjEsInJvbGUiOiJhZG1pbiJ9   <-- Payload (claims / data)",
          "  .SflKxwRJSMeKKF2QT4fwpMeJf36POk6yJV_adQ  <-- Signature (HMAC of header+payload)",
        ]),
        spacer(80),
        infoBox([
          "JWT Flow:",
          "1.  User logs in with credentials.",
          "2.  Server validates and returns a JWT.",
          "3.  Client stores JWT (localStorage / httpOnly cookie).",
          "4.  Client sends JWT in every subsequent request.",
          "5.  Server verifies signature — no database lookup needed.",
          "6.  Token expires (typically 15 min - 24h).",
          "7.  Client uses refresh token to get a new access token.",
        ], C.bg3, C.accent, "JWT Authentication Flow"),
        spacer(80),
        codeBlock([
          "  // Client sends token in every request:",
          "  Authorization: Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...",
        ]),
        spacer(120),

        h2("6.5  OAuth 2.0"),
        body("OAuth 2.0 is an authorisation framework that allows a user to grant a third-party application access to their data on another service, without sharing credentials."),
        spacer(60),
        codeBlock([
          "  OAuth 2.0 Flow (Authorization Code Grant):",
          "",
          "  1. User clicks 'Login with Google'",
          "  2. App redirects to Google's auth server",
          "  3. User grants permission on Google's page",
          "  4. Google redirects back with an authorization code",
          "  5. App exchanges code for access token (server-side)",
          "  6. App uses access token to call Google APIs",
        ]),
        spacer(80),
        compTable(
          ["OAuth Grant Type", "Use Case"],
          [
            ["Authorization Code", "Web apps with a backend (most secure)"],
            ["Authorization Code + PKCE", "Mobile / SPA apps (no backend secret)"],
            ["Client Credentials", "Server-to-server (no user involved)"],
            ["Implicit (deprecated)", "Old SPA flow — no longer recommended"],
          ]
        ),
        spacer(120),

        h2("6.6  API Security Best Practices"),
        spacer(60),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [480, 8880],
          rows: [
            ["Always use HTTPS / TLS 1.2+", "Encrypts data in transit; never use plain HTTP"],
            ["Validate all input", "Prevent SQL injection, XSS, and other attacks"],
            ["Use JWT or OAuth", "Avoid Basic Auth in production"],
            ["Implement Rate Limiting", "Prevent brute force and DDoS attacks"],
            ["Return minimal data", "Only expose what the client needs (least privilege)"],
            ["Use CORS properly", "Restrict which origins can call your API"],
            ["Log & Monitor", "Track errors, access patterns, and anomalies"],
            ["Version your API", "Avoid breaking changes for existing consumers"],
            ["Sanitise error messages", "Never expose stack traces or DB structure"],
          ].map((row, i) => new TableRow({
            children: [
              new TableCell({
                borders: noBorders(),
                shading: { fill: C.accent, type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 160, right: 160 },
                width: { size: 480, type: WidthType.DXA },
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text: "\u2713", font: "Arial", size: 24, bold: true, color: C.white })],
                })],
              }),
              new TableCell({
                borders: { top: noBorder(), bottom: border(C.border), left: noBorder(), right: noBorder() },
                shading: { fill: i % 2 === 0 ? C.bg3 : C.white, type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 200, right: 160 },
                width: { size: 8880, type: WidthType.DXA },
                children: [new Paragraph({
                  children: [
                    new TextRun({ text: row[0] + " — ", font: "Arial", size: 22, bold: true, color: C.dark }),
                    new TextRun({ text: row[1], font: "Arial", size: 22, color: C.mid }),
                  ],
                })],
              }),
            ],
          })),
        }),
        spacer(100),

        pageBreak(),

        // ─── MODULE 07 ───────────────────────────────────────────────────
        sectionBanner("07", "Other API Paradigms"),
        spacer(80),

        h1("Module 07 — Other API Paradigms"),

        h2("7.1  GraphQL"),
        body("Developed by Meta (Facebook) in 2012, open-sourced in 2015. GraphQL solves a fundamental REST limitation: over-fetching and under-fetching data."),
        spacer(60),
        infoBox([
          "REST Problem:  GET /users/1  returns ALL user fields even if you only need the name.",
          "GraphQL Solution:  you ask for exactly what you need — nothing more, nothing less.",
        ], C.bg1, C.primary),
        spacer(80),
        h3("GraphQL Query Example"),
        codeBlock([
          "  # Request exactly the fields you need:",
          "  query {",
          "    user(id: 1) {",
          "      name",
          "      email",
          "      posts {",
          "        title",
          "      }",
          "    }",
          "  }",
          "",
          "  # Response — only what was requested:",
          "  {",
          '    "data": {',
          '      "user": {',
          '        "name": "Laraib",',
          '        "email": "laraib@example.com",',
          '        "posts": [{ "title": "My first post" }]',
          "      }",
          "    }",
          "  }",
        ]),
        spacer(80),
        compTable(
          ["REST", "GraphQL"],
          [
            ["Multiple endpoints (/users, /posts)", "Single endpoint /graphql"],
            ["Server decides what data is returned", "Client specifies exact fields"],
            ["Over-fetching is common", "No over/under-fetching"],
            ["Easy to cache with HTTP", "Caching is more complex"],
            ["Great for simple CRUD", "Great for complex, nested data"],
          ]
        ),
        spacer(120),

        h2("7.2  gRPC"),
        body("gRPC (Google Remote Procedure Call) is a high-performance framework developed by Google. Instead of JSON, it uses Protocol Buffers (protobuf) — a binary serialisation format that is much faster and smaller."),
        codeBlock([
          "  // Define the service in a .proto file:",
          "  syntax = 'proto3';",
          "",
          "  service UserService {",
          "    rpc GetUser(UserRequest) returns (UserResponse);",
          "    rpc ListUsers(Empty) returns (stream UserResponse);",
          "  }",
          "",
          "  message UserRequest { int32 id = 1; }",
          "  message UserResponse {",
          "    int32 id = 1;",
          "    string name = 2;",
          "    string email = 3;",
          "  }",
        ]),
        spacer(80),
        compTable(
          ["Feature", "REST", "gRPC"],
          [
            ["Protocol", "HTTP/1.1", "HTTP/2"],
            ["Format", "JSON (text)", "Protocol Buffers (binary)"],
            ["Speed", "Slower", "5-10x faster"],
            ["Browser Support", "Native", "Requires proxy"],
            ["Best For", "Public APIs, web apps", "Microservices, internal services"],
          ],
          [2400, 3480, 3480]
        ),
        spacer(120),

        h2("7.3  WebSocket API"),
        body("HTTP is request-response: the client must ask before the server can respond. WebSockets provide a persistent two-way connection where both sides can send data at any time."),
        codeBlock([
          "  // HTTP:  Client --request--> Server --response--> Client (then connection closes)",
          "  // WebSocket: Client <====== persistent connection ======> Server",
          "",
          "  // JavaScript WebSocket client:",
          "  const ws = new WebSocket('wss://chat.example.com/socket');",
          "",
          "  ws.onopen = () => ws.send(JSON.stringify({ type: 'join', room: 'general' }));",
          "  ws.onmessage = (event) => console.log('New message:', event.data);",
          "  ws.onclose = () => console.log('Connection closed');",
        ]),
        spacer(80),
        body("WebSockets are used in: chat apps, live sports scores, collaborative editors (Google Docs style), online multiplayer games, financial trading platforms, and IoT dashboards."),
        spacer(120),

        h2("7.4  SOAP API"),
        body("SOAP (Simple Object Access Protocol) is an older, XML-based messaging protocol. It is still used in enterprise, banking, healthcare, and government systems due to its strict standards and built-in security (WS-Security)."),
        codeBlock([
          "  <?xml version='1.0'?>",
          "  <soap:Envelope xmlns:soap='http://www.w3.org/2003/05/soap-envelope'>",
          "    <soap:Header>",
          "      <Authentication>",
          "        <Token>abc123</Token>",
          "      </Authentication>",
          "    </soap:Header>",
          "    <soap:Body>",
          "      <GetUserRequest>",
          "        <UserId>42</UserId>",
          "      </GetUserRequest>",
          "    </soap:Body>",
          "  </soap:Envelope>",
        ]),
        spacer(100),

        pageBreak(),

        // ─── MODULE 08 ───────────────────────────────────────────────────
        sectionBanner("08", "API Design Best Practices"),
        spacer(80),

        h1("Module 08 — API Design Best Practices"),

        h2("8.1  RESTful Design Principles"),
        bullet("Use nouns for resources, never verbs in URLs"),
        bullet("Use plural nouns: /users not /user"),
        bullet("Use HTTP methods to express the action"),
        bullet("Use correct status codes in every response"),
        bullet("Version your API from the start: /api/v1/"),
        bullet("Keep URLs lowercase with hyphens: /blog-posts"),
        bullet("Filter, sort, paginate via query parameters"),
        bullet("Use consistent response envelope format"),
        spacer(100),

        h2("8.2  Consistent Response Format"),
        body("A consistent response structure makes your API predictable and easier to consume."),
        codeBlock([
          "  // Success Response:",
          "  {",
          '    "status": "success",',
          '    "code": 200,',
          '    "data": { ... },',
          '    "meta": { "page": 1, "total": 100, "per_page": 10 }',
          "  }",
          "",
          "  // Error Response:",
          "  {",
          '    "status": "error",',
          '    "code": 422,',
          '    "message": "Validation failed",',
          '    "errors": [',
          '      { "field": "email", "message": "Must be a valid email address" }',
          "    ]",
          "  }",
        ]),
        spacer(120),

        h2("8.3  API Versioning Strategies"),
        compTable(
          ["Strategy", "Example", "Pros / Cons"],
          [
            ["URL Versioning", "/api/v1/users", "Easy to see and test. Recommended."],
            ["Header Versioning", "Accept: application/vnd.api.v2+json", "Clean URLs. Less discoverable."],
            ["Query Parameter", "/api/users?version=2", "Easy for testing. Not RESTful."],
            ["Subdomain", "v2.api.example.com/users", "Separate deployments. Complex."],
          ],
          [2000, 2800, 4560]
        ),
        spacer(120),

        h2("8.4  Rate Limiting"),
        body("Rate limiting controls how many requests a client can make in a given time window. It protects your server from abuse, DDoS attacks, and unintentional overload."),
        codeBlock([
          "  // Response headers tell the client about limits:",
          "  X-RateLimit-Limit: 1000",
          "  X-RateLimit-Remaining: 450",
          "  X-RateLimit-Reset: 1718012400",
          "  Retry-After: 60    (sent with 429 Too Many Requests)",
        ]),
        spacer(80),
        compTable(
          ["Algorithm", "How it Works", "Best For"],
          [
            ["Fixed Window", "N requests per fixed time window", "Simple use cases"],
            ["Sliding Window", "N requests in any rolling time window", "Smoother limiting"],
            ["Token Bucket", "Tokens refill at constant rate; each request consumes one", "Burst traffic"],
            ["Leaky Bucket", "Requests processed at fixed rate; excess queued or dropped", "Steady output"],
          ],
          [2000, 3500, 3860]
        ),
        spacer(120),

        h2("8.5  Idempotency"),
        body("An operation is idempotent if calling it multiple times produces the same result as calling it once. This is critical for reliability in distributed systems."),
        compTable(
          ["Method", "Idempotent?", "Why?"],
          [
            ["GET", "Yes", "Reading data never changes state"],
            ["PUT", "Yes", "Replacing a record with the same data has the same result"],
            ["DELETE", "Yes", "Deleting an already-deleted resource still results in it being gone"],
            ["POST", "No", "Creates a new resource each time"],
            ["PATCH", "No", "May have cumulative effects (e.g. increment a counter)"],
          ],
          [2000, 1800, 5560]
        ),
        spacer(100),

        pageBreak(),

        // ─── MODULE 09 ───────────────────────────────────────────────────
        sectionBanner("09", "Building APIs with Code"),
        spacer(80),

        h1("Module 09 — Building APIs with Code"),

        h2("9.1  Python — FastAPI"),
        body("FastAPI is the most modern and fastest Python web framework for building APIs. It uses Python type hints to generate automatic validation, serialisation, and documentation."),
        codeBlock([
          "  # Install: pip install fastapi uvicorn",
          "",
          "  from fastapi import FastAPI, HTTPException, Depends",
          "  from pydantic import BaseModel",
          "  from typing import Optional",
          "",
          "  app = FastAPI(title='My API', version='1.0')",
          "",
          "  class UserCreate(BaseModel):",
          "      name: str",
          "      email: str",
          "      age: Optional[int] = None",
          "",
          "  fake_db = {}",
          "",
          "  @app.get('/')",
          "  def root():",
          "      return {'message': 'Welcome to My API'}",
          "",
          "  @app.get('/users/{user_id}')",
          "  def get_user(user_id: int):",
          "      if user_id not in fake_db:",
          "          raise HTTPException(status_code=404, detail='User not found')",
          "      return fake_db[user_id]",
          "",
          "  @app.post('/users', status_code=201)",
          "  def create_user(user: UserCreate):",
          "      user_id = len(fake_db) + 1",
          "      fake_db[user_id] = {'id': user_id, **user.dict()}",
          "      return fake_db[user_id]",
          "",
          "  # Run: uvicorn main:app --reload",
          "  # Docs at: http://localhost:8000/docs",
        ]),
        spacer(100),

        h2("9.2  Python — Flask"),
        codeBlock([
          "  # Install: pip install flask",
          "",
          "  from flask import Flask, request, jsonify",
          "",
          "  app = Flask(__name__)",
          "  users = []",
          "",
          "  @app.route('/users', methods=['GET'])",
          "  def get_users():",
          "      return jsonify({'data': users, 'count': len(users)}), 200",
          "",
          "  @app.route('/users', methods=['POST'])",
          "  def create_user():",
          "      data = request.get_json()",
          "      users.append(data)",
          "      return jsonify({'message': 'Created', 'data': data}), 201",
          "",
          "  if __name__ == '__main__':",
          "      app.run(debug=True)",
        ]),
        spacer(100),

        h2("9.3  Node.js — Express"),
        codeBlock([
          "  // Install: npm install express",
          "",
          "  const express = require('express');",
          "  const app = express();",
          "  app.use(express.json());",
          "",
          "  const users = [];",
          "",
          "  app.get('/users', (req, res) => {",
          "    res.json({ data: users, count: users.length });",
          "  });",
          "",
          "  app.post('/users', (req, res) => {",
          "    const user = { id: users.length + 1, ...req.body };",
          "    users.push(user);",
          "    res.status(201).json({ data: user });",
          "  });",
          "",
          "  app.get('/users/:id', (req, res) => {",
          "    const user = users.find(u => u.id === parseInt(req.params.id));",
          "    if (!user) return res.status(404).json({ error: 'Not found' });",
          "    res.json({ data: user });",
          "  });",
          "",
          "  app.listen(3000, () => console.log('Server running on port 3000'));",
        ]),
        spacer(100),

        h2("9.4  Typical FastAPI Project Structure"),
        codeBlock([
          "  project/",
          "  |",
          "  |-- main.py                  # App entry point",
          "  |-- requirements.txt",
          "  |-- .env                     # Environment variables (never commit!)",
          "  |",
          "  |-- routers/",
          "  |   |-- users.py             # /users endpoints",
          "  |   |-- products.py          # /products endpoints",
          "  |   +-- auth.py              # /auth endpoints",
          "  |",
          "  |-- models/",
          "  |   +-- user.py              # SQLAlchemy / database models",
          "  |",
          "  |-- schemas/",
          "  |   +-- user.py              # Pydantic request/response schemas",
          "  |",
          "  |-- services/",
          "  |   +-- user_service.py      # Business logic layer",
          "  |",
          "  |-- database/",
          "  |   +-- connection.py        # DB connection and session",
          "  |",
          "  +-- middleware/",
          "      +-- auth.py              # JWT verification, CORS",
        ]),
        spacer(100),

        pageBreak(),

        // ─── MODULE 10 ───────────────────────────────────────────────────
        sectionBanner("10", "Testing & Documentation"),
        spacer(80),

        h1("Module 10 — Testing & Documentation"),

        h2("10.1  API Testing with Postman"),
        body("Postman is the most popular API testing tool. It allows you to build, test, and document API requests without writing code."),
        spacer(60),
        compTable(
          ["Postman Feature", "What it Does"],
          [
            ["Collections", "Group related requests together"],
            ["Environments", "Switch between dev/staging/prod variables"],
            ["Tests (JavaScript)", "Write automated assertions on responses"],
            ["Mock Servers", "Simulate an API before it is built"],
            ["Pre-request Scripts", "Run code before a request (e.g. get a token)"],
            ["Newman (CLI)", "Run Postman collections in CI/CD pipelines"],
          ]
        ),
        spacer(80),
        codeBlock([
          "  // Postman test script example (JavaScript):",
          "  pm.test('Status is 200', () => {",
          "      pm.response.to.have.status(200);",
          "  });",
          "",
          "  pm.test('Response has user id', () => {",
          "      const json = pm.response.json();",
          "      pm.expect(json.data.id).to.be.a('number');",
          "  });",
          "",
          "  // Save token from login response to environment:",
          "  const token = pm.response.json().token;",
          "  pm.environment.set('auth_token', token);",
        ]),
        spacer(120),

        h2("10.2  Testing with cURL"),
        body("cURL is a command-line tool for making HTTP requests. Essential for quick tests and scripting."),
        codeBlock([
          "  # GET request:",
          "  curl https://api.example.com/users",
          "",
          "  # POST request with JSON body:",
          "  curl -X POST https://api.example.com/users \\",
          "       -H 'Content-Type: application/json' \\",
          "       -H 'Authorization: Bearer YOUR_TOKEN' \\",
          "       -d '{\"name\": \"Laraib\", \"email\": \"laraib@example.com\"}'",
          "",
          "  # PATCH request:",
          "  curl -X PATCH https://api.example.com/users/1 \\",
          "       -H 'Content-Type: application/json' \\",
          "       -d '{\"email\": \"new@example.com\"}'",
          "",
          "  # DELETE request:",
          "  curl -X DELETE https://api.example.com/users/1 \\",
          "       -H 'Authorization: Bearer YOUR_TOKEN'",
        ]),
        spacer(120),

        h2("10.3  API Documentation"),
        body("Good API documentation is essential. It is the first thing a developer sees when integrating with your API. Poor docs = poor adoption."),
        spacer(60),
        compTable(
          ["Tool", "Description", "Format"],
          [
            ["Swagger UI", "Interactive docs — test endpoints in the browser", "OpenAPI (YAML/JSON)"],
            ["Redoc", "Beautiful read-only documentation", "OpenAPI"],
            ["Postman Docs", "Auto-generate docs from collections", "Postman format"],
            ["GitBook", "Write docs manually as a book", "Markdown"],
            ["Stoplight", "Visual API design and documentation", "OpenAPI"],
          ],
          [2000, 4360, 2000 + 1000]
        ),
        spacer(80),

        h3("What Good API Docs Must Include"),
        bullet("Authentication method and how to obtain credentials"),
        bullet("Base URL and environments (production, sandbox)"),
        bullet("Complete endpoint reference with request/response examples"),
        bullet("All parameters: path, query, header, body"),
        bullet("All possible response codes and error messages"),
        bullet("Rate limiting information"),
        bullet("SDK / code examples in multiple languages"),
        bullet("Changelog and versioning history"),
        spacer(100),

        pageBreak(),

        // ─── MODULE 11 ───────────────────────────────────────────────────
        sectionBanner("11", "API Lifecycle & Advanced Concepts"),
        spacer(80),

        h1("Module 11 — API Lifecycle & Advanced Concepts"),

        h2("11.1  The API Development Lifecycle"),
        spacer(60),
        infoBox([
          "1.  Design      — Define endpoints, data models, auth strategy, and versioning",
          "2.  Development — Implement the API with proper error handling and validation",
          "3.  Testing     — Unit tests, integration tests, contract tests, load tests",
          "4.  Documentation — Write or auto-generate documentation",
          "5.  Deployment  — Deploy to cloud (AWS, GCP, Azure, Railway, Render...)",
          "6.  Monitoring  — Track response times, error rates, throughput",
          "7.  Maintenance — Bug fixes, security patches, version deprecation",
        ], C.bg1, C.primary, "API Lifecycle"),
        spacer(120),

        h2("11.2  API Gateway"),
        body("An API Gateway is a server that acts as the single entry point for all client requests. It routes requests to the correct microservice and handles cross-cutting concerns."),
        spacer(60),
        codeBlock([
          "  Client",
          "    |",
          "    v",
          "  API Gateway  <-- Authentication, Rate Limiting, Logging, CORS, SSL termination",
          "  /   |   \\",
          " v    v    v",
          "User  Order  Payment   <-- Individual microservices",
          "Service Service Service",
        ]),
        spacer(80),
        body("Popular API Gateways: AWS API Gateway, Kong, NGINX, Traefik, Apigee, Azure API Management."),
        spacer(120),

        h2("11.3  Microservices and APIs"),
        body("In a microservices architecture, each service is a small, independent application with its own API. Services communicate with each other through APIs."),
        compTable(
          ["Communication Pattern", "Description", "When to Use"],
          [
            ["Synchronous (REST/gRPC)", "Service waits for response before continuing", "Real-time data, user-facing requests"],
            ["Asynchronous (Message Queue)", "Service sends message and continues; no wait", "Background jobs, email, notifications"],
            ["Event-Driven", "Services react to events published by others", "Decoupled, scalable workflows"],
          ],
          [2400, 3500, 3460]
        ),
        spacer(120),

        h2("11.4  Caching in APIs"),
        body("Caching stores API responses so repeated requests are served faster without hitting the database."),
        compTable(
          ["Cache Type", "Where", "Example"],
          [
            ["HTTP Cache Headers", "Browser / CDN", "Cache-Control: max-age=3600"],
            ["Redis Cache", "Server-side in-memory", "Cache DB query results for 5 minutes"],
            ["CDN Cache", "Edge servers globally", "Cache static API responses near users"],
            ["ETags", "Browser + Server", "Only re-fetch if resource has changed"],
          ],
          [2000, 3000, 4360]
        ),
        spacer(120),

        h2("11.5  API Monitoring & Observability"),
        body("The three pillars of API observability are Metrics, Logs, and Traces. Together they help you understand what is happening inside your API."),
        spacer(60),
        compTable(
          ["Pillar", "What it Captures", "Tools"],
          [
            ["Metrics", "Throughput, latency, error rate, uptime", "Prometheus + Grafana, Datadog"],
            ["Logs", "Every request/response, error details, auth events", "ELK Stack, Loki, CloudWatch"],
            ["Traces", "End-to-end journey of a single request across services", "Jaeger, Zipkin, OpenTelemetry"],
          ],
          [1600, 3760, 4000]
        ),
        spacer(100),

        pageBreak(),

        // ─── MODULE 12 ───────────────────────────────────────────────────
        sectionBanner("12", "Interview Prep & Quick Reference"),
        spacer(80),

        h1("Module 12 — Interview Prep & Quick Reference"),

        h2("12.1  Essential Interview Questions"),
        spacer(40),

        h3("Beginner Questions"),
        bullet("What is an API and why are they used?"),
        bullet("What is the difference between GET and POST?"),
        bullet("What is REST? What makes an API RESTful?"),
        bullet("What are HTTP status codes? Give examples."),
        bullet("What is JSON and why do APIs use it?"),
        bullet("What is the difference between authentication and authorisation?"),
        spacer(80),

        h3("Intermediate Questions"),
        bullet("What is the difference between PUT and PATCH?"),
        bullet("What is JWT? How does it work?"),
        bullet("What is OAuth 2.0 and when would you use it?"),
        bullet("What is rate limiting and how do you implement it?"),
        bullet("What is API versioning? What strategy do you prefer and why?"),
        bullet("What is CORS and why is it needed?"),
        bullet("What is an API Gateway?"),
        bullet("What is idempotency? Which HTTP methods are idempotent?"),
        spacer(80),

        h3("Advanced Questions"),
        bullet("What is the difference between REST, GraphQL, and gRPC?"),
        bullet("How would you design an API for a Twitter-like feed?"),
        bullet("How do you handle API backward compatibility?"),
        bullet("How do you secure an API in production?"),
        bullet("What is HATEOAS?"),
        bullet("How do you implement pagination for a large dataset?"),
        bullet("What is the N+1 problem in GraphQL and how do you solve it?"),
        bullet("How do you monitor API performance in production?"),
        spacer(120),

        h2("12.2  Model Answers — Key Questions"),
        spacer(60),

        infoBox([
          "Q: What is the difference between REST and GraphQL?",
          "",
          "REST: Multiple endpoints, server decides what data is returned, easy to cache.",
          "GraphQL: Single endpoint (/graphql), client specifies exact fields, reduces over-fetching.",
          "Use REST for: simple CRUD APIs, public APIs, microservices.",
          "Use GraphQL for: complex frontends with many data needs, mobile apps where bandwidth matters.",
        ], C.bg1, C.primary),
        spacer(80),

        infoBox([
          "Q: How does JWT authentication work?",
          "",
          "1. User logs in --> Server validates credentials",
          "2. Server creates JWT: Header.Payload.Signature (signed with secret key)",
          "3. Client stores JWT and sends it in Authorization: Bearer <token>",
          "4. Server verifies signature on every request -- no DB lookup needed",
          "5. Token expires; client uses refresh token to obtain new access token",
          "Key point: JWTs are stateless -- the server doesn't store sessions.",
        ], C.bg3, C.accent),
        spacer(80),

        infoBox([
          "Q: What is idempotency and why does it matter?",
          "",
          "An operation is idempotent if calling it multiple times produces the same result.",
          "GET, PUT, DELETE are idempotent. POST and PATCH are NOT.",
          "",
          "Why it matters: In distributed systems, network failures happen. If a client",
          "retries a request, idempotent APIs prevent duplicate side effects.",
          "Best practice: Use an Idempotency-Key header on POST requests for payment APIs.",
        ], C.bg2, C.secondary),
        spacer(120),

        h2("12.3  Quick Reference Cheatsheet"),
        spacer(60),
        compTable(
          ["Concept", "One-Line Summary"],
          [
            ["REST", "Architectural style using HTTP methods on resource URLs"],
            ["GraphQL", "Query language — client asks for exactly the data it needs"],
            ["gRPC", "Binary RPC framework — fast microservice communication"],
            ["WebSocket", "Persistent two-way connection for real-time communication"],
            ["SOAP", "XML-based protocol for enterprise / legacy systems"],
            ["JWT", "Self-contained signed token for stateless authentication"],
            ["OAuth 2.0", "Delegated authorisation — log in with Google/GitHub"],
            ["API Gateway", "Single entry point handling auth, routing, rate limiting"],
            ["Rate Limiting", "Restrict requests per time window to prevent abuse"],
            ["Idempotency", "Calling the same operation multiple times = same result"],
            ["CORS", "Browser security policy restricting cross-origin requests"],
            ["OpenAPI", "Standard specification format for documenting REST APIs"],
          ]
        ),
        spacer(100),

        divider(),
        spacer(60),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "End of API Mastery — Complete Course Notes", font: "Arial", size: 24, bold: true, color: C.primary })],
          spacing: { before: 80, after: 40 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Good luck in your interviews and projects!", font: "Arial", size: 22, italic: true, color: C.light })],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  const outputPath = path.join(__dirname, 'API_Mastery_Complete_Course_Notes.docx');
  fs.writeFileSync(outputPath, buffer);
  console.log(`Done! Wrote ${outputPath}`);
}).catch(console.error);
