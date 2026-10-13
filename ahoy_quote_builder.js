// ahoy_quote_builder.js
// Usage: node ahoy_quote_builder.js spec.json out.docx [logo.png]
// If logo.png is omitted, the embedded Ahoy logo below is used.
// Builds an Ahoy Creative quote / work order as .docx (A4, RTL, Assistant font, logo header, footer).
// Output: a .docx handed to the user; integrity is verified by sha256 (see Project instructions).

const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
  ImageRun, AlignmentType, HeadingLevel, WidthType, BorderStyle, ShadingType,
  LevelFormat, PageNumber, NumberFormat, TabStopType, PageBreak, VerticalAlign
} = require("docx");

const [,, specPath, outPath, logoPath] = process.argv;
const spec = JSON.parse(fs.readFileSync(specPath, "utf8"));

const FONT = "Assistant";
const YELLOW = "FFCC33";
const BLACK = "000000";
const GREY = "999999";
const PAGE_W = 11906, PAGE_H = 16838;         // A4 in DXA
const MARGIN = 1134;                           // 2cm
const CONTENT_W = PAGE_W - 2 * MARGIN;         // 9638

const run = (text, opts = {}) => new TextRun({
  text, font: opts.font || FONT, rightToLeft: true, size: opts.size || 22, bold: !!opts.bold,
  color: opts.color || BLACK, italics: !!opts.italics
});

const para = (children, opts = {}) => new Paragraph({
  bidirectional: true,
  alignment: opts.align || AlignmentType.START,
  spacing: { after: opts.after ?? 120, before: opts.before ?? 0, line: 300 },
  numbering: opts.numbering,
  heading: opts.heading,
  border: opts.border,
  keepNext: opts.keepNext,
  children: Array.isArray(children) ? children : [children]
});

// Inline markup: **bold** segments
function rich(text, base = {}) {
  const parts = String(text).split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  return parts.map(p => p.startsWith("**") && p.endsWith("**")
    ? run(p.slice(2, -2), { ...base, bold: true })
    : run(p, base));
}

const HEAVY = FONT;   // ExtraBold cut is not installed on client machines; bold is the safe heaviest
const H1 = (t) => para([run(t, { size: 36, bold: true })], { align: AlignmentType.CENTER, after: 240, before: 120, keepNext: true });
const H2 = (t) => para([run(t, { size: 28, bold: true, color: YELLOW, font: HEAVY })], { before: 320, after: 120, keepNext: true });
const H3 = (t) => para([run(t, { size: 24, bold: true, color: YELLOW, font: HEAVY })], { before: 200, after: 80, keepNext: true });
const P  = (t, o) => para(rich(t, o), { after: 120 });
const Small = (t) => para([run(t, { size: 18, color: GREY })], { after: 80 });

let numInstance = 0;
const numbered = (items) => { numInstance++; return items.map(t => new Paragraph({
  bidirectional: true, alignment: AlignmentType.START, spacing: { after: 80, line: 300 },
  numbering: { reference: "num", level: 0, instance: numInstance },
  children: rich(t)
})); };
const bullets = (items) => items.map(t => new Paragraph({
  bidirectional: true, alignment: AlignmentType.START, spacing: { after: 80, line: 300 },
  numbering: { reference: "bul", level: 0 },
  children: rich(t)
}));

const border = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function table(t) {
  // t = { columns: ["מרכיב","כמות","מחיר","סה״כ"], rows: [[...]], total: ["סה״כ", "", "", "11,300 ש״ח"] }
  const n = t.columns.length;
  const widths = t.widths || (() => { const w = []; const first = Math.round(CONTENT_W * 0.46); w.push(first);
    const rest = Math.floor((CONTENT_W - first) / (n - 1)); for (let i = 1; i < n; i++) w.push(rest);
    w[n - 1] += CONTENT_W - w.reduce((a, b) => a + b, 0); return w; })();
  const cell = (text, i, o = {}) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({ bidirectional: true, alignment: AlignmentType.START,
      spacing: { after: 0 }, children: [run(text, { bold: !!o.bold, size: 21 })] })]
  });
  const rows = [];
  rows.push(new TableRow({ tableHeader: true, children: t.columns.map((c, i) => cell(c, i, { bold: true, fill: YELLOW })) }));
  t.rows.forEach(r => rows.push(new TableRow({ children: r.map((c, i) => cell(c, i)) })));
  if (t.total) rows.push(new TableRow({ children: t.total.map((c, i) => cell(c, i, { bold: true, fill: "F2F2F2" })) }));
  return new Table({ visuallyRightToLeft: true, width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, rows });
}

// ---- block renderer ----
function renderBlocks(blocks) {
  const out = [];
  for (const b of blocks) {
    if (typeof b === "string") { out.push(P(b)); continue; }
    switch (b.type) {
      case "h1": out.push(H1(b.text)); break;
      case "h2": out.push(H2(b.text)); break;
      case "h3": out.push(H3(b.text)); break;
      case "p": out.push(P(b.text)); break;
      case "small": out.push(Small(b.text)); break;
      case "numbered": out.push(...numbered(b.items)); break;
      case "bullets": out.push(...bullets(b.items)); break;
      case "table": out.push(table(b)); out.push(para([run("")], { after: 60 })); break;
      case "pagebreak": out.push(new Paragraph({ children: [new PageBreak()] })); break;
      case "spacer": out.push(para([run("")], { after: b.size || 120 })); break;
      default: throw new Error("unknown block type " + b.type);
    }
  }
  return out;
}

// ---- header / footer ----
const LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAlgAAAEUCAMAAADAyrhkAAAA/1BMVEXo6OiioqLh4eHp2aXnqx4kJCSgoKD+1l5fTx/Z2NhlZWWxqpVQUFDb29ssJA1ZWVmvjSOkpKSUjHRrZlaScxtFORbFu54tLS2DfGg5OTn/3jPStFt7fYR+gIYAAAAAAAD+8cz7+/v+yzIAAAACAgIZGRj77cgDAwPJycm2trbo6Oj+1zXX19epqaknJyd3d3Q3NzVXVlSHh4dISEeYl5d7e3v4+Pg4ODhXV1dnZ2cnJydoaGgYGBgVFRX86bLY2NgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArIrD2AAAAQHRSTlOgpGL//2kY///Pi/YVIv9t/1///////yz/gv////8B///9/y5P//9l/v72//3+//7///7//gfQCQf+DgkSLf/SmtQfGQAARBRJREFUeNrtnQd72sjasMF0g8E12d3zfWAMoghUQGCq4f//q1dTJE15RhpREpxozrl2s4Sum6eXXDs72bnCyWVfQXYysLKTgfVLzxqdxrrR8P9Jzyq7thlYJ57Varvc7fb7Pfi3/u1bnzTVX2cnAws6O184Rf91F5xcLvhTmQGwsdpmdGVgJUuqJYWqXPZR+vRPFzjodh+wRoMwtW2sMrgysGKsqfIWM/XgI9XVOp++GHsij25kkisDCzpLLKl2d3eaTLF03d1he36bsZWBJZhVSySpfDOKBWY6nc4dx7EN4Xj+jY7/lwJdTwiufaYUM7BCu2rd8Ing1N9ovPHcTsJxvdl4bHKSq5HJrQysECv/Hw8RVeZ47tod7WPYmzlDVw67jPuMrb8drN0aq8AAqpEz65xyZs7I5NnK0PqbwSJYBVRZM6Nz+jHcxTRi62eG1t8L1qqBdCBlQUGVbdvuRjj+bQpd6W7Ggb2FzC0gML/eoQTRqsGfjJU/CaxdhJXpuJBpPp8vxv4ZCQfdtpjPNy5k3tuh3Mo9iWitl0v4rWxXWfrxTwFrvQ6xGnmCsDJmm4U1xgChYwknuH3k/42zcSVJ51omiBYOlD08lOWD7f0MmD8BLF900EjoyOMljofE1CjgaRyeQFZh1KK/HCHp5Qmq0XBGEVoBMSvkJXyqQqw+Whkx3x6sXSMQVyPOC7Q3i0BIBTQF8spB1tWciqxRIM5C9BaOIPY8am3lyhSZVfsuLnzv3y8j67uDhcSVLK0QVaMAKv8P1nzueGCY1PZcbz6fW0SyBRRazswG0WogZFbtXEJmKCPrm4O1KrcfPonJzigvbz6OoLIcx7WN5Niou3Esi4VrwelVwyLI+HruJ0HZXICH3K/V62XQfF+wVoG4siJyjA3Cg8Ixn7mp4lce1Z/k4ZbjyVLLF1pP6N+OClZ74f9trfeVkfVtwdq1y1gnTSPjynasUYiFe0qQ1HYdqkax2Nowz+ERD/EJvWhcVB8JraeMrG8L1rpdxmpwYUQeXOjrOXGiSh0WpX8/C9kaWQxahhNmtmMFnw9gLgPru4IVmFcOowQpVnMPNNM9jyg6kxQ9YAD9G0HGQqcSPV0EqUtjD16s1EPKsNTLzKxvCdaSmtDhRfdCrGzJdEI1V2ofbjr1LXzDkNkKQ1/MUxKhhf+7WZUPyWMjsPoZWd8RLMrVOLjixpwy4NiCnLIszepRy5KElxeKrUgh2qMArOJ/E/H895aB9a3BKhOuLFFc8dLK24wEeI417hxFuEYLQcm5TpDzmbGaDr/M6+RDPBMGrH4G1rcDaydw5WCLacRh5S6YilDzmM/nS6XSsM8f/6aW/zcme1drw9FpBx7BItS6GwJWVQXWIgRrv1rvlussd/hNwBK4Miziv3msfTRimLovDUWixNO6zzHia8wV3oRobUJR6MESa1LEf+2/dq3fZ3RhVj3/LcBat8ssVzNqXBl8jBKfmg8V4iaJKwJXK3+Mqpq5FNGcOpGcLKuqwBp3u/k+NrJQUwcuQ82E1u2DtVrjOEPIFbnmMxkrX1QN+0NOWPXgw+jGfK4WOIusSvQWI15oIUn5AYNlm1RiPX1maenvA9a63Ja5ikRJgJUvq4ZDhqpe8gklV5hjXjBoUX04N+LAesXWXRdLLDZT/XmXkXXrYBGHcMxyNXK42CTCiqOK50d4PpCukC0GLZeYclGgv1MUdeEEx7Fs/3H3+UDujUgErdXOfMSbBotwZRqsvAq58qYMViJVCU8ssFXK10S0DOJ8Rk7Co2RkNdHN8yj0unGD9PVnI4s+3DJYqzVOELosV8GFNhYSVppQyXShB+ep2IpqJzyiDoMXlN1Cg5GaPlasHM2yhzcN1m6Frvac5crlig/MfIhVSqgEuBiNGGlalyOrOQGdQiI2x2wafJplD28brCVrYHkcVySJJ2J14uuEbLUQrUxQy5iPrJCsNxGsR3yzhJX/VtFby0TWzYK12uFIgxHERUWuzHsOq7NeCz9DqV+LHAVyGLIMwXonTqEnYeXfc0p0YUbWbYK1xNXm1HRZsP4gDrTnS5Sr87EK0GoFr2cUXw2GLGJ3iWBh233DYtV8fA1iprUse3irYO3aD5Ei3LD+4ChQg5fDCp0vXIaMOan+Nyk+NilZKHNoyG4hAYuNzRc/JsRT3KCIQ1bwcKNgrXFo1KYld+jqslzdM1xd6AX3SEJOAw9wMpkQtOa+xMKykk/qUNudwWoyCWJbBg6aZiLrJsEiuWcacreieLtxLa627UboghKt56P15r8gem3UxhgHFsEqtOgDXZiRdXtgLZeRwHIYr3/OcNW75KVrPNwFmjB0ALHUsscF/+VtEaxHGaswuDXParRuFSwScydWlY3Tdow/eA2ufNXra0JTjIUitOz/T14fSOiIWNEaLTfThbcKFrawzMgjpAFxm+fqkq/YQMGNsZwWxGj9v/HIg8HisAqCEChGesx04Q2CtcNVWFYYGh1twl6rwB+8NFdYROJ6nKaYbUZozS3hZiSajFceq9D0mmeF8LcJFolh2aLAsq7Gla97c0G6G6jpQ2hx8QYEkPH6MZnA9cqZLrxNsFbLB05g0SgkypWY1+GK5CXHcIEMQavKB7Kqjx/A/QKj3szAukWwSJbQDQXWIopg3V+JK6x750qwEFqC0gPvxRQsm5kuvDmw1lGs0o5yhM71FGF7iYMNqn6cVKcZNIW1MpF1O2Btt9v1ell+CMUHimFZkeVOEoSXlwRMsOHxPLCi4HsuA+sWwNqt17t1MCs2NKRx0H3DCiwawPrarvz/XcyoW4VeqHGevGJ0YRZw+L1grVaN5TpoxiujtXB3SC1No3YvO7CHuyZuRs23WnQSNtpreYmtlsSo86CCvtRgfYQBh0wX/jawVr7eo398QCsrmWaXOTXdLRp0d4XZn7nc01M5QPPcof/LSEa+Ts4VWWGTReYX/hawfKbo7kq8ulKcrIANdpQBpgNHp/BmS7TWkhhn53Qg4+J6K8YnTAPWY/B+MyPr14O1IwtRV9DuSnOKFsPRQneLBkcNvFwCbYWThxT5wgvLrpONrhUONiygsPvJfqHlf5DMyPq1YK2xkb6ToDItCw1HM9jGUWvsSLPO0FQ1yxL3DqI9OL7RtT0ZLGJiSQGrdOIKPTqo9svA+pVg7ZZbbFGxUJnmyNkY8ji0DmobnamnHzsOu3eQ7u86YcsSI7E6RvW1eCJc/qOKRTqPjUqsTBf+GrBWuwahittdqdzeZUc+YcwMUWvMtrffnSK3cBgrmjXarD4+FlPBhSXV42s1+iTBTIcMrKuDtd7h3ZXM8ubxfBOLjReaWEmr4Zillp85vHcwldzCYPHDRo3m6+OHluhC9yk+vr5xRfAo6nafgfULwFrv2uzuyq652CQuRJ2HwQa9pZbCTsu0ulAeY+uLruIkDi4CVbUpLXwySRFpBtaVwUKLdhvRot3RXGvYPwLLSzG33Z1bEVt7JLb0ycLEb4AnbWLRpTgQVKjTtRsKrAysK4KFsFqGKnDkaO5utlEbQ6p1E3hOtxmoxIa+2NpvGznimDrQWPjH2FIGSS0HfdoZWFcFC62RDC0r09EnxdCw3aGHhToRLx7UlVr7XGT6Ce5EMz7KzjgSkZua7w8zTXhVsNbMot2xl2Yzicd0E6bcZRIMu/2805davacaEwBZRNPgmzHR+LfgFY3NnImsmVeqG8vACk5jG2I1XWhLHzKf30FOYee57p/n50ff70JHl8xgzTPa8qyXp/YpeMrVmLgYDc0aMVxNArI27Bzw/P0w4+qqYO2W7d2dOEBKrcOM5svLy49CofDun24TJXTmnVq3UKhUwrH9xSIKQlabyYBFq3jbemh9oaEzrVYwUnmRzJVEllmr3/87vHw7bQYWb12Fi3aTsDLc6qI+GAzeyUF/KqCEzshrFipSKBI5+ZXK4/MsYeeX4RC0ak96+jCY8JdnuGomZKV5ssjUymgITkbTFcBahXvB47GyX+qFQ0TUAPM1eK+jhM5oBtRHVSqFA31AofD8Ymiitdciq1fiuDI0qh0ksi47rSQDS1CDwebK6SwOqufCICDqnTkqsCqFAiPWyEMLz7MY642ihYIPew2yvvA0I115BcqsTFxdEazVMhBXytC5MasXQsUnngCsFxasSijZ3gW43n8sqklS666nQdZXu/EZL6+AXsIPgx0skXF1RbB8NUjElSoOZTsEEoApBqzx2I1GNFYK74r7kicpPFfjlgr6DmKyPtyThZsx8mryCmw/CSaEYE+0lHF1NbB27SUJZMPGlev8UEkqEaww6F05vMcd8lyFuoItly55/umLpJizXSVyVYSKTDmyjllD4bXA2lE1CIorXwPGiaoQrGcWrASsAriw3IJl5GZKjfiv2GBuLsFuR4P8oE1gLFm5XhZxvwpYSzzlUZHSfR4kU4UV22NnHIClhVXE1g+wwItOhn/q9fZqrnB/kKG22wlAxRiypuQ1MrAuD1aZcGUCdQlefaBFFQFrMR7XO68TfaxCtA7/QGiR4fA59TzjMlKEph3nDzZVzdIBWahgppbZWFcAi8qrsXxpZz+wjtM8xSa5voX3tAfHweoulNXGNnxLoQ7JPN1NR92tEySdizFkoRK/VgbWxcEKuJLlxQ9dYUVFVuXRP5X0XL1TfCG0SOKl1fuCRW04wAFuLwzjCtWYQTO4ozATWZcGi3IlBa+qKbEiYfXTqIpDi6x2yrWhxREYrKmaK6ZKJkaeGSSWlYF1UbCINpHM9uZzeqxCPM5CS1LIBiELkigMWGBlH1PXB4qs/x4zsK4EloKrl8F5iJyM1uB9IJvxcxVZywgsnxzVdnGVyJp8EHkWgpWRdTGw1iBXRuH3YEWl1sEDyar9lK48Mxqk03mTN18W43ZET4pv0Y6mDKyLgrUql3Oy3e6Lq9+FFUVrYUB5aZmsLe7VGYWbcSZQqKFDx8CJq3Uema1yZikD65Jgkbi1wFX9/bdyhdEqiPUVNiarIVrwjSX3AV6hSdseHR7CiaxADdK0ZD6rHr0kWGRLyYi/gvXTjPZLCy3RiMd9fzVxqeB+v+MyUbw6NIIOZ1sUWZNik3UNzKyh8JJg7fCo4y6XqXv77eIqNOKrgMySNuzuaU0GsAmaCiwnHOdcFf8KhcnwdLhW1p9zSbDIbPbZDXJFhNYzSJZw/bftJ1xLHZplgXcYWO5m9OOhIotRgziwb5ayhsJLglVm13XdGldEHTY1ZNa2/RM36kxtvveLxkZJIpsVWf89GlzstV4aZgLrgmCRjfOmcaNcYbQGfEHNjGZ3eAS+er08X5uBovBMuoYTWZMoGB/tqc5M9wuCVZYU4Y1xhb1DT45nSfli1F6I1WE45+bN56cZ7VwJRVZzMikGXNm4+/rYyri6LFirh4dw0jH9pm+NK/x+HJmskkTWV+8rxxX+NB+roYw7DuvhL+g1VIMzU9gnnGF0GbCW0fKsMH51a1xh75AnawQmd9AG+ieoEQQZUf8O/xV+QoEaNDOuLg/WFllY7Lf93L09rrDQcqSMdF40s3ATGBnkYNmCgEPw5IXfkIfVYK2UcXVxsCSBVX2/Sa4kstwuXJbXC2x4pg7WpcvHhiUzrF8OnQBWDWZcXQ4s3DzB5Z1vFCyRrBmWNQBZoQ0/pwi5oxCffJchywqmy7Bc7Vfr3XK5XGU0nQXWkpk6TGyOm+VKImsMRbOoOizVwk5uG5tRNYpPDc8NsFEz7JRVg4G4CoG64L6fvxGsMrNfEAusw/sNH54so6tqq0GU0NkTIzLK7ViiYqlFZxOOTEgN+izt7vApt7VGRmRgqSRWOceZ7v/csMDCZL0JZhbcVoMHOTDj6I+lUCyxN5stkatyOBYzd7fOyDrHK/zksjmFWwerwDp7FugZBjLrK5hGaJLhj2SMTCDL0ENLvNW+bd+xCzc+GxlZp4K1ZtumUKS68H7bBw3fYpShCYVJI7RarVwun2+VOK6Cm/Olfp+bhtUgBRLWzLbpmNunP5SsdblcXq5XVwRrybRN4RjWbQssOhmC7wnL9dRkBfPYWH5QPAK42Tfb79jwF3IlPxt/YvxhGewDXF0TLD7Y8L+bB8t/g54QgFf2mCJ1WCr1xSF95OZSnx+ytsXC2+JDsLne15/HFd4weXf3sNcn6wRV+MlWJDdvXROSw5hZqP/h2IshizmK28lN+3JZrB1CivbpxmKm62U6HSadXTQJG8+lvhpYOTatVv0OYA3efwj2e0sdM4ew4m7nvomxlEfMga7B78OK6rDG6eJqx6zZyjX2oulVXjYuAhZKFEbh0dfC4FuQ9cIHs47x1XlqccZesodwqEh0SDTjZrjyJdXDQy73UPb/uD8RzAfk95pzz/DGpK9cNr3ajf25YDWwXbH4TiYW8QybfH65dXaaDxub0vAmXxzeUGI6HOHZzd2dJvCIf2LSVBfKSISNdEgY4m3xdw9teVB1LvVPgI82VL6FiYVHu6URWRqWx8OdVFCDoTXPeuZlebdabbfr3SWyQ8vyQxRly23TR0KWK/Qpu2OXV/XUdYmWkKDFDduzwNrulmUu2vC/9+9y3MuKLCywXBVYpz5zmdNiZ3O1xLMwZ7PFlIzOTEnWjqjBBfebpH10+/bDHRsafhLISgPWCi8Nf/qWYHFhUvt8kYXjxOMOBNbw9GfeoSWPn/7J3S3Pj7TumKzu5oTo7ZpwNROqJUmsZkdML8d2bbJw5o5/cn2wdshSo8sto7hQ8f3bkOXxsazSWWDtQAvrTLCohCDn7lyufO/iM7pUuK0ofiirTLnEVQgW2fnosFVGd9yT64K1XKNa90CnRsKxePiOIssjrfGnk7XELqFxWbCWmITu1LJGZK7XeTKrzAtVXEDb0A/f7gB5hcDCc6K/xGJPpytE8PTA2mFLLfg1sdsnioVvKbLIdvAzwEJfLDTP9yywkIQghrJrpZUvEFh33IJiRJZ+LISEGUbC9B6b2hC9hsicJUTwcppYhfudTYt7rWLl/TuKLAeLrJON7PW6LA4YOB8sZhQc7ZZtnO9eeELRf+2nXvh2RRxKVx4KhYagfKGntjqykvxKARZaRx9itRCXWxYnh+9Dls2Z77l+6dQLhxmwOhcGa/XJXknr3LyjBBbOOOV6P3V+OQ9tQA/icjY8BOVL/l25JFOmDVa0La47llfRVycf30hkPcvm+2lk4WyODYNFtqCcFMzkxvegaNtZ/oUMFi76z+vQWpY6/IJvLY8T8U/A72rKxYZzSd9ggJVpQV8kGtDybays90GTq5452XzHQzKnHQVYMU/7UN41Gg0wt0aupcNbLU/n5B2xHejKg6RbyR9azZWJCh2/UOXjDNhh9KQJ1qrRLufo2nBwUSAZlfF9HMMXznw/OZSFZcEsNVjrKM6s4oD98XqqWldtsSqQSq/+Z6KcXsJcjfF0HVw5VANc4hmZTK4B1i4QV6rdls0Psvvm9hg6FCoVaS8dV+SwOF0X4ljDVLl6TAXWrr2l0c+H9nYPX02D98DO8lwbq7I4Go9InaT6CzJl3ZF2QQYLGrFPKFuYNvfRcxoZTNVy+epHsKzrxrAKFwBLGwlsXimcdt2wolBsZlRLrCUT/cwBMksCC8WMzkoPrJABbgP2d7yG3ZUfugA4G9KhhAXWE7g0QhOs4ItQLSFk58FWTpUrhcLl9eiB3StdUelC+2Qja73bf4LBUflnK3+d0/FYEfzcSWCNzwbrDhBZM/W2DvKopdAtI3OF0nqbBGGdi9Gzn6qtXhirCtI26JxkZg0K9JEfSGcNrsUVTxanC0dBrO8kgcVKcVcDLJJfGaNojTEHg58QWOZ52xC3CBEHqOypfalFFjYgzQ4g6EhHJQq7P0H59wWnAnKx9ls0605Y7szuJxlUUoss4eKn2SWXeMSn5rVhUzCyTpAIuDibi2hrgIVpHEe//pp0FwisU5V1ANb+CQhyojip2ivAs89M6DFHyhXJ58jr3qZcCC8X52+Gowy413j5gbcQvnNoHc7BSr7+Z9hX0jMPYF3onbitZCfWNcy7yWCtH7hy0zHg9F8BLNT2CFDixsQcdkKfu8gVAmsrB8jIs+b7SWBReTUHqKqDuy3T7O0KLz1Sgv6hOrFynj48KAQWhyyrC1EY+hQja1fmYw3+83iJYAlBqg3nmavAGp0Nli+zcoADOwemk7MX3onlqr0FS9HmGmCVFVy5eBP9mbPcK/KS+sHB9+P+O4cs//GE7QMgDNnnLRhspPgEI2uFZqKwprvDYRaA1YsNUuH0x1cSWJvzwdq3f0Ie3ggAm9YHfwLXHWXES9FYTFQyMwU0YbeUANYa3hY3uwBVlCvAXEdsHc7TfwdQE3KexeCdS8aZ6Uce45LkOZ9/m8nBp15sLMF/kOTwyWB554Plk9VS5PwAZbgulxUCjgSwyOfao8F7U0ATHofxYO3W2DEWuMJ9XoNLmNZKtVkpnCcFKzpgPXMBv/TWO/a0hDJnD6griQ9SjXTAml0ArDZ24SSLaA56hqS+wpC9PZYrn9UyTB+rCSGwyux29yC6UJeF1eGQXsJU/qtcIbV4+NAHq35etd9SrGuYCr9BjEz/dsBq/0SJPYkWyDPkFqLxiWt2PPS+vQSCGEgTDmPB2rHb3YOlXgdeWuGMyWTim9+HlJLlOhnrSmilJ4LFGFmnrITDQ5xcvhaVC0KOQbCwnDMEg4R/5WuB1f5CWxLGycqQWECS1jSjABa5ewPKQWJNGAvW6qEsbcl55rAaFHymJsXih/+/Saq184fKlZoQD6FbmQCWfwy2LqWWEqydqAZG4k5sPbBGsha+Glj7XgOIk0qeISntswAA6xxXJLntqTShEqwGVoRcvL3QZbTgoPLhU/VaJV9B9fFjcmo+50Aj975/eHZ49BCslAa8QuH9eaz1fkxpvZeFugZfYB2PGmDtBbA2vxAsambNEjzDMjAxgNRwD/m9LrgYLUkTSmARecg9PzdZDUmrxyqfMjwhbH6ocPEmuRLhYoF3IfQ6eP+HtUqPqd1CbtQOeor7mgAWRMNvBav9s52XpwEIypBMETLAShl+X9B2/yndUdKEEljEcncVXB18cRUu/ugYYZHDIS1WaBMScxQhiDMMeZXAYmPvVmq3cCekCW1U+XbkDNLRWWA5VwGr/YVGFY7ilCHpcPdiuOqxn8WSU9TCQtAcFBp1FLO2Cx+TD7quz6CH7MxKZTshOn1t+vhYfUOn+vhYDNi6BFqFWAuLjb3PUlfO4ECMzSf0hyZ3kwM+51oQR78YrDYqzZPKnxjPEE84ke5gQVxhnbaQCeQ1oQDWGtWGcXFlZpkJsq7odlEjOqTgr5IOq+Jrk5WlRvO1iNG6SCyiEpeDZHt13LQSqyyUpaNL8y9at2MnRTUBsO51wNJNZvoXbrlcb/2zJ2e93vn/yYAFRLPcsIBmv9p+yrRIgQZQalM/SNCEAlhEYDFhmReeq0cRK4pWdVLQvuqIqg4j9iJbbXKhMgcmyS09nQhWK41bWBayrxv8xec5sGa6YOWFJqGTJdZ6t1yr8k/rYHjRlzAaIQqToudvrIBYF8wV+SiunIAWFoLmBBsixwksZqtXyJUhHUzWQe+STz6IPyk9Q4TW6Z7hQSghBQQgKLE0wcL7ik1eYN0DYOVPAmsNgZVPenvr8g4jj+YJfTInhwbP46YNutUAN0CMYGWI62BE497uigEsxsp3peD8veAG5aSPz364HzpcYS60GlcP/nM0FU8S7fs+SWgNaIlE+N/owB0VVa54LQVYS+HSU1coz33PM7BhUQTLMCWwtu2VDFYuVmKt8eSz8kOOGbnHnVzuroHWqmNl+FPOLwfK8E6u6OQqsEThIwo3/76CJuTBKi95geXpcWVgMysZhwJdUWoYqmfBC05Pqc0KtV+i/c902qcEa40HrBicbXs/LCGw+KS0BlgdCKylAqye6v20wykt9IwtcsYsXHdkYUYPr9jYAJ4h7pffyMLsOAQWE6/QIICRJNxqcWCt0UdjPEljwJpGj/FMVBMvaYVEKgz1Cfd9F0432CuJYIkSS9d6X7JFoDTWUCISa85dEYAGECz+Xtv2Lg1YCKtlQJU5dRwvslfxl+w5znRK0HpCaKHtefIgExQmRaJsJGeeTYgrOfcQ6mz+nebEihDGOI22LyEXvhkHhcbcmcKkaMRzRURf8QSyCpAbiCL7hziw7LRg8caFE9i2vwOscjSlZTr3FJ0dvs7BI9e6uQZaf43NrKmsDE054i5UyohgzaXGMNHEEiQWV0YfbV/yr9vkLRYKLLLilRXmykg6/l1OIAvqyzl8gF0eDFhBx7geWOUHvpsQFWKR/Sf8F22iFoh+MljHM8BahmMap47diT8e3ZiByPqqSVDgjcZiIN1RckWcV2m2kVmKAWu14vvJwv2WA5+r1wRhkzSDTZOr08gaQJH2CqwZTwZLHGa7oYZF/0usm+lqgdUVjBIZLMOETUBfC67JHuux19E4s2lAFhTNGsMBrrxikXpZnjhno1/JUA3WTojqhy6hb2B9aCix2MncFU2u0DO9pbXgD78ArDX+2RncBbkf0skrU/7XngyWIVevymD5WhUucyadeebc6OgdY0S7VLEyFFq7bFPiypQqGvjfl2Co2dDXmBPrvAwmhhUpwmoSWP7DmpU4A6upyRV6pmpasirXB0vsJnTpIIMeAsvUAsvjewaHp4FVJlpQH6to4B4lSzCobE92CGsqrghYySYWB9YDF0JzGIFVTOQKfcpHZenyAZFpGNpkPaYcNgIZ78k2Virjvbzm04Rzqi16X8I3DTbZS9HPOQzWIgmsNW10WaTBKhrlCMYcgAzhv5BDSJLVW0nmLeLBWgkreQNNSC13HUFTUEYvHzupwOoU08XgB1C4IdErTBNuWArt6jad6NOTwHK0wFrAYJlJYJH6XtPrpD0cWXb8HU2F4Y4p2UuZwhH+LAlgeZIm1BNY2OquKERW4cNIwVVkZlVOsLIO14ljiZOLglhDr/dTAMuAwNppgjWNB4u00XTHRif9WZBF62BqR8gQqrlC0+dFx9Ig/Tk9FVhCD6yTUmAZ6v0ng8prKq7QcxVTKkMaek9OBwmR93s9sHC5kin4fuTb/ymEt0CwtuhXK4BVSg8WqZZzOqccY0QXNaISeOVTGDEOYSi5F+JDgJ9njn/TU6leRldgIRpeYbAKRSM1WM3UU2y0R9e40kTNZLB2wuQiL6wTIQ3nm4SxVtt2QwRLxG+1330mgLWEknraxw72cTyplSGi76g03IF8KVvkpwBryQ3tDaKj1CXUpAG+sJVqSq4wpJNLTnSIq27QAqvMTy6ywhlkZD7CjMveyltPAIklybWd1GAsgFU+iytMwLHXBz1DJoZCtqyrMpTS3MGgtiwOLFZiGYyFZeiC1QGNrINvYRmpwWperqb0fLDEKclGN6iY7JH+zRnnVdWkJ5Uk1lwbrPBOO2E/VvpjUpHV+1QoQ2ThKx1CFVgW+C2yKR22+4dW+B00gu5JRlZqC4u1sipXBIuWJmuAtVOkCXHHuQgW1E4hgWWnBmu1BOcqnCCylJ6hHW+4hzOcbHF8EVDVJoC1iVoJB6lMdxp8B8AaVJongUWipJeenMu2QuOSbp2ymbI4uWhKou74kVjLbaSwDgDWRhr6kwYsrFEU7pxtb+b0eHa8yMILS3o5aIioYcYb7gFYQjXWCJDQarD+IYhUSFmDtmEEgFUonsCV/2TGZedmAV06iy79Sno6sYaNotkJgzVKmOWB77Q4B6wlyoGDNre3GHPVfdPxzFDHqPJYZP2EnsuKi7gzbadjsSIwDVi0OeeD1GEpSbLxMYLqdRvyCaungUXD75VLg/WP3LDa00gTTnlll48FKw+CNeXButcAKxdesjW8bsyYTwOgTNMM/zjeqBzDY7+kEFlOMKwo7gtBm5vH4kzb9GDF+YQd23Pm4xE6C1QRpHILK28dIY4q/elX6sIo2kBb7Pt6Y88XnMowo2anfUMoqXSBCfIQWGLZ8TJOYq3hnQUzzNK0ls/nS8NSqXSfzx8JXvAI9RGdGYHnOdhKA0vNlbym0YODgaeC1fGc0dQ/GKzp1Bwt0AVrFmQJ8T/m8f7bn20WtH52M7Pj3MVAF17afOdmN9zrgIWX3PD9EnWmEhdphxHPgxTIWq2EmVLAHC0IrDCK2niABBaaMNTN32P9hUvD8GnljxgtIO8TOKO4accSFWGSgQWC5Z4Clio62nEtyhTCispgywAk1uD9MXx8x9gsMI3BGS1marSu4hey/aq2ZuB9+SBMtR5FsQYM1loCS5523BCoceU5WktpQY8ZFkpsQYGFuKr/OxySGmL2tI7wgP5ATcs7AFwdRUjAmiXXNvBgsS2FsWB1NiMGq0DDIyp/iNb7oBAEGzrGbEQsgfCxPmVjJVqBkXVRXcg6hY4mWDshTWgL3ZmAxJILZ0SwgglZPfXrsGA1FOtLcINWSBWOUbFoWUDGhoLVFmbVz5kx7urTkMEaw8Ogcvx4gSkHFopiAT6dPZewwpXTG65xOrDdm5QrhBU1LwW0OrHB94uKLNYpnHehrDzQ9CkIi0UUdY8quu2EiiwJLHkqzU4aRhyBBS2Z8IhRNIyYig66KQ+QRQoMe6Qmg3vCBe2mj/82GmLQLoiPSg/L8Z9LBksOj9pWyIXJeLkwWIMCtd0dXLQf3plFa2HDZDWvAZbLNzglg4U1lFjUPhTH+rgJQ2lhidXnW4xFsELRt28Da3HCiWgSVoQtHAadSyGFWjSr3RZqpYeJBudarni3utN+MlgjLgcNgqXgSiWxKkRijbumyTdTMmhZoDoMwRpcJe6O24CTbfc1ShOyVcWe2ETXEHvON3K8Ybt+ACRWqf/1td2u1uvdbrdu7JZ3Kom1BTThhplcJX0GdEsJZ5sNESwzAosvwq8nCyxo75kFr/4UwAr7rB0VWCqu8Ggvo6BwCsddU2rTNSOyxl4MWBc0slgTK6rG6iWmCecdqda9pwZLLJxZltFsBUH1QItyGgJYI9RxUcISS168OWLSLz1ovgyJVi1EsIig/BKXWqPGjeuARQr9XDalUyBgcVfbmcrmFQWl02kqog0QV6w+nE7tGLAKVzGxHL0U9E4utzL5nzYEFhtv2JHZCvyeFFxv7x80YyE8oipEgXKcNd7LiRT/RabxYXKSEzQliVUKwZqLjY7JHrLcSpEI1h79YEJ5W6UJHQGszkbJFXpBBVgwV6w69LWhMt5wSSNrYPAVIonfpbTgZMNG3SOwPD7nZkZ3WLUfgoZlTxxeKh9P6vnL+c+0b0izzuzwffRiZkRKu0ks6sOJYM262mBNAeUaB1Z7tfyMvkODAYslazxV6EEMliGD9YwdDsXQiliyLg8Wa2KhjGutn/hdYtnviBpImEQmFlUumIqscrhSW2zTSQbLwTfVvvAet7Gc9kvSXXv0uGuDNUoEq8HOMMKT/CSwOmGgYSp/J/4nn0lOIdI9lhosxs5yOr8ArH+A8GgvKU3Iyn5Xnn8hg2VFbuEWj8ozLTxbgVNl6skozJ3sOeL/JzLDHbnnKsmLw2DNrguW0U0Ca7Uts83WyL87kBx09LFdNVcm+m7/SQsWS5Yr2VgfF0/quFwUq5ToX4sVfiSOKH2RnzxYTBUf/rWOTml+CH4AeIaVvCDQkYYzKsAa64KlGyxOCxa2UqPXmhGwPtge6EhgQaSAYSyUQokDK0YZ0lYdYcvShfI5dEyPlia0hdCS/CgBLDsqXdiXP+W1EKkOXugkqTQahY1/+1sILMZ458d6la4DFhnHPOJmGNHIe0eysKYQIjYKq0JWTSxYDFlCCD70Ci8H1osUbIj/KqUFJ4AmBMCKiq22wMDOtAex3ILAaiWDBcTJzSjcYMhjKS8OFs60cr8tX/ogicUWkHoxFhZisiqFRzXAMlUi6+JgsetO5l0+k6wZtlmIsa/VarvfSxLLDMLqW2A2S9rjv9WWDNZcwyhqQHFycxiCZacGK72NtUPJqOjNG0QXVriymThNaCIl+gKD5cWCpRJZnUurQnbzl54mlNKEkSbs7Vfr3XLZCCSWKa4Jb5XQnfbAJNi0xwPB8jSmEcrZR9JC9AvBIt2q0c8OZfgGJN7A6MJYTegCtQ1aYEUia9GBqhs+Jpc33cnmr2GiJhTShFgTcinXlX+w12MI/HXN2lOv/fMCYM1AsAwNsLaNTyFAGmIgJqGvBNaKOMXhvZ+7daQLybCNsObdHcX4hL4l4iqyvglgMVlDFwJrUjxcOp1Dd4Ym/uKXwnKPOfGd+q1WuXyXQyf8EIZUgNftfj610brvzUXAcqXOBzPh/e/kqjw7UuWfvwKsHT8u+QU3ottkJEI4wqgzi/MJ3bCzR7icbzgWqSeyWF3Y6VBFOHksXFpgoSBWLTGZj0dnLgRJdPRPDfgQfJDKc8YjMkjvScqqpD2orvAeWhthJmw/WMuDsx0WrFHn6mCt0Yye8PO7dK9fHU9xCXVhDFhYYHUKqhkc1im6MDSxJs3nC6x25QTWphtMTUtlotjiG/fPiHRdAeoO1zU+tXNnuoW4iqEk9y47SWs8l0DV6TyIKvSEhTjngJVXgkV6bB22FgvFfJrYeo/Gjy4CsGCB5UCXH1/PmbYutOVmio+J8c9FwHI5yWMmCqzlWlhwQlMs6N06M/8khqdQfxmuihrbJzFluK43Ir+BGrTKshYLVnktl8kHFV49NCji+mDhWIPJtxQO3h3f1Cp8MCJrobLdsUtoFJSCIkkXToHoe5jQuYzE4lzCcJlHPFhimrCzWbj+SReCymOyzLl3WqChS99qTV75Rre0xiYNppJhVhuGzRTjjhgg1QLL1Adry3bXN8NL4fiyq8JYWQFYUr0M/l3XFVcfFRRo60JPTuj4XF9EFXZkgRVvoSzLn9JY9PSRAl+qtPCHN62Fx3cvR2ezIMcyLUlCmrhMNC+Hwxx5bYqoCEWLPxwMg8FyOmKu8MISi9eEzFVsukRkfZCsqVJioffvxtnM2gGHeUeseL8MWGzQXVNgSWnC0wzvWr/UL+XNMDnBlod2wUw+J7HqtKr9SX4vyOIr9b7Aj7DCUwGg7W/3EVguD1ZOM/LO/9jiwbpjPnQhuhiHav2d+GXxYLk4A6QQWO+OukhE1IUsWKEmfD0fLM5y1xNYUoXfSYe0TKDOBsxWHFgm7qITwQqq2vuAc4m2R7QAslZoxQ7iagq0UpAf1F6sIdOUWHInkRcD1pKtl3ljL8d7gYSy/sPjFS0QLGxgKRUhvaaawfcQrMh0R4r4XLC4UndsuegIrAflFKkTwCJdM7ZQQnp/f5+nB0mm4ZF3FsKWrD6uMt4A06zwXuftNtget14u8Y45PFhZSn5He1C3osN7MlhzxQNDsCx+GkhAFh1rjc2sjQnaWLFcBVWbI61ENANWILA+PowLSKw3fraVhsBawrMSUieQaxSrUi8vNomZQqspNtGFPrLgFwBUGVOyck8/cYmMT9dqhd86rViViyrMsDRDAku1rEAamSlmH50YsHYsWBIih4+QLI/0qU5ZcWVukrjC5o2rFSINwGIF1iMcIEslsKrCxbhPFFhymvDUkhfanvXVE5Y5zXEjH9/DLIA1Dyswel97kHMLB/hzd3c0aVm+C3fMzaFZM7VhBBZvK+mCVZbBAuO0ElgFxXR+4houcGc9i4SbwFVg38zjRJYAVrD2hL6uMbhcpAFnCWvJAqt8gXoXMpaG9tFAYImt8T0BLGar1l4Rwp8Fv3Ky/zIsNZnb4HSsoOh0K3l31wBrHQ9WQBbeK4eFVtBVj+c1JHBFdGETTzTQB+sxVIS+JpydpQl5rrDNnCyw2rvGZ/w4dJ345qIbbC9CL/YEgCV0mwpVhazHtVeF8F1rKjY/Oa4iKhbWkq1gsPpaYI0vBdYh8PvxZrnNeBq4MFPEbjORK18XNpuoSEkTLKZgBtfcnxN49x/5DMw70KrwG52H1dzky+p7gCrkuMIxABtKGWMbCliwFLDlBRN8rLmn2jHnsmmsU8HabkWwRmeAFZKFOys6hmuZ0QSml4Fo7QMFdp3OK/JJTC3jPdhZSE7zLE0ocoWShMfEYT1Qu3va4+APyw55+CmCJfbLbsW2d640BousM6qcp2zfGgTWUWcc60qc6GfqgzVQL34gO1L9n9JmTGbGuXWepoFKZFUfsR+UBNYsWngSCixQEw7CE4eVD31V/NVqDOsB0oSpD/qodUbn4rbNeWwJ6CoWLP8ZameYfUGhkBqs7lAHrK2Qjph26/1YsEZMyUwMWT5azeg6cVj5VjqsFXGpRPExSOLGgeXiEbmRvELFYADqAx4xtbiqG+KvVmO6GFDhJys6dDYOczxRQEzv2d23GKxprJ+uBIv23OwbJ6+liDaZkGeTwRprgiV3l6jBarD75OzYlSI+Wf8V3yKs2Es8eJ91VGTVfbPpjSwAikvp4CoAhivkinogV/XZy8z//0uhoIAL3TJ4kcM+NQ2u2itVmtB13Zm1WCymCckYJBvrwxLzUhgsUzKg+HEzsWARz/A0sjwhKCwNvwHWkanHzdhijUUsWBtVICvyDSeTj+IrEVi280PUfLgwpaB4tOvz4pPlgr6hGdVj8Vw9grrZl4BMwNN9qQ8EtoiGLDiGzNVRQxGSWIMAymZjwQVDTDMJL7HqPMKNXe5MsEjL/ClkeaLLch5Yjkaja47MZ2VG07sqE/xQeayiEJPtPj8XBpJBRbr23gbKnEoRGUy2BQitYJKpY/BcIUUIVXkNmoJeeqlLPoeEFclXlnQEVkNKEy6ktF6dZGJK9JhiWcs0GAvEjtGKB2uXBFb759dJZCGueMxPBmsl7BpXNrrSshkmWRgTlirU6/UfjKoR1FNVTAkJytCXQM0OO9mPBWs6tVzeH8QGljtICkwFU0ye6wdq0vvv88UF0x9JexeCUncxTTgj7/NYO6L0XgkY+pkHwOKbNXYXAKuNQ11pyUK/ihovqU8Ga4smik45ZmPAWmHlPYtaVRMc/AF0tYPClEIcWTgUhkeTM2yZaOKthYzfKssVynw3wffSjDGo5QEIabkiQawNn0jr5v/9tyThFAuWEEzYCWavAqx5fCcOIStNRSrSEUGRRPBMW3GIsz5Y+y3X+TOHJ5CG7V8r1kioKsFSuveD98NL/KORae+T9fGIvv/ZmBsyOR3PkbTzxRXDFQ6a1SHLvXBCMCcFV+u1sJuQxqwBORUEzvs9qRBP2tcAzMgVqNEDC5Olnx7H8TSBK6LQTgJLjPCNgP0uDFjb7R0jsjrP3dSxbuZi/1A92EER9UkRGWodezMfm9NarTY1FyQD0XzksPp4U2hlvrZK08iIopXJYEkVfkhg/ctRJc77lBKBQOBQHN8pT3rXAQuR9YS6hCytUjEP/aDMMO5xAbAavDgf+0o2Bqz9fs21MD2/pyMLNXlFTYmqkr8uCjpMkNR6E7+AZvXxg7WusN0OJ4vSg2Vg09vU5EqegjwPp32CU2R7dGy604lNdYj5W1ty1JcwWMJl8yku4f6zZLRwJwY1r7gPDoC10ARrxc3ZNZQB+xwTm2MM1urgfTBIA5aTDBZSYq6B0suTSfGxGoXEq4gqFquP/x6J3T5ImhmjL66OQ02ullITghkt8QuO9KifosRyQLA8vrGBvyYNweNS9Oz5ZOGxtf71jZsK4TqjLi+uhPdyUuSdzufbcEX9MWDtt9sa524Y9fdUaL3X7WSvEgU27WoRx1l9uIIz4akiOUkjRmqmsbHscdDposcVWWE84531eswU2XAu7VQDrAWXZBF+7FtxkbehuGxoJHKemKcLcI2c4dLdTSY7sJvbwHI6WKynN1KO1cqxo5S4L/StHqTkdM2sOj6F+Jzw+3M15IeeD/4g19FwCjqDRI3FLN4/8kgACvzZqn6QYoXfiAqsmMfut2txSoK0QrUhzmofy2Ct9MAi07bztCF7PJ4x34I7my0C16h2P1T8ntZArrCmt8WY5Cy7o/lstpl2cTtZHFiULNPmQkMpSjc1+UMR8cpHzJl8VDvGSyHGyGM6blCZ8chRsWU7NFKeL0E/W83x265Oc6scF5KXy63W/EY2CaxVe6sJVhv5C73+UzQ6YmpOyf/DW4750nCo+tzrE8tmCFmNaMaAqRormGN+LzmRLJ+tWRASDcFRE5RcbkDROqjRmjy+zVCKZqBn0BGBNLVmkiHrOUFnUO1+qM9Ve7cXBoGIG040xyTIwU2x4gQAS5BYtmIuApVZ/R7q/qnJkyTMY82nqs9uQ2kngGWkAMt/9Sf6qseSqrQtJwxvhsK6aG2A6y0GcSVXKY8SrQrOKscXwzA9Nwsmeu84zsZ/qzP/39Y0/OXW2HUzyV+Z3Jwz1ZpLI3Wv2NJCwpWwzCkZrHCalWKMO13I1Grlc/lc7tg95nC/z32pH1LVh30NCaxNl1m6mSCwfFuw3W/hFERfKctzbYksi7eNX+rPKDcYYlW4zOyXQ6EiwVWpHDSUKgwWXJmK1UFfW1wBg0Dkue5Ki9YTRIApVsXw1c7JNhbZfNOP2RAQLvvCjUBw9LankLCmkLS61wPLx6pMGzZq+XvlLzbH+cwkqmtyP716NzK8cUHmy/ulZoIeKhVKF/rD4aAb2+DAOt7na2AB4TF/H6sOlI3pptg8nLzKCbt84mLJowgWry6lbferlGBRtHowU31VZASskJ3rDbdFv7xyjv2SW4oHcWC1v2iAZMwUTtejRA6uV3lR2O2DU+nyn/0wSBWN5cFCJmrJtzXC7KNpmke0yjbcOqqLFQ1iOWDzcMomA5kbsSpmLGZD0oMVoMVLLibgFhs/FxZl6Eyb8d/jkjQtzjc0opHraYDl+xotbJUxNvyAjUJ1IK7ef+2BwKLGxj0+Qk5PHysfrGVX1TOa9IWLS7jnIjdiTgcGa5QOLJYt8cQ6G2Uu3+hbc1ol7+sy7rFe2EGFOibrK0kVkqhul5tSEsbR/X/NAK6Q1fVikIBmUq3w5Vd4WVFIHT69NFiReU5WatO9DW2qlPbU74RWRQsEy0wNFkRXciqZl7B2V++VcPR4yvw6Nqjwo/eVCFabZAuij19l5ATMFaqos1/qLzTiPvgFEsth41h1JVe637JQ4TcTWpmHJ9WC02wH88C9EHr3LglW2rPdc6MiN3obT8psEXvU9/QVFyBlo1lRuXfYekWKVSCumqi47/29+9IM9OU/9SujFV1AthYmvUKQfscrMeo+11kzEBotrlw4I65Gmd8KWLyEHWk5hbigSKzYGWOR1UsAi3SCWFIJTMjVQC7mbNa7SPsN7APJM3deCr/KxPJoS2j/fKjC5pwZqAmTdcudmIamy+zZe3EmlAzWfvUpgmVeCSx+e7nRpbUNGgVFC2A+RTJYeElG5C443YArlJ8zQK7CIoRBMBO3e2VN6HFhLGoC9Xhj46Rvuyxs6ZUW1sepwoYIlrSFG2sDQw2WVL18RbD4+oaFliZcL8tdYO4A+vH14tbKhfHB8LHMWNEZUHTMc0VOtfPSHfwqgdWZ6kzV1udKas6xNHa3MZUJI2AaleAWmp10YB2vBBYxJxfhOzG1FkKDg1IW2mCF30/U0ucBdVYyV/iWwvsvs7D0Ni2lAEseGmXqacL2drsV6xtm4oxQIe+z0QGr27oSWNtt6KW5pt7XWBaDX+HnKPWTwFpyYFGUiH/vggV3PziuCh2NuPzgvNkxL3Lb+GUEFtYOYp6jrqcJwWn9olsoVfXqgHV/JbBo0cHIcRZdeKUZLOQ6MFjSg2PBIv06VPkUxPZUZM7X+Vbol5iK9wNO4FQOg3NCEoPuP/wkhlp/eDmBBWxKvdfltiG29mBySrzX12AdI3m1xC8Fi5LVpdXLWnL/Ux4KHmSDdMCKRpq4BRzurIIzHZrSjWgQ1kBRzRBmnCvPEqIpuPohzLm4H15MEy5FmZNCE1J5t4mt9ROarqQ+Hhms8TXBau9JW4beOAs1WFMdsLDXzH67jn850cWU1ogj6WQMJJsL1IRcaR8eVsuPPhroqkFmpzMZR328nOneXj8IvXZpNCGULSSDIjldyAxodGVPTBrPv7kqWMiDbqFubs2qohWsCqdgo04O+Hq4LKxbxwLrWZQydaDRC17UJFaM+mQ571GiqK7buBH1xBJ5MurqlOBpH3lXlqWzjjz66nbSkl4xktVgB+ZaGmC51wWr/RWlr3s6hQ3QVk9FP4UYeW+gHgxh94A8j43EtV6E0VgFqA8aKET2yXKZ+73gvgmNwtOCDc7m613QJ9yI/QX63MoDH0iBAxfIegozIii2W0oCy7syWKjE+Uu7Wm0r9/awP794sFb8grkOPDULp+uMgwBWXXId/dteJx8QWTYTpnjpVH/Ee4t4JtFM9EWiWYsXkVhlYZDwRrOwgf3mZmLAgY8W7PdBDtfrSgP9fgdY7VRZilVDHt3shnNW48Fqo958ObxaB4STI8L2AkzfK3QgsopNBlVMaRUXuoNs4RsLz4Y0SExnSK2+hdVey9059ymeX2rvAsFCc326U9zlIcf0fwNYAVqaH/GpK/54xsAPBAZL7tYRrXQyON2QAvG23FM48M3tIiCyiiyXZHxMk/YEDfiDb6qLmh1PysW7dy8msMRuUf10DuM0TePB8s/PwMc/yhG43wJW+gjFhp84AnvOOcWD+fLkKhAalfpSB2Cz6sAwijxawUKxHwxZBdwX7T7XJZk1KDy/NMHp5tpOsmawQagdnaeN6q+2opEFgdUmPv7xfig/OQRW7pbAIhN2fYlLupNtPHGkNexrgUUfzPVUCEOv0OCzJlR+J+8sRCKrypN1eK9gkWXzw40pPfbLP4UfBXx+FOovL9DIok3U23wprghYNjRuP1XBwDwRLN8TK5VKoCu2BJqwjzcFVvtnEFUdLxaka1PhmOfUD546MFhEc72AFTTQfCOjU+TIqnS7FSyyGPGGDKxn3aFP8y47medC33pDKAH1NPMc7A9S7LMHwWpHBerim5f299hJW1R//fkKoqphzH4YNx8LxjKUWbw5hSc1FqA5ewa8yKbJk3VAZBUNfo6uODlbPYxhdA2ucG3xVOxTTfcCki6EwaJtW4DNLIHlSqOOboEsFFVF0zLNWtRi3tYCi2JpgmARCwsYDWqDcXd891eOLF9kFUSRxZVZxZ1N9ypcYbAsvgk5bfB1JSYbVcUJKhdfikB6NwgWjqrSbWV9dXA1p1ClvSeVxCKpQykQOlDE3bGAQyKLIevdF1nYymLvfdAaVz7tgpPELgPWmBNY+dRR/X2brwg3pTQzi5ZGxMJTDWL8rafX+xJbC9q6YOEtZiBYRGAB8yAHqhHHSEe+cmT5IuuAR9fWhZBD0qQrq6uYJHYZsKIPbJ9UkLPlakPwEvtUKfLVVhhIY91YuIHT5f3YDvOc6iv65Cb8dQfsLFEo1oCTh/Box0KHiKyILP+5sS6cMWAljVMzNvLIpwt+WXxv8EkCC0018L+2kRH5GK10tRcNvtfHUIu8W0ArNmafi4uEGWyffYDAwegAxTEEjAOc5fOVJA9WgepCThTGjVOzZxbtc86nmvGRJkDK7PsyTqwg3O5xJ/nc9c+4q4hJJ1lpVsy6ndtjS/XXSrCeeHPhmZYgwLGGAKyCcmFglSOrQv1CTheq5h4as024ESIc8tG7tOGx3a4+QwNnfJLAInmLz8gXT/0cpElqzvgppZsFK+HkYn46fEjlB21ydsHByEQVFpTtD0aRI+sdhbKafF4n6m825uFZMPPga6lnx6Qzsu7IGDcbjRo2Tyx53kc5G/N+mLoKEW+X71q+NjTwwOP88A8DCye+hLq1t+cB2T0ILXIa0D4e+LxR870Y6UICliuUc6mnE5nR3MNrYNWm6wKjQXUnFnrhSXv5Ix51c0KfB60XHpNpj/nLFV7fCljYjBTKizrNfwq41k7ShEiSzWJ2B7wEujAkC4H1ykVUWZElriav5Vul/jDlSKITEmGBHjNLp6eLel/s7JfUz0GHwZJPPfy2AksJFtH28nZrvFDwB1S0PuuoljcP3v8X+IUhWIcK2e3FPZXBrZHsdo/o5PP30uz+K30ZexxUrtXyZ3mdGr54/OMbuVrteETv4vtypVaFuMAZXujZlGs7n19QvlsJFmro4cECdpqIwz7u/5WWjPSuiZUcoLmEM35GZPsazu/vB4vWpEFkiSXJ9SDyolo3jy3+VwmsRxYsvLWLHfZxFLfX9K6NFQkq9/oXsON657xZBqzed+UqxnjHyXqQrB98rUNkGgVgibOTUe2MaGRFYIUFfVFNOx1PNBxqj6i7QuTvt0e2f/fbuBJYgZc0lapZ+Mwhs9lm9h4NjRxIJTV0DYVCYhXqL2Ix1FBciPRLg8q3Edhut/88sAIHxXTEcnO2OH3QEcFCI9peBKmGlJxCYr3/qNedalVaCSrMvWr/Zee7f+hcrBlJ/W8eLbseKjx+OxMBi+4T5AsdEH6vEFhGEyoRdWhKpf9XMvVHnFyCg9IgcWTT4XrPZ88/BuHcIgmsGbAErBqAVeS9QvhMozUjGVZ/Ilh47WKwpZ2Pltr2DO0VKHRksFyg1sY3oN4EsP57VG4FZQb1ZZfojwQLr12kya95chkeKjAOgwbVOLDQ0q9qTJGoedFOiezcHFg4Q1HC9fNm4ppAd8BFo56FDPNrsJ3Q9wwfH6uxxcf33zjonB0dsMgGM7TzWBx5ARwEUjcEC62gD6JUBbqhlxzDSOrB6WcC608HCxtaX3jLTiJZb4iayBZ7CcfcFl4096GSMtH6xUvas3ODYJGFnjhYOku3jtl5/sc/z7Om1k5wY7PophsElp3vDRYZ0ZW83/q043quZ/knrIW672dc/S1g4fVNdOPk3DUuydVMXDF4lV6J7NwoWNiID4ehWgvPZVKIzvxCXJlBEVIWGv17wMLqkN2BOBptZuhYU6AgUFcNWtEa49r9fbhiMOPqLwKL2PD5I7TN1DqFKns+og2opdLJGwaz8/3BIiGtfimfqyWSBdhhruV4tCPU9rzNeBQ1SXEVfRlWfx1YJKKF91qjxekBXvjfU7RB3nbwKnnPsbrmxuBDCUTpTfFhiIxauk5bMZidPwEsYXP6sH+Pdy/n4R3yY8tauJ7rWJZlwvc45lv9EKteRtVfDFa0Ob3EbE7P17qpTy2Xbwn7wDOq/mawwM3ppQAt2rhs1iQZZeZLJV/A0UPM9eGva5PIzu2D1ZZ2W/ts9Fut1v19y2fsHv9reJ8/Bg7k8YhDCYrlzRlVGVhquPhmuFCQkdNP2jKfYZWBpaUZNU6vl5lVGVipdSOE2kX2gWfnLwNLyVbMyb79DKwLwZV93xlY2clOBlZ2MrCyk4GVnexkYGUnAys7GVjZyc5p5/8A1tf7O+RLiDYAAAAASUVORK5CYII=";
const logoBuf = logoPath ? fs.readFileSync(logoPath) : Buffer.from(LOGO_B64, "base64");
const header = new Header({ children: [
  new Table({ visuallyRightToLeft: true, width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W - 2600, 2600],
    borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: CONTENT_W - 2600, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
        children: [new Paragraph({ bidirectional: true, alignment: AlignmentType.START, spacing: { after: 0 }, children: [
          run(spec.date + "  |  עמוד ", { size: 18, color: GREY }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: GREY }),
          run(" מתוך ", { size: 18, color: GREY }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: GREY })
        ] })] }),
      new TableCell({ width: { size: 2600, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
        children: [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 0 }, children: [
          new ImageRun({ type: "png", data: logoBuf, transformation: { width: 150, height: 69 } })
        ] })] })
    ] })]
  }),
  new Paragraph({ spacing: { after: 120 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD", space: 1 } }, children: [] })
] });

const footer = new Footer({ children: [
  new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { before: 80, after: 0 },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD", space: 4 } },
    children: [run("כל הזכויות שמורות " + new Date().getFullYear() + " © אהוי קריאייטיב ויזמות בע\"מ  |  רחוב יגאל אלון 104, תל אביב  |  טלפון: 03-6024205", { size: 16, color: GREY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 },
    children: [new TextRun({ text: "info@ahoy-creative.com  |  www.ahoy-creative.com  |  www.agvertising.biz", font: FONT, size: 16, color: GREY })] })
] });

// ---- document ----
const headerBlock = [
  ...(spec.number ? [para([run("הזמנת עבודה " + spec.number, { bold: true, size: 24 })], { after: 40 })] : []),
  para([run("אל: ", { bold: true }), run(spec.to)], { after: 40 }),
  para([run("מאת: ", { bold: true }), run(spec.from || "גלעד ורשבסקי ועמרי ירושלמי | אהוי קריאייטיב ויזמות בע\"מ")], { after: 240 }),
  H1(spec.title)
];

const signoff = [
  para([run("")], { after: 200 }),
  para([run("בברכה,")], { after: 40 }),
  para([run("גלעד ורשבסקי | עמרי ירושלמי", { bold: true })], { after: 40 }),
  para([run("אהוי קריאייטיב ויזמות בע\"מ")], { after: 120 })
];

const signature = spec.kind === "work_order" ? [
  new Paragraph({ children: [new PageBreak()] }),
  H2("אישור הזמנת העבודה"),
  P("מורשה החתימה של הלקוח מצהיר בזאת כי הוא מורשה לבצע את ההזמנה בשם הלקוח וערב אישית לביצוע התשלום."),
  para([run("")], { after: 200 }),
  table({ columns: ["שם החברה", "תאריך", "חתימת מורשה + חותמת החברה"], widths: [3200, 2200, CONTENT_W - 5400],
    rows: [["", "", ""], ["", "", ""]] }),
  para([run("")], { after: 200 }),
  P("הצהרת מורשה החתימה: הנני מצהיר כי קראתי את כל התנאים בהזמנת העבודה הנ\"ל וכי אני מורשה לבצע את ההזמנה בשם הלקוח וערב אישית לביצוע התשלום."),
  para([run("")], { after: 200 }),
  P("שם מורשה החתימה: ______________________________"),
  para([run("")], { after: 120 }),
  P("תפקיד: ______________________________"),
  para([run("")], { after: 120 }),
  P("חתימה: ______________________________")
] : [];

const doc = new Document({
  creator: "Ahoy Creative",
  title: spec.fileTitle || spec.title,
  styles: {
    default: { document: { run: { font: FONT, size: 22, rightToLeft: true } } },
    paragraphStyles: [{ id: "Normal", name: "Normal", run: { font: FONT, size: 22, rightToLeft: true },
      paragraph: { alignment: AlignmentType.START } }]
  },
  numbering: { config: [
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.RIGHT,
      style: { paragraph: { indent: { left: 0, right: 500, hanging: 360 } }, run: { font: FONT } } }] },
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.RIGHT,
      style: { paragraph: { indent: { left: 0, right: 500, hanging: 360 } }, run: { font: FONT } } }] }
  ] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 1500, bottom: 1300, left: MARGIN, right: MARGIN, header: 500, footer: 500 } },
      rtlGutter: true, titlePage: false },
    headers: { default: header },
    footers: { default: footer },
    children: [...headerBlock, ...renderBlocks(spec.blocks), ...signoff, ...signature]
  }]
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(outPath, buf); console.log("wrote", outPath, buf.length, "bytes"); });
