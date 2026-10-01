// House style for PDF reports. Usage:
//   #import "/templates/report.typ": *
//   #show: report.with(title: "…", subtitle: "…", author: "…", date: "1 октября 2026")
// Compile: python tools/render.py doc.typ   (fonts come from /fonts, root is the repo)

// Design tokens — change here, never inline in documents.
#let accent = rgb("#1F4E79")
#let ink = rgb("#1A1A1A")
#let muted = rgb("#6B7280")
#let rule-c = rgb("#D1D5DB")
#let tint = rgb("#F3F6FA")
#let serif = ("PT Serif",)
#let sans = ("PT Sans",)

#let callout(body, title: none) = block(
  width: 100%, fill: tint, inset: (x: 12pt, y: 10pt), radius: 2pt,
  stroke: (left: 2.5pt + accent), breakable: false,
)[
  #if title != none { text(font: sans, weight: "bold", fill: accent, title); linebreak() }
  #body
]

#let report(
  title: "", subtitle: none, author: none, date: none,
  lang: "ru", toc: false, numbered: true, paper: "a4", body,
) = {
  set document(title: title, author: if author == none { () } else { author })
  set page(
    paper: paper,
    margin: (x: 24mm, top: 24mm, bottom: 26mm),
    header: context if counter(page).get().first() > 1 {
      set text(font: sans, size: 8pt, fill: muted)
      title
      h(1fr)
      if date != none { date }
    },
    footer: context {
      set text(font: sans, size: 8pt, fill: muted)
      h(1fr)
      counter(page).display("1 / 1", both: true)
    },
  )
  set text(font: serif, size: 10.5pt, fill: ink, lang: lang, hyphenate: true,
           costs: (widow: 100%, orphan: 100%))
  set par(justify: true, leading: 0.7em, spacing: 1.15em)
  set heading(numbering: if numbered { "1.1" } else { none })
  // Headings are sticky (never stranded at a page bottom) — Typst default since 0.12.
  show heading: set text(font: sans, fill: accent, hyphenate: false)
  show heading: set par(justify: false)
  show heading.where(level: 1): set text(size: 17pt)
  show heading.where(level: 1): set block(above: 1.8em, below: 0.9em)
  show heading.where(level: 2): set text(size: 13pt)
  show heading.where(level: 2): set block(above: 1.5em, below: 0.7em)
  show heading.where(level: 3): set text(size: 11pt, fill: ink)
  // Russian typography: glue short words to the next word and numbers to their units.
  show regex("(?i)(^|\s)(в|во|и|с|со|к|ко|о|об|у|а|я|на|не|но|по|за|из|до|от|для|без|или)\s"): it => it.text.trim(at: end) + sym.space.nobreak
  show regex("\d [%‰₽$€]|\d (млн|млрд|тыс|руб|кг|км|мм|см|п\.п\.|г\.)"): it => it.text.replace(" ", sym.space.nobreak.narrow)
  show regex(" — "): sym.space.nobreak + sym.dash.em + " "
  show link: set text(fill: accent)
  show raw: set text(size: 9pt)
  set list(marker: text(fill: accent)[•], indent: 0.4em, body-indent: 0.6em)
  set enum(indent: 0.2em)

  // Booktabs tables: heavy top, rule under header, hairlines between rows, no verticals.
  set table(
    inset: (x: 6pt, y: 5pt),
    stroke: (x, y) => if y == 0 { (top: 1pt + ink, bottom: 0.6pt + ink) } else { (bottom: 0.4pt + rule-c) },
  )
  show table.cell.where(y: 0): set text(font: sans, weight: "bold", size: 9pt)
  show table: set text(size: 9.5pt, number-type: "lining", number-width: "tabular")
  show table: set par(justify: false)
  show figure.where(kind: table): set figure.caption(position: top)
  show figure.caption: set text(font: sans, size: 8.5pt, fill: muted)
  // Tables that fit on a page never split (no orphan rows). Taller ones must split: if QA
  // reports "table starts at the page bottom", put #pagebreak(weak: true) before that table.
  // Row count is estimated from the cells (deterministic; measure() proved unreliable here).
  show figure.where(kind: table): it => {
    let t = it.body
    let ncol = if type(t.columns) == array { t.columns.len() } else if type(t.columns) == int { t.columns } else { 1 }
    let rows = t.children.len() / calc.max(ncol, 1)
    set block(breakable: rows > 25)  // must reach the figure's own block, a wrapper won't do
    it
  }

  // Title block
  block(below: 1.6em)[
    #text(font: sans, size: 24pt, weight: "bold", fill: accent, hyphenate: false, title)
    #if subtitle != none { v(0.2em); text(font: sans, size: 13pt, fill: muted, subtitle) }
    #if author != none or date != none {
      v(0.6em)
      text(font: sans, size: 9pt, fill: muted, (author, date).filter(x => x != none).join("  ·  "))
    }
    #v(0.4em)
    #line(length: 100%, stroke: 0.6pt + rule-c)
  ]
  if toc {
    outline(title: if lang == "ru" { "Содержание" } else { "Contents" }, depth: 2)
    pagebreak(weak: true)
  }
  body
}
