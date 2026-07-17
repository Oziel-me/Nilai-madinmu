# Graph Report - .  (2026-07-17)

## Corpus Check
- Corpus is ~11,132 words - fits in a single context window. You may not need a graph.

## Summary
- 66 nodes · 64 edges · 16 communities (8 shown, 8 thin omitted)
- Extraction: 100% EXTRACTED · 0% INFERRED · 0% AMBIGUOUS
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- Development Dependencies & Linter
- Main Application Components
- Package Manifest & Scripts
- Oxlint Rules Configuration
- ECharts & DOM Dependencies
- ECharts Integration for React
- File Saving Utilities
- Animation Library (Framer Motion)
- Image Generation Utilities
- PDF Generation Utilities
- Icon Library (Lucide React)
- React Core Library
- Table Library (TanStack Table)

## God Nodes (most connected - your core abstractions)
1. `scripts` - 5 edges
2. `plugins` - 3 edges
3. `react` - 3 edges
4. `rules` - 3 edges
5. `xlsx` - 3 edges
6. `App()` - 3 edges
7. `react/only-export-components` - 2 edges
8. `@tanstack/react-table` - 2 edges
9. `echarts` - 2 edges
10. `echarts-for-react` - 2 edges

## Surprising Connections (you probably didn't know these)
- `App()` --references--> `xlsx`  [EXTRACTED]
  src/App.jsx → package.json

## Import Cycles
- None detected.

## Communities (16 total, 8 thin omitted)

### Community 0 - "Development Dependencies & Linter"
Cohesion: 0.13
Nodes (15): oxlint, devDependencies, oxlint, tailwindcss, @tailwindcss/vite, @types/react, @types/react-dom, vite (+7 more)

### Community 1 - "Main Application Components"
Cohesion: 0.22
Nodes (6): plugins, xlsx, oxc, react, App(), xlsx

### Community 2 - "Package Manifest & Scripts"
Cohesion: 0.20
Nodes (9): name, private, scripts, build, dev, lint, preview, type (+1 more)

### Community 3 - "Oxlint Rules Configuration"
Cohesion: 0.33
Nodes (5): rules, react/only-export-components, react/rules-of-hooks, $schema, warn

### Community 4 - "ECharts & DOM Dependencies"
Cohesion: 0.40
Nodes (5): echarts, dependencies, echarts, react-dom, react-dom

## Knowledge Gaps
- **30 isolated node(s):** `$schema`, `oxc`, `react/rules-of-hooks`, `warn`, `name` (+25 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **8 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `dependencies` connect `ECharts & DOM Dependencies` to `Main Application Components`, `Package Manifest & Scripts`, `ECharts Integration for React`, `File Saving Utilities`, `Animation Library (Framer Motion)`, `Image Generation Utilities`, `PDF Generation Utilities`, `Icon Library (Lucide React)`, `React Core Library`, `Table Library (TanStack Table)`?**
  _High betweenness centrality (0.695) - this node is a cross-community bridge._
- **Why does `devDependencies` connect `Development Dependencies & Linter` to `Package Manifest & Scripts`?**
  _High betweenness centrality (0.363) - this node is a cross-community bridge._
- **Why does `xlsx` connect `Main Application Components` to `ECharts & DOM Dependencies`?**
  _High betweenness centrality (0.361) - this node is a cross-community bridge._
- **What connects `$schema`, `oxc`, `react/rules-of-hooks` to the rest of the system?**
  _30 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `Development Dependencies & Linter` be split into smaller, more focused modules?**
  _Cohesion score 0.13333333333333333 - nodes in this community are weakly interconnected._