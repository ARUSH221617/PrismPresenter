# Slide Storyboard Specification Schema (`T711-structure.md`)

## 1. Document Metadata
- **Module_Code**: [e.g., CS-301 / Data Structures & Algorithms]
- **Unit_Topic**: [e.g., Unit 4: Balanced Search Trees]
- **Target_Audience**: [e.g., Undergraduate Engineering / Technical Executive Briefing]
- **Deck_Version**: [e.g., v1.2]
- **Slide_Range**: [e.g., Slides 1–10]

---

## 2. Template Layout & Archetype Mapping
Deck `T711.pptx` follows a structured, methodical rhythm designed for instructional hierarchy, progression, and multi-metric syntheses. Every storyboarded unit must be mapped to the matching archetype slots:

| Slide Index | Archetype Name | Slot Budget | Pedagogical Intent & Visual Logic |
| :--- | :--- | :--- | :--- |
| **0** | `title_cover` | 3 slots | Deck Title, Subtitle/Module, Presenter/Date metadata. |
| **1** | `process_timeline` | 6 slots | Agenda sequence, lifecycle phases, or chronological procedural steps. |
| **2** | `content_bullets` (Dual-Split) | 2 slots | High-contrast conceptual pairing (e.g., Problem vs. Solution, Before vs. After). |
| **3** | `content_bullets` (Framework Grid) | 10 slots | Multi-category taxonomy (e.g., 2–4 column cards with title + explanation). |
| **4** | `content_bullets` (Deep Breakdown) | 19 slots | Granular rule matrix, nested component steps, or comprehensive technical specs. |
| **5** | `content_bullets` (Comparative T-Chart) | 9 slots | 3-column evaluation, trade-off analysis, or structured criterion breakdown. |
| **6** | `content_bullets` (Execution Roadmap) | 23 slots | Multi-tier pipeline breakdown, operational instructions, or deep drill-downs. |
| **7** | `metrics_stats` (Dashboard A) | 26 slots | Primary KPI/benchmark metrics paired with analytical descriptors. |
| **8** | `metrics_stats` (Dashboard B) | 25 slots | Experimental outcomes, efficiency bounds, or validation scorecards. |
| **9** | `content_bullets` (Summary & Next Steps)| 17 slots | Key takeaways, operational checklist, and future milestone commitments. |

---

## 3. Slide Content Schema

For each slide, extract and specify content according to this standardized structure:

### Slide [Slide_Number]: [Slide Title]
- **Template_Index**: [0 to 9]
- **Archetype**: [`title_cover` | `process_timeline` | `content_bullets` | `metrics_stats`]
- **Slide_Type**: [Agenda | Theory / Definition | Classification | Worked Example | Benchmark / Performance | Summary]
- **Core_Concept**: [Concise 1-line statement capturing the essential pedagogical takeaway]

#### Animation Sequence (Sequential Reveals)
Specify build order strictly matching sequenced indicators (`①`, `②`, `③`...):
1. **Step 1 (`①`)**: [Anchor concept / Baseline state / Initial equation]
2. **Step 2 (`②`)**: [Transformation / Intermediate condition / Comparative point]
3. **Step 3 (`③`)**: [Evaluation / Highlight / Final deduction]

#### Structured & Mathematical Elements
- **Formal Definitions / Rules**: [Rigorous theoretical rule, theorem statement, or architectural principle]
- **Formulas / LaTeX Expressions**: [Standard LaTeX notation; e.g., $T(n) = 2T(n/2) + \mathcal{O}(n)$]
- **Data Sets / Metrics / Invariants**: [Target values, complexity bounds, or parameter constraints]

#### Visual & Non-Textual Annotations
- **Diagrams / Topology Maps**:
  - `Root / Focal Point`: [Central node, system boundary, or primary axis]
  - `Connectors / Child Nodes`: [Dependencies, transitions, or sub-components]
- **Eliminations / Deprecations (Crossed Out)**: [Suboptimal alternatives crossed out with rationale, e.g., $\text{O}(n^2)$ naive search]
- **Validations (Checkmarked)**: [Validated design choices or verified assertions tagged with `✓`]
- **Key Invariants (Boxed / Highlighted)**: [Immutable properties or definitive results placed in visual focus frames]
- **Instructor / Strategic Callouts**: [Pragmatic tips prefixed with `[Note]`, `[Warning]`, or `[Industry Standard]`]

---

## 4. Transcription Example (Applied to Slide 8: Algorithmic Complexity Benchmark)

### Slide 8: Empirical Performance & Scaling Metrics
- **Template_Index**: 8
- **Archetype**: `metrics_stats` (Dashboard B)
- **Slide_Type**: Benchmark / Performance
- **Core_Concept**: Empirical verification of $O(\log n)$ balanced tree queries versus linear degradation in unbalanced trees.

#### Animation Sequence (Sequential Reveals)
1. **Step 1 (`①`)**: Display primary throughput indicator: **42.8k ops/sec** under sustained concurrent query loads.
2. **Step 2 (`②`)**: Reveal comparative memory allocation metric: **14.2 MB** heap overhead at $N = 10^6$ elements.
3. **Step 3 (`③`)**: Render the tail-latency threshold: **$p_{99} < 1.4\text{ ms}$** across synthetic enterprise workloads.
4. **Step 4 (`④`)**: Unfold comparative summary table contextualizing worst-case versus amortized operations.

#### Structured & Mathematical Elements
- **Formal Definitions / Rules**: An AVL tree maintains height balance such that balance factor $BF(v) \in \{-1, 0, 1\}$ for every vertex $v$.
- **Formulas / LaTeX Expressions**:
  - Height bound: $h < 1.44 \log_2(N + 2) - 0.328$
  - Search Cost: $C_{\text{lookup}} = \mathcal{O}(\log n)$
- **Data Sets / Metrics / Invariants**:
  - Metric 1: `Throughput`: $42,800\text{ req/s}$ ($+18\%$ vs. Red-Black baseline)
  - Metric 2: `Rebalance Cost`: $1.2\text{ rotations per insertion}$
  - Metric 3: `Worst-Case Depth`: $18\text{ hops}$ at $N = 100,000$

#### Visual & Non-Textual Annotations
- **Diagrams / Topology Maps**:
  - `Root / Focal Point`: Central comparison card displaying Tree Depth vs. Traversal Latency
  - `Connectors / Child Nodes`: Sub-charts linking rotate-left (`LL`) and rotate-right (`RR`) cost curves
- **Eliminations / Deprecations (Crossed Out)**: Unbalanced Binary Search Tree Worst Case $\mathcal{O}(n)$ struck through as unviable for SLA compliance.
- **Validations (Checkmarked)**: Dual-rotation rebalancing confirmed (`✓ AVL invariant preserved`).
- **Key Invariants (Boxed / Highlighted)**: Boxed upper-bound latency constraint: `max_latency <= 2.0ms`.
- **Instructor / Strategic Callouts**: `[Industry Standard]`: In write-heavy engines, consider Red-Black trees over AVL to reduce rotation overhead during batch inserts.
