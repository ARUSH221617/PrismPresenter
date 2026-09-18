# Slide Storyboard Specification Schema (`T711-structure.md`)

## 1. Document Metadata
- **Course_Domain**: [e.g., Business Analytics & Operations / High School Advanced Mathematics / Corporate Strategy]
- **Module_Chapter**: [e.g., Module 4: Quantitative Decision Modeling]
- **Lesson_Topic**: [e.g., Inventory Optimization & Economic Order Quantity (EOQ)]
- **Deck_ID / Page_Number**: [e.g., Deck-711 / Unit 2]
- **Slide_Range**: [e.g., Slides 1–10]
- **Target_Template**: `T711.pptx` (Professional, Clean & Balanced Layouts)

---

## 2. Layout & Archetype Mapping

The `T711.pptx` template provides a balanced, clean aesthetic featuring high-clarity bullet grids, process timelines, and high-density metric/KPI dashboards. Each storyboarded slide maps directly to one of the following template archetypes:

| Template Index | Archetype Name | Slot Count | Primary Layout Focus & Usage |
| :--- | :--- | :--- | :--- |
| **Slide 0** | `title_cover` | 3 Slots | Hero Title, Subtitle, Presenter/Course Metadata |
| **Slide 1** | `process_timeline` | 6 Slots | Sequential Workflow, Milestones, Step-by-Step Evolution |
| **Slide 2** | `content_bullets` (2-Col / Split) | 2 Slots | Core Definition vs. Intuition, Problem vs. Solution |
| **Slide 3** | `content_bullets` (4–5 Cards) | 10 Slots | Categorical Groupings, Structural Pillars, Core Properties |
| **Slide 4** | `content_bullets` (Modular Matrix) | 19 Slots | High-Density Matrix, Multi-Attribute Comparison Grid |
| **Slide 5** | `content_bullets` (3-Pillar Deep-Dive) | 9 Slots | 3-Column Theoretical Framework, Factor Breakdown |
| **Slide 6** | `content_bullets` (Comprehensive Exercise) | 23 Slots | Multi-Step Worked Problem, Scenario Analysis, Q&A Matrix |
| **Slide 7** | `metrics_stats` (Dashboard Grid) | 26 Slots | Multi-KPI Metrics, Operational Targets, Variances |
| **Slide 8** | `metrics_stats` (Comparative Benchmark) | 25 Slots | Benchmark Performance, Parameter Sensitivity, Scenarios |
| **Slide 9** | `content_bullets` (Executive Summary) | 17 Slots | Key Takeaways, Action Roadmap, Summary Checklist |

### Spatial & Quadrant Placement Keys
When organizing content inside archetypes, map positional containers using:
- `Header_Area`: Top Banner (Slide Title, Category Pill, Context Breadcrumb)
- `Slot_L` / `Slot_R`: Left / Right Split Containers (for 2-Slot archetypes)
- `Col_1` / `Col_2` / `Col_3`: Vertical Pillar Columns (for 3–6 slot layouts)
- `Quadrant_TL` / `Quadrant_TR` / `Quadrant_BL` / `Quadrant_BR`: 2×2 Grid Layouts
- `Card_[1..N]`: Discrete modular content containers / metric cards
- `Footer_Callout`: Bottom Takeaway Bar or Pro-Tip Callout Banner

---

## 3. Slide Content Schema

For every slide in the storyboard, transcribe and structure content using the following specification:

### Slide [Slide_Number]: [Slide Title]
- **Archetype**: [`title_cover` | `process_timeline` | `content_bullets` | `metrics_stats`]
- **Layout_Target**: [Slide Index 0–9 from `T711.pptx`]
- **Slide_Type**: [`Agenda / Timeline` | `Theory / Definition` | `Classification / Matrix` | `Worked Example / Case Study` | `Metrics & Evaluation` | `Summary / Action Plan`]
- **Core_Concept**: [Concise 1-sentence pedagogical or strategic objective]

#### Layout Slots Allocation
- **Container / Slot Reference**: [Assign text blocks, cards, or metric tiles to specific template slots]

#### Animation Sequence (Numbered Steps)
Strictly order progressive reveals using circled step numbers (`①`, `②`, `③`, `④`, `⑤`...):
1. **Step 1 (`①`)**: [Trigger / Initial state - Concept framing, baseline equation, or primary card]
2. **Step 2 (`②`)**: [Secondary reveal - Intermediate derivation, benchmark comparison, or branching step]
3. **Step 3 (`③`)**: [Tertiary reveal - Secondary parameter or constraint application]
4. **Step 4 (`④`)**: [Synthesis / Conclusion - Final result, boxed answer, or KPI trigger]

#### Mathematical / Core Structured Elements
- **Definitions / Governing Rules**: [Formal axioms, business rules, or mathematical statements]
- **Formulas / Equations**: [LaTeX formatted expressions, e.g., $TC = \frac{D}{Q}S + \frac{Q}{2}H$]
- **Parameters & Sets**: [Explicit definitions of sets, variables, constraints, or matrices, e.g., $D \in \mathbb{R}^+, Q^* > 0$]

#### Visual / Non-Textual Annotations
- **Workflows & Diagrams**:
  - `Source Node`: [Initial State / Input]
  - `Transitions / Edges`: [Condition or Transformation]
  - `Target Node`: [Resulting State / Output]
- **Eliminations (Crossed Out)**: [List suboptimal alternatives, invalid solutions, or eliminated terms with strikethrough `~~text~~` and rationale]
- **Confirmations (Checkmarked `✓`)**: [List validated conditions, optimal values, or verified invariants]
- **Key Invariants (Boxed / Highlighted)**: [List boxed final formulas, optimal operational parameters, or primary takeaways]
- **Instructor / Strategic Callouts**: [Callouts prefixed with `💡 Pro-Tip`, `⚠️ Common Pitfall`, `📌 Key Theorem`, or `🔍 Insight`]

---

## 4. Concrete Transcription Example

### Slide 1: End-to-End Inventory Decision Flow
- **Archetype**: `process_timeline`
- **Layout_Target**: Slide Index 1 (6 Slots)
- **Slide_Type**: Agenda / Timeline
- **Core_Concept**: Establish the sequential 6-phase quantitative optimization workflow for deterministic demand systems.

#### Layout Slots Allocation
- `Slot_1`: Phase 1 – Demand Forecasting ($\hat{D}$)
- `Slot_2`: Phase 2 – Cost Parameter Estimation ($S, H, C$)
- `Slot_3`: Phase 3 – Total Cost Function Formulation ($TC(Q)$)
- `Slot_4`: Phase 4 – First-Order Optimization ($\frac{dTC}{dQ} = 0$)
- `Slot_5`: Phase 5 – Robustness & Sensitivity Testing ($\pm 20\%$)
- `Slot_6`: Phase 6 – Operational Reorder Point Execution ($ROP$)

#### Animation Sequence (Numbered Steps)
1. **Step 1 (`①`)**: Highlight Slot 1 & 2 (`Data Acquisition Phase`).
2. **Step 2 (`②`)**: Reveal Slot 3 & 4 with mathematical curve overlay (`Analytical Formulation`).
3. **Step 3 (`③`)**: Reveal Slot 5 & 6 with target operational parameters (`Operational Deployment`).

#### Mathematical / Core Structured Elements
- **Definitions / Governing Rules**: Continuous-review inventory policy $(s, Q)$ with constant deterministic lead time $L$.
- **Formulas / Equations**:
  $$\text{Reorder Point: } ROP = d \times L$$
  $$\text{Economic Order Quantity: } Q^* = \sqrt{\frac{2DS}{H}}$$
- **Parameters & Sets**:
  - $D$: Annual Demand (Units/Year)
  - $S$: Setup / Ordering Cost (\$/Order)
  - $H$: Unit Holding Cost (\$/Unit/Year)
  - $L$: Replenishment Lead Time (Days)

#### Visual / Non-Textual Annotations
- **Workflows & Diagrams**:
  - `Demand Stream (d)` $\rightarrow$ `Inventory Depletion` $\rightarrow$ `Hit ROP` $\rightarrow$ `Trigger Batch Q*`
- **Confirmations (`✓`)**: `✓ Lead time is strictly deterministic ($L = \text{const}$)`.
- **Instructor / Strategic Callouts**:
  - `💡 Pro-Tip`: "Never optimize order quantity $Q$ in isolation from storage capacity constraints."

---

### Slide 4: Multi-Model Inventory Policy Matrix
- **Archetype**: `content_bullets` (Modular Matrix)
- **Layout_Target**: Slide Index 4 (19 Slots)
- **Slide_Type**: Classification / Matrix
- **Core_Concept**: Systematically compare Classical EOQ, Production Order Quantity (POQ), and Quantity Discount models across structural assumptions.

#### Layout Slots Allocation
- `Header_Area`: Comparative Model Taxonomy
- `Col_1 (Classical EOQ)`: Instantaneous replenishment, fixed unit cost $C$, zero stockouts allowed.
- `Col_2 (POQ / EPQ)`: Gradual production replenishment rate $p > d$, finite batch buildup.
- `Col_3 (Quantity Discounts)`: Piecewise linear unit cost $C(Q)$, step-level price breaks.
- `Row_Bottom`: Optimal batch size formula expressions per archetype.

#### Animation Sequence (Numbered Steps)
1. **Step 1 (`①`)**: Populate column headers and baseline Classical EOQ assumptions.
2. **Step 2 (`②`)**: Contrast POQ model with finite replenishment dynamics ($p > d$).
3. **Step 3 (`③`)**: Display Quantity Discount price break curves and feasible region tests.
4. **Step 4 (`④`)**: Highlight bottom cost-curve comparison row.

#### Mathematical / Core Structured Elements
- **Formulas / Equations**:
  - Classical EOQ: $Q^* = \sqrt{\frac{2DS}{H}}$
  - POQ Model: $Q^*_p = \sqrt{\frac{2DS}{H \left(1 - \frac{d}{p}\right)}}$
  - Quantity Discount Total Cost: $TC(Q) = \frac{D}{Q}S + \frac{Q}{2}iC_k + D \cdot C_k$
- **Parameters & Sets**:
  - $p$: Production rate ($p > d$)
  - $C_k$: Tiered unit acquisition cost where $C_1 > C_2 > C_3$ for order quantities $q_1 < q_2 < q_3$.

#### Visual / Non-Textual Annotations
- **Eliminations (Crossed Out)**:
  - ~~Instantaneous Delivery Assumption~~ (Eliminated in POQ formulation due to finite line speed $p$).
  - ~~Constant Unit Price $C$~~ (Eliminated in Quantity Discount evaluation).
- **Confirmations (`✓`)**:
  - `✓ $1 - \frac{d}{p} < 1 \implies Q^*_p > Q^*$ (POQ batch sizes are always strictly larger than EOQ batch sizes)`.
- **Key Invariants (Boxed)**:
  - $\boxed{TC_{\text{min}} = \min_k \left\{ TC(Q_k^*) \mid Q_k^* \text{ is feasible} \right\}}$
- **Instructor / Strategic Callouts**:
  - `⚠️ Common Pitfall`: "Always check if calculated $Q^*$ meets the minimum volume qualification for that discount tier before validating."

---

### Slide 7: Operational Performance & Cost Variance Dashboard
- **Archetype**: `metrics_stats` (Dashboard Grid)
- **Layout_Target**: Slide Index 7 (26 Slots)
- **Slide_Type**: Metrics & Evaluation
- **Core_Concept**: Evaluate simulated distribution metrics and sensitivity variances against enterprise KPI targets.

#### Layout Slots Allocation
- `Card_1 (Top-Left KPI)`: Annual Holding Cost ($\$ 42,500$ vs Target $\$ 45,000$ | **-5.6%**)
- `Card_2 (Top-Right KPI)`: Annual Setup Cost ($\$ 42,800$ vs Target $\$ 45,000$ | **-4.9%**)
- `Card_3 (Mid-Left KPI)`: Order Cycle Time ($14.2 \text{ Days}$ vs Baseline $28.0 \text{ Days}$)
- `Card_4 (Mid-Right KPI)`: Service Level Fill Rate ($99.4\%$ vs SLA $98.0\%$ | **+1.4%**)
- `Card_5 (Bottom Summary)`: Total Annual System Cost ($\$ 85,300$ vs Budget $\$ 102,000$ | **$\Delta = -\$16,700$**)
- `Card_6 (Risk Diagnostic)`: Stockout Probability ($\alpha = 0.6\%$)

#### Animation Sequence (Numbered Steps)
1. **Step 1 (`①`)**: Fade in primary cost KPI cards (Holding Cost & Setup Cost balancing).
2. **Step 2 (`②`)**: Reveal Cycle Time and Fill Rate SLA compliance metrics.
3. **Step 3 (`③`)**: Trigger bottom aggregate savings summary card with positive delta indicator.

#### Mathematical / Core Structured Elements
- **Formulas / Equations**:
  $$\text{Total Cost Variance: } \Delta TC = TC_{\text{actual}} - TC_{\text{baseline}}$$
  $$\text{Optimality Condition: } \text{Holding Cost} \approx \text{Ordering Cost } \left(\frac{Q^*}{2}H = \frac{D}{Q^*}S\right)$$
- **Parameters & Sets**:
  - $\text{Budget} = \$102,000$
  - $\text{Actual Cost} = \$85,300$
  - Net Efficiency Gain $= +16.37\%$

#### Visual / Non-Textual Annotations
- **Confirmations (`✓`)**:
  - `✓ Holding Cost (\$42.5k) \approx Setup Cost (\$42.8k)` (Verified trade-off optimality).
  - `✓ Fill Rate SLA Met (99.4% \ge 98.0%)`.
- **Key Invariants (Boxed)**:
  - $\boxed{\text{Net Annual Savings: } \$16,700 \text{ / Year}}$
- **Instructor / Strategic Callouts**:
  - `🔍 Insight`: "When holding and ordering costs are roughly equal, the system operates near the global minimum of the convex cost curve."

---

### Slide 9: Executive Synthesis & Implementation Roadmap
- **Archetype**: `content_bullets` (Executive Summary)
- **Layout_Target**: Slide Index 9 (17 Slots)
- **Slide_Type**: Summary / Action Plan
- **Core_Concept**: Synthesize analytical findings into a 4-step governance and deployment checklist.

#### Layout Slots Allocation
- `Header_Area`: Executive Summary & Next Operational Steps
- `Card_1`: Policy Institutionalization (Adopt $Q^* = 1,200 \text{ units}$, $ROP = 350 \text{ units}$)
- `Card_2`: Vendor SLA Realignment (Negotiate lead time variance $L \in [3, 5] \text{ days}$)
- `Card_3`: Automated ERP Thresholds (Encode dynamic triggers into SAP/Oracle SCM)
- `Card_4`: Continuous Review Schedule (Quarterly parameter re-estimation for $H$ and $S$)
- `Footer_Callout`: Key Organizational Invariant & Milestone Target

#### Animation Sequence (Numbered Steps)
1. **Step 1 (`①`)**: Reveal core policy adoption parameters in Card 1.
2. **Step 2 (`②`)**: Cascade Cards 2, 3, and 4 (Operational Execution).
3. **Step 3 (`③`)**: Illuminate bottom governance milestone badge.

#### Mathematical / Core Structured Elements
- **Governing Policy Tuple**:
  $$\mathcal{P}^* = \langle Q^* = 1200, \; ROP = 350, \; SS = 50 \rangle$$
- **Parameters**:
  - Safety Stock ($SS$): $z_{\alpha} \cdot \sigma_L = 1.645 \cdot 30.4 \approx 50 \text{ units}$

#### Visual / Non-Textual Annotations
- **Confirmations (`✓`)**:
  - `✓ ERP Integration Approved`
  - `✓ Vendor Lead Time Bound Validated`
- **Key Invariants (Boxed)**:
  - $\boxed{\text{Target Go-Live: Q1 Execution | ROI Projected: 4.2x}}$
- **Instructor / Strategic Callouts**:
  - `📌 Key Theorem`: "Optimal mathematical solutions only succeed when paired with tight lead-time variance controls in the supply contract."
