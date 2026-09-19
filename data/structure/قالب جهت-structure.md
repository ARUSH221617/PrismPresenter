# Slide Storyboard Specification Schema (`قالب جهت-structure.md`)

## 1. Document Metadata
- **Grade_Level**: [e.g., پایه نهم / پایه دهم / K-10 / K-11]
- **Subject**: [e.g., ریاضیات / هندسه / فیزیک / STEM]
- **Chapter**: [e.g., فصل سوم: استدلال و اثبات در هندسه]
- **Lesson_ID**: [e.g., درس دوم: هم‌نهشتی مثلث‌ها]
- **Template_Name**: قالب جهت.pptx
- **Slide_Range**: [e.g., Slides 1–5]
- **Page_or_Sheet_Reference**: [Book Page Number or Worksheet Index]

---

## 2. Layout & Card Spatial Mapping
`قالب جهت.pptx` utilizes high-contrast instructional cards and directional flow indicators for progressive disclosure. The pedagogical grid organizes slides into modular zones and reading sequences:

- **Full_Cover** (`Index 0`): Title cover, subtitle, and author/grade banner.
- **Card_Split_L2R / R2L** (`Index 1`): Two-column setup (Left: Theory/Rule, Right: Conceptual diagram/Example).
- **Stepwise_Worked_Vertical** (`Index 2`): Problem statement header with progressive vertical derivation cards (Steps 1 to 4).
- **Dual_Comparative_Cards** (`Index 3`): Dual-card comparative container for Counter-examples, Definitions vs. Properties, or Case A vs. Case B.
- **Dense_Matrix_Grid** (`Index 4`): Multi-slot analytical layout (up to 19 granular slots) for comprehensive summaries, multi-step geometric proofs, or theorem grids.

---

## 3. Slide Content Schema

For each slide, construct the storyboard entry adhering strictly to the schema below:

### Slide [Slide_Number]: [Slide Title / Headline]
- **Archetype**: [title_cover | content_bullets | card_split | stepwise_worked | dual_comparative | dense_matrix]
- **Position / Layout_Slot**: [Full_Cover | Card_Left | Card_Right | Header_Statement | Derivation_Container]
- **Slide_Type**: [Agenda | Theory / Definition | Classification | Worked Example | Exercise | Summary & Cheatsheet]
- **Core_Concept**: [1-line pedagogical takeaway]

#### Animation Sequence (Numbered Progressive Reveal)
Transcribe elements strictly following circled numerical triggers (`①`, `②`, `③`, `④`...) corresponding to teacher lecture beats:
1. **Step 1 (`①`)**: [Trigger: Immediate / On Click] -> [Problem setup, Given statement, or Base Theorem]
2. **Step 2 (`②`)**: [Trigger: On Click] -> [First transformation, geometric deduction, or formula substitution]
3. **Step 3 (`③`)**: [Trigger: On Click] -> [Intermediate resolution or elimination step]
4. **Step 4 (`④`)**: [Trigger: On Click] -> [Final target answer, Q.E.D. closure, or boxed invariant]

#### Mathematical & Conceptual Elements
- **Definitions / Theorems**: [Exact formal statement or pedagogical rule]
- **Formulas / Equations**: [Standard LaTeX formatting, e.g., $\Delta ABC \cong \Delta A'B'C' \implies \frac{AB}{A'B'} = 1$]
- **Given & Goal (Hypothesis & Thesis)**:
  - `Hypothesis (فرض)`: [e.g., $AB = AC$, $AM \perp BC$]
  - `Thesis (حکم)`: [e.g., $BM = MC$]

#### Visual & Pedagogical Annotations
- **Directional Vectors / Geometric Diagrams**:
  - `Base Shape / Geometry`: [e.g., Isosceles Triangle $ABC$ with altitude $AM$]
  - `Markings`: [e.g., Right-angle square at $M$, double tick mark on $AB$ and $AC$]
- **Eliminations & Invalidations (Crossed Out)**: [List terms crossed out or falsified counter-examples, e.g., $\cancel{SAS}$ due to non-included angle]
- **Confirmations (Checkmarked)**: [Validated assertions with `✓`, e.g., حالت (ض‌زض) برقرار است ✓]
- **Key Invariants (Boxed / Highlighted)**: [Final values, conclusions, or boxed answers, e.g., $\boxed{x = 12}$]
- **Instructional Callouts (جهت / نکته)**: [Teacher callouts prefixed with `💡 جهت راهنما`, `⚠️ دام تستی`, or `📌 نکته کلیدی`]

---

## 4. Transcription Example (Applied to Lesson Module)

### Slide 1: فصل سوم: استدلال و اثبات در هندسه
- **Archetype**: title_cover
- **Position / Layout_Slot**: Full_Cover
- **Slide_Type**: Agenda
- **Core_Concept**: معرفی مبحث هم‌نهشتی مثلث‌ها و کاربرد آن در حل مسائل پایه نهم

---

### Slide 2: هم‌نهشتی به حالت (ض ز ض) — قضیه و تحلیل
- **Archetype**: content_bullets (Card_Split)
- **Position / Layout_Slot**: Card_Left (Theorem) + Card_Right (Diagram)
- **Slide_Type**: Theory / Definition
- **Core_Concept**: هرگاه دو ضلع و زاویه بین آن‌ها از یک مثلث با اجزای نظیر از مثلث دیگر برابر باشد، دو مثلث هم‌نهشت‌اند.

#### Animation Sequence
1. **Step 1 (`①`)**: تعریف رسمی حالت دو ضلع و زاویه بین (ض‌زض).
2. **Step 2 (`②`)**: نمایش نمادین زاویه محصور بین دو ضلع نشاندار شده.
3. **Step 3 (`③`)**: تأکید بر شرط «زاویه بین» و جلوگیری از اشتباه رایج ض‌ض‌ز.

#### Mathematical & Conceptual Elements
- **Definitions / Theorems**: قضیه هم‌نهشتی دو ضلع و زاویه محصور:
  $$\begin{cases} AB = A'B' \\ \hat{A} = \hat{A}' \\ AC = A'C' \end{cases} \implies \Delta ABC \cong \Delta A'B'C'$$
- **Given & Goal**:
  - `Hypothesis (فرض)`: $AB = A'B'$, $AC = A'C'$, $\hat{A} = \hat{A}'$
  - `Thesis (حکم)`: $\Delta ABC \cong \Delta A'B'C'$

#### Visual & Pedagogical Annotations
- **Directional Vectors / Geometric Diagrams**:
  - `Base Shape`: دو مثلث همنهشت $\Delta ABC$ و $\Delta A'B'C'$
  - `Markings`: زاویه $\hat{A}$ با کمان رنگی مشخص و اضلاع $AB$ و $AC$ دارای علامت‌گذاری نظیر.
- **Eliminations & Invalidations**: $\cancel{\text{SSA}}$ (حالت دو ضلع و زاویه غیرمحصور معتبر نیست).
- **Instructional Callouts**: `⚠️ دام تستی`: اگر زاویه دقیقا بین دو ضلع مساوی نباشد، نمی‌توان نتیجه هم‌نهشتی گرفت!

---

### Slide 3: حل گام‌به‌گام مسأله هم‌نهشتی (ارتفاع در مثلث متساوی‌الساقین)
- **Archetype**: content_bullets (Stepwise_Worked_Vertical)
- **Position / Layout_Slot**: Derivation_Container
- **Slide_Type**: Worked Example
- **Core_Concept**: اثبات برابری پاره‌خط‌های پایه حاصل از رسم ارتفاع در مثلث متساوی‌الساقین.

#### Animation Sequence
1. **Step 1 (`①`)**: تفکیک فرض و حکم مسأله و شماره‌گذاری زوایای قائمه مجاور.
2. **Step 2 (`②`)**: معرفی مثلث‌های قائم‌الزاویه ایجاد شده $\Delta ABH$ و $\Delta ACH$.
3. **Step 3 (`③`)**: تطبیق حالت وتر و یک ضلع مشترک (و‌ض).
4. **Step 4 (`④`)**: استخراج تساوی اجزای متناظر و اثبات نهایی حکم.

#### Mathematical & Conceptual Elements
- **Formulas / Equations**:
  $$\begin{cases} AH = AH & (\text{ضلع مشترک}) \\ AB = AC & (\text{فرض - ساق‌ها}) \\ \hat{H}_1 = \hat{H}_2 = 90^\circ & (\text{فرض - ارتفاع}) \end{cases} \implies \Delta ABH \cong \Delta ACH \quad (\text{حالت و‌ض})$$
- **Conclusion**:
  $$BH = CH \quad (\text{تساوی اجزای متناظر})$$

#### Visual & Pedagogical Annotations
- **Directional Vectors / Geometric Diagrams**:
  - `Base Shape`: مثلث متساوی‌الساقین با خط ارتفاع عمود بر قاعده $BC$.
- **Confirmations**: $\Delta ABH \cong \Delta ACH$ با برچسب `(و‌ض) ✓`
- **Key Invariants**: $\boxed{BH = CH}$
- **Instructional Callouts**: `💡 جهت راهنما`: ارتفاع وارد بر قاعده، همزمان میانه و نیم‌ساز نیز می‌باشد.
