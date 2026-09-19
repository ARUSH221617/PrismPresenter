# Slide Storyboard & Vision Parsing Specification Schema (`test-endpoint-detection-struct.md`)

This specification provides the operational schema, visual detection heuristics, and structured transcription format for an AI Vision & Parsing Agent to detect, extract, and serialize handwritten and hybrid mathematical lecture sheets (4-quadrant slide sheets) into production-ready slide decks.

---

## 1. Document Metadata Schema

Every physical or digital sheet processed by the vision pipeline must populate the following metadata block prior to quadrant extraction:

```yaml
Document_Metadata:
  Subject: "ریاضیات (Mathematics)"
  Grade_Level: "پایه هفتم (Grade 7)"
  Chapter_Number: 5
  Chapter_Title: "شمارنده‌ها و اعداد اول (Divisors and Prime Numbers)"
  Lesson_ID: 1
  Page_Number: 1
  Slide_Range: "Slides 1–4"
  Language_Direction: "RTL"
  Grid_Type: "4-Quadrant Landscape (2x2)"
```

---

## 2. Visual Layout & Quadrant Mapping Detection Rules

Physical pages are structured in a standard 2×2 grid following the Persian/Arabic right-to-left (RTL) reading order:

```
+-----------------------------+-----------------------------+
|                             |                             |
|       Quadrant_TL           |       Quadrant_TR           |
|       [Slide N + 1]         |       [Slide N]             |
|                             |                             |
+-----------------------------+-----------------------------+
|                             |                             |
|       Quadrant_BL           |       Quadrant_BR           |
|       [Slide N + 3]         |       [Slide N + 2]         |
|                             |                             |
+-----------------------------+-----------------------------+
```

### Quadrant Detection Heuristics:
1. **Quadrant_TR (Top-Right)**:
   - Coordinates: $X \in [0.50, 1.00]$, $Y \in [0.00, 0.50]$
   - Order Index: **Slide $N$** (First slide of the sheet)
2. **Quadrant_TL (Top-Left)**:
   - Coordinates: $X \in [0.00, 0.50]$, $Y \in [0.00, 0.50]$
   - Order Index: **Slide $N + 1$** (Second slide of the sheet)
3. **Quadrant_BR (Bottom-Right)**:
   - Coordinates: $X \in [0.50, 1.00]$, $Y \in [0.50, 1.00]$
   - Order Index: **Slide $N + 2$** (Third slide of the sheet)
4. **Quadrant_BL (Bottom-Left)**:
   - Coordinates: $X \in [0.00, 0.50]$, $Y \in [0.50, 1.00]$
   - Order Index: **Slide $N + 3$** (Fourth slide of the sheet)

---

## 3. Slide Content Extraction Schema

For each detected quadrant, extract data strictly adhering to the standardized block below:

```markdown
### Slide [Slide_Number]: [Slide Title / عنوان اسلاید]
- **Quadrant**: [TR | TL | BR | BL]
- **Slide_Type**: [Agenda | Theory / Definition | Classification | Worked Example | Exercise | Summary]
- **Core_Concept**: [Brief 1-line summary of pedagogical target]

#### Animation Sequence (Numbered Steps)
Strictly follow circled visual step markers (①, ②, ③, ④, ⑤...):
1. **Step 1 (`①`)**: [Text / Prompt / Initial problem state]
2. **Step 2 (`②`)**: [Intermediate deduction / Operational breakdown]
3. **Step 3 (`③`)**: [Secondary proof / Evaluation]
4. **Step 4 (`④`)**: [Final answer / Boxed rule / Generalization]

#### Mathematical Elements
- **Definitions / Rules**: [Formal statement in Persian/English]
- **Formulas / Equations**: [Standard LaTeX formatting, inline $...$ or display $$...$$]
- **Factorizations / Sets**: [Set notations, divisor lists, e.g., $D_{24} = \{1, 2, 3, 4, 6, 8, 12, 24\}$]

#### Visual / Non-Textual Annotations
- **Trees / Diagrams**:
  - `Root`: [Node value / Category]
  - `Branches`: [Left child, Right child, sub-branches]
- **Eliminations (Crossed Out)**: [Strikethrough elements with reason, e.g., `\cancel{9}` (عدم اول بودن / composite)]
- **Confirmations (Checkmarked)**: [Validated items marked with `✓`]
- **Key Invariants (Boxed)**: [Formulas or values inside rectangular frames `\boxed{...}`]
- **Teacher Callouts**: [Marginal notes prefixed with `نکته`, `اگه دقت کنی!`, `هشدار`, `توجه`]
```

---

## 4. Concrete Transcription Example

Direct transcription modeled from **Sheet 1 (Prime Numbers & Worked Examples)** representing Slides 1 through 4.

---

### Slide 1: مفهوم شمارنده و تعریف عدد اول (Concept of Divisors & Prime Numbers)
- **Quadrant**: TR
- **Slide_Type**: Theory / Definition
- **Core_Concept**: تعریف مفهوم شمارنده طبیعی و مشخص کردن ویژگی اعداد اول با دو شمارنده متمایز

#### Animation Sequence (Numbered Steps)
1. **Step 1 (`①`)**: بررسی ضرب‌های تولیدکننده عدد ۶ و تعیین تمام شمارنده‌ها:
   $$6 = 1 \times 6 = 2 \times 3 \implies \text{شمارنده‌های } 6 = \{1, 2, 3, 6\}$$
2. **Step 2 (`②`)**: شمارش تعداد مقسوم‌علیه‌ها: عدد ۶ دارای ۴ شمارنده مختلف است، پس عدد اول نیست.
3. **Step 3 (`③`)**: بررسی عدد ۷:
   $$7 = 1 \times 7 \implies \text{شمارنده‌های } 7 = \{1, 7\}$$
4. **Step 4 (`④`)**: استخراج تعریف رسمی: هر عدد طبیعی بزرگتر از ۱ که **دقیقاً دو شمارنده متمایز** (یک و خودش) داشته باشد، **عدد اول** است.

#### Mathematical Elements
- **Definitions / Rules**: عدد طبیعی $p > 1$ اول است اگر و تنها اگر تنها مقسوم‌علیه‌های مثبت آن $1$ و $p$ باشند.
- **Formulas / Equations**:
  $$\text{Divisors}(p) = \{1, p\} \quad \text{where } p \in \mathbb{N}, p > 1$$
- **Factorizations / Sets**:
  - $D_6 = \{1, 2, 3, 6\}$ (۴ شمارنده)
  - $D_7 = \{1, 7\}$ (۲ شمارنده)

#### Visual / Non-Textual Annotations
- **Trees / Diagrams**: ندارد.
- **Eliminations (Crossed Out)**: ندارد.
- **Confirmations (Checkmarked)**: $7 \implies \checkmark$ (عدد اول است).
- **Key Invariants (Boxed)**:
  $$\boxed{\text{عدد اول} = \text{دقیقاً دو شمارنده متمایز}}$$
- **Teacher Callouts**:
  - `اگه دقت کنی!`: عدد ۱ فقط **یک** شمارنده دارد ($1 \times 1 = 1$)، بنابراین عدد ۱ **نه اول است و نه مرکب**!

---

### Slide 2: دسته‌بندی اعداد طبیعی (Classification of Natural Numbers)
- **Quadrant**: TL
- **Slide_Type**: Classification
- **Core_Concept**: افراز مجموعه اعداد طبیعی به سه دسته بر اساس تعداد شمارنده‌ها

#### Animation Sequence (Numbered Steps)
1. **Step 1 (`①`)**: دسته‌بندی اعداد طبیعی به سه گروه مجزا: عدد ۱، اعداد اول، اعداد مرکب.
2. **Step 2 (`②`)**: تعریف اعداد مرکب: اعدادی که بیش از دو شمارنده طبیعی دارند (حداقل ۳ شمارنده).
3. **Step 3 (`③`)**: بررسی عدد ۲ به عنوان تنها عدد اول زوج:
   $$2 = 1 \times 2 \implies D_2 = \{1, 2\}$$
4. **Step 4 (`④`)**: نتیجه‌گیری مهم در خصوص تمامی اعداد زوج دیگر ($4, 6, 8, \dots$) که همگی بر ۲ بخش‌پذیرند و مرکب هستند.

#### Mathematical Elements
- **Definitions / Rules**:
  $$\mathbb{N} = \{1\} \cup \{\text{اعداد اول}\} \cup \{\text{اعداد مرکب}\}$$
- **Formulas / Equations**:
  $$\forall n \in \text{اعداد مرکب} \implies |D_n| \ge 3$$

#### Visual / Non-Textual Annotations
- **Trees / Diagrams**:
  - `Root`: اعداد طبیعی ($\mathbb{N}$)
  - `Branches`:
    - شاخه ۱: عدد ۱ (تک شمارنده)
    - شاخه ۲: اعداد اول (دو شمارنده: $2, 3, 5, 7, 11, \dots$)
    - شاخه ۳: اعداد مرکب (بیش از دو شمارنده: $4, 6, 8, 9, \dots$)
- **Eliminations (Crossed Out)**: $\cancel{1 \in \text{اعداد اول}}$ (عدد ۱ اول نیست).
- **Confirmations (Checkmarked)**:
  - $2 \implies \checkmark$ (تنها عدد اول و زوج).
- **Key Invariants (Boxed)**:
  $$\boxed{\text{تنها عدد اول زوج} = 2}$$
- **Teacher Callouts**:
  - `نکته طلایی`: هر عدد زوج بزرگتر از ۲، حتماً یک عدد مرکب است چون حداقل ۳ شمارنده ($1$ و $2$ و خودش) دارد.

---

### Slide 3: مثال‌های کاربردی و بررسی اول بودن (Worked Examples: Primality Test)
- **Quadrant**: BR
- **Slide_Type**: Worked Example
- **Core_Concept**: بررسی اول یا مرکب بودن اعداد با استفاده از ضرب و شمارش شمارنده‌ها

#### Animation Sequence (Numbered Steps)
1. **Step 1 (`①`)**: صورت سوال: آیا اعداد $21$، $29$ و $91$ اول هستند یا مرکب؟
2. **Step 2 (`②`)**: تحلیل عدد ۲۱:
   $$21 = 3 \times 7 \implies D_{21} = \{1, 3, 7, 21\} \implies \text{مرکب}$$
3. **Step 3 (`③`)**: تحلیل عدد ۲۹:
   $$29 = 1 \times 29 \implies D_{29} = \{1, 29\} \implies \text{اول}$$
4. **Step 4 (`④`)**: تحلیل دام‌دار عدد ۹۱ (آزمون بخش‌پذیری):
   $$91 = 7 \times 13 \implies \text{مرکب}$$

#### Mathematical Elements
- **Formulas / Equations**:
  - $21 = 3 \times 7$
  - $29 = 1 \times 29$
  - $91 = 7 \times 13$
- **Factorizations / Sets**:
  - $D_{21} = \{1, 3, 7, 21\}$
  - $D_{29} = \{1, 29\}$
  - $D_{91} = \{1, 7, 13, 91\}$

#### Visual / Non-Textual Annotations
- **Trees / Diagrams**: ندارد.
- **Eliminations (Crossed Out)**: $\cancel{91 \text{ اول است}}$ (تصور اشتباه رد می‌شود).
- **Confirmations (Checkmarked)**:
  - $21 \rightarrow \text{مرکب} \ \checkmark$
  - $29 \rightarrow \text{اول} \ \checkmark$
  - $91 \rightarrow \text{مرکب} \ \checkmark$
- **Key Invariants (Boxed)**:
  $$\boxed{91 = 7 \times 13 \implies \text{مرکب}}$$
- **Teacher Callouts**:
  - `هشدار امتحانی!`: عدد ۹۱ در ظاهر شبیه به اعداد اول است، اما چون بر ۷ و ۱۳ بخش‌پذیر است، حتماً مرکب می‌باشد!

---

### Slide 4: تجزیه به شمارنده‌های اول و نمودار درختی (Tree Diagram Factorization)
- **Quadrant**: BL
- **Slide_Type**: Worked Example / Method
- **Core_Concept**: نمایش تجزیه عدد به عوامل اول با استفاده از رسم نمودار درختی

#### Animation Sequence (Numbered Steps)
1. **Step 1 (`①`)**: نوشتن عدد مورد نظر در رأس درخت (مثال: عدد ۶۰).
2. **Step 2 (`②`)**: تجزیه عدد به دو شمارنده اختیاری (مثلاً $6 \times 10$).
3. **Step 3 (`③`)**: ادامه شاخه‌ها تا رسیدن به اعداد اول نهایی:
   $$6 \rightarrow 2 \times 3, \quad 10 \rightarrow 2 \times 5$$
4. **Step 4 (`④`)**: نوشتن عدد به فرم حاصل‌ضرب شمارنده‌های اول و نمایش توان‌دار:
   $$60 = 2 \times 2 \times 3 \times 5 = 2^2 \times 3 \times 5$$

#### Mathematical Elements
- **Formulas / Equations**:
  $$60 = 6 \times 10 = (2 \times 3) \times (2 \times 5) = 2^2 \times 3 \times 5$$
- **Factorizations / Sets**:
  - شمارنده‌های اول عدد ۶۰: $\{2, 3, 5\}$

#### Visual / Non-Textual Annotations
- **Trees / Diagrams**:
  - `Root`: $60$
  - `Branches`:
    - شاخه چپ: $6 \longrightarrow 2 \text{ (اول)} \text{ و } 3 \text{ (اول)}$
    - شاخه راست: $10 \longrightarrow 2 \text{ (اول)} \text{ و } 5 \text{ (اول)}$
- **Eliminations (Crossed Out)**: ندارد.
- **Confirmations (Checkmarked)**: دور برگ‌های درخت دایره کشیده شده که نشانه رسیدن به عامل اول است ($\text{Circle: } 2, 3, 2, 5$).
- **Key Invariants (Boxed)**:
  $$\boxed{60 = 2^2 \times 3 \times 5}$$
- **Teacher Callouts**:
  - `اگه دقت کنی!`: مهم نیست تجزیه را با چه ضربی شروع کنی (مثلاً $60 = 4 \times 15$ یا $2 \times 30$)، انتهای شاخه‌ها همیشه به عوامل اول یکسان ختم خواهد شد!
