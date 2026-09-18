# Slide Storyboard Specification Schema (`structure.md`)

## 1. Document Metadata
- **Grade_Level**: [e.g., 7 / پایه هفتم]
- **Chapter**: [e.g., 5 / شمارنده‌ها و اعداد اول]
- **Lesson_ID**: [e.g., 1]
- **Page_Number**: [Integer]
- **Slide_Range**: [e.g., Slides 1–4]

---

## 2. Quadrant Layout Mapping
Each physical page is processed as four slides in reading order:
- `Quadrant_TR` (Top-Right): Slide $N$
- `Quadrant_TL` (Top-Left): Slide $N+1$
- `Quadrant_BR` (Bottom-Right): Slide $N+2$
- `Quadrant_BL` (Bottom-Left): Slide $N+3$

---

## 3. Slide Content Schema

For each quadrant, extract data using the following structure:

### Slide [Slide_Number]: [Slide Title]
- **Quadrant**: [TR | TL | BR | BL]
- **Slide_Type**: [Agenda | Theory / Definition | Classification | Worked Example | Exercise]
- **Core_Concept**: [Brief 1-line summary]

#### Animation Sequence (Numbered Steps)
Transcribe elements strictly following circled step indicators (`①`, `②`, `③`...):
1. **Step 1 (`①`)**: [Text / Formula / Statement]
2. **Step 2 (`②`)**: [Text / Formula / Statement]
3. **Step 3 (`③`)**: [Text / Formula / Statement]

#### Mathematical Elements
- **Definitions / Rules**: [Formal mathematical rule stated]
- **Formulas / Equations**: [Standard LaTeX formatting, e.g., $a \times b = 62$]
- **Factorizations / Sets**: [e.g., Factors of $24 = \{1, 2, 3, 4, 6, 8, 12, 24\}$]

#### Visual / Non-Textual Annotations
- **Trees / Diagrams**:
  - `Root`: [Number or Category]
  - `Branches`: [Branch 1, Branch 2]
- **Eliminations (Crossed Out)**: [List terms crossed out and reason, e.g., $25$ eliminated from primes list]
- **Confirmations (Checkmarked)**: [List terms validated with `✓`]
- **Key Invariants (Boxed)**: [List boxed final answers or invariants, e.g., boxed 97]
- **Teacher Callouts**: [Notes prefixed with `اگه دقت کنی!`, `نکته`, etc.]

---

## 4. Transcription Example (Applied to Sheet 1, Slide 4)

### Slide 4: تعریف و ویژگی اعداد اول
- **Quadrant**: BL
- **Slide_Type**: Theory & Worked Observation
- **Core_Concept**: Definition of prime numbers, listing primes $< 50$, and the parity invariant of the number 2.

#### Animation Sequence
1. **Step 1 (`①`)**: هر عدد طبیعی که فقط ۲ شمارنده داشته باشد که یک و خودش هستند.
   - List: $2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47$
2. **Step 2 (`②`)**: اگه دقت کنی! تنها عدد اول زوج ۲ است. بقیه اعداد اول فرد هستند.
3. **Step 3 (`③`)**: از اونجایی که اگه جمع و تفریق دو عدد فرد بشه حتماً یکی از اونها زوجه و دیگری فرد.
4. **Step 4 (`④`)**: پس حتماً یکی از اون دو عدد ۲ هست.

#### Visual / Non-Textual Annotations
- **Eliminations**: Number `۲۵` crossed out in the prime sequence (corrected on the fly).
- **Key Invariants**: Box around `۲`.