# AI Coding Agent Instructions: Taste-Skill Integration

This workspace has integrated the **Taste-Skill** framework (`/taste-skill-main/`). All AI agents working in this repository must strictly adhere to the following directives to ensure high-end, premium, and functional UI/UX craftsmanship that is free of "AI slop" or default templates.

---

## 1. PRE-FLIGHT BRIEF INFERENCE
Before writing code or editing files, you **MUST** output a single-line "Design Read" at the very beginning of your turn:
> **"Reading this as: <page kind> for <audience>, with a <vibe> language, leaning toward <design system or aesthetic family>."**

*Example:*
> *"Reading this as: Professional payroll and timesheet management hub for internal operations and finance teams, with a minimal and elegant tabular language, leaning toward structured, high-contrast neubrutalism & responsive grid layouts."*

---

## 2. THE THREE DIALS
Ensure any layout or interface matches the configured parameters:
*   **`DESIGN_VARIANCE: 5`** - (Clean, balanced alignment, minimal clutter)
*   **`MOTION_INTENSITY: 4`** - (Calm, purposeful transitions, micro-interactions on active buttons)
*   **`VISUAL_DENSITY: 8`** - (Sleek, high-density dashboard cockpit for complex payroll processing)

---

## 3. CORE DESIGN Engineering Guidelines (Anti-Slop Rules)

### 3.A Typography Pairs
*   **Display & Title Fonts**: Sans-serif Display fonts like `Outfit`, `Geist`, `Inter` with snug letter-spacing (`tracking-tight` or `tracking-tighter`) and crisp weights.
*   **Data & Tables Font**: Sử dụng font được cấu hình (var(--font-table, var(--font-main))) kèm theo class tabular-nums for numbers, codes, dates, and currency values to align columns vertically perfectly.
*   **Emphasis**: Avoid mixing font-families within a single text block (e.g., no random serif word in a sans headline). Use italic or bold of the same family.

### 3.B Color Calibration
*   **Accent Color**: Maximum 1 primary accent. Saturation < 80%.
*   **Anti-Purple Gradient Glow**: Do NOT inject random violet/magenta glow or mesh gradients into professional business applications unless requested.
*   **High-Contrast Neutrals**: Stick to Slate/Zinc/Stone neutrals. Ensure WCAG AA contrast ratios (4.5:1 min for body text, 3:1 for large display elements) are respected.

### 3.C Shape & Spacing Consistency
*   **Corner Radius Scale**: Lock corner radiuses globally across the app (e.g., tables have `12px` radius, buttons have `pill` shape, containers have `1.25rem` radius). Never mix sharp and rounded elements randomly.
*   **Padding**: Avoid overly deep top paddings (max `pt-24` on main headers) or giant padding margins that waste screen space.

### 3.D Interactive States & Button Rules
*   **Tactile Active Feedbacks**: On `:active`, use a light physical push translation (`active:scale-[0.98]` or `active:translate-y-[1px]`).
*   **No Button Wrap**: Button text must always reside on a single line on desktop. Shorten button text (1-3 words max) and provide sufficient width.
*   **No Duplicate CTAs**: One primary call-to-action intent per view.
*   **Placeholder Restraint**: Never use placeholders as standard input labels. Keep labels above inputs, helper text below.

### 3.E Layout Discipline
*   **Navigation**: Keep navigation strictly on a single line at desktop width. Limit navigation bar height to 64-72px.
*   **Bento Cells**: Bento grids must have a cell count that matches the content precisely (no empty or placeholder tiles).
*   **Z-Pattern / Alternation**: Do not chain more than 2 consecutive alternating sections of the same pattern (e.g., left-image/right-text). Break up layout monotony.

---

## 4. CODE MAINTENANCE & SANITIZATION (The Logic Folder Rule)
Keep the timesheet and payroll core business rules separate from UI components.
*   **Zero Junk Code**: Commented out blocks of inactive code, temporary console.logs, or placeholder test arrays inside `/src/app/constants/timesheet-logic.ts` must be pruned or encapsulated cleanly.
*   **Helper Modularity**: Avoid embedding continuous states (like window scroll or mouse move) in React state. Use custom hooks, memoized helpers, and web workers to keep calculation logic completely distinct from layout rendering.

## Cloud workspace changes

Read [docs/cloud-workspace.md](docs/cloud-workspace.md) before changing synchronization, local persistence or the Save contract. It records the cross-device requirements, cache correctness boundary and verification limitations.
