---
# officecli: v1.0.41
name: officecli-word-form
description: "Use this skill to create fillable Word forms (.docx) with real Content Controls (SDT) and document protection. Triggers when the user needs: employee onboarding forms, HR intake forms, fillable contracts or SOW templates, customer surveys, compliance checklists, or any Word document where specific fields must be editable while the rest is locked. Key triggers: 'fillable form', 'form fields', 'content controls', 'SDT', 'word form', 'fill in', 'only editable fields', 'protect document', 'onboarding form', 'survey template'. Do NOT trigger for regular reports, letters, or documents that don't require user-fillable fields — those belong to officecli-docx."
---
# officecli: v1.0.41

# officecli-word-form

## BEFORE YOU START (CRITICAL)

**If `officecli` is not installed:**

`macOS / Linux`

```bash
if ! command -v officecli >/dev/null 2>&1; then
    curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
fi
```

`Windows (PowerShell)`

```powershell
if (-not (Get-Command officecli -ErrorAction SilentlyContinue)) {
    irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
}
```

Verify: `officecli --version`

**zsh path quoting — REQUIRED when path contains `[N]`:**

```bash
# WRONG — zsh interprets [1] as a glob pattern and fails
officecli get form.docx /body/sdt[1]

# CORRECT — always quote paths containing brackets
officecli get form.docx '/body/sdt[1]'
officecli get form.docx '/formfield[empName]'
officecli set form.docx '/body/sdt[2]' --prop text="Jane Smith"
```

---
# officecli: v1.0.41

## What Makes a Real Form

A **real fillable form** requires two things: SDT Content Controls (the fields) + `protection=forms` (the lock).

| Approach | What users see in Word | CLI-readable? | Is a real form? |
|----------|----------------------|---------------|-----------------|
| SDT controls + `protection=forms` | Gray-bordered fields, rest is locked | Yes — `query sdt` returns all fields | **YES** |
| Underscores `___` or spaces | Visual-only lines, entire doc editable | No — no structured fields | **NO** |
| Plain paragraphs left blank | Empty text, entire doc editable | No | **NO** |

**CRITICAL WARNING**: Do NOT simulate form fields with underscores (`姓名：_______________`) or blank lines. These produce no structured data, cannot be programmatically read or filled, and violate the Hard Rules of this skill. Always use `--type sdt` or `--type formfield` controls.

**How `protection=forms` works**: After running `officecli set form.docx / --prop protection=forms`, the document becomes partially locked in Word — users can only interact with SDT content controls and legacy formfields. All other paragraphs, headings, and table cells become read-only. The CLI itself is not restricted by document protection (see [Document Protection](#document-protection) for details).

---
# officecli: v1.0.41

## Execution Model

**Run commands one at a time. Read output before proceeding.**

OfficeCLI is incremental: every `add`, `set`, and `remove` immediately modifies the file. Use this to catch errors early:

1. **One command at a time.** Check the output before running the next command.
2. **Non-zero exit = stop and fix immediately.** Do not continue building on a broken state.
3. **`protection=forms` must be the LAST command executed.** This is the most critical form-specific rule:

> **CRITICAL ORDERING RULE**: Run all `add` commands first, then `set protection=forms` last. After protection is enabled, any subsequent `add` command (for SDT or formfield) will fail with `Exit 1: Document is protected (mode: forms)`. If you must add controls to a protected document, use `--force` flag — but the recommended practice is always: build the form first, protect it last.

```bash
# CORRECT ordering
officecli add form.docx /body --type sdt ...    # all add commands first
officecli add form.docx /body --type sdt ...
officecli set form.docx / --prop protection=forms   # LAST

# WRONG — adding SDT after protection (will fail without --force)
officecli set form.docx / --prop protection=forms
officecli add form.docx /body --type sdt ...    # Exit 1: Document is protected
```

**Modifying a protected form**: If you need to add controls to an already-protected document:

```bash
# Option A: Remove protection, add controls, re-enable
officecli set form.docx / --prop protection=none
officecli add form.docx /body --type sdt --prop sdtType=text ...
officecli set form.docx / --prop protection=forms

# Option B: Use --force to bypass protection (not recommended for bulk edits)
officecli add form.docx /body --type sdt --prop sdtType=text ... --force
```

---
# officecli: v1.0.41

## Core Workflow (Standard Form Creation)

The 5-step workflow for every fillable form:

```
1. create + open      → initialize document
2. set metadata       → title, font defaults
3. add structure      → headings, label paragraphs
4. add SDT controls   → all fillable fields (all add commands go here)
5. set protection     → protection=forms LAST
```

### Complete Example: Employee Onboarding Form

```bash
# Step 1: Create and open
officecli create onboarding.docx
officecli open onboarding.docx

# Step 2: Set document metadata and defaults
officecli set onboarding.docx / \
  --prop title="Employee Onboarding Form" \
  --prop docDefaults.font="Calibri" \
  --prop docDefaults.fontSize="12pt"

# Step 3: Document title and intro
officecli add onboarding.docx /body --type paragraph \
  --prop text="Employee Onboarding Form" \
  --prop style=Heading1 --prop size=20 --prop bold=true \
  --prop spaceBefore=0pt --prop spaceAfter=12pt

officecli add onboarding.docx /body --type paragraph \
  --prop text="Please complete all fields below and return to HR on your first day. Fields marked with * are required." \
  --prop size=11 --prop italic=true --prop color=666666 --prop spaceAfter=18pt

# ── Section 1: Basic Information ──────────────────────────────────
officecli add onboarding.docx /body --type paragraph \
  --prop text="Section 1: Basic Information" \
  --prop style=Heading2 --prop size=14 --prop bold=true \
  --prop spaceBefore=18pt --prop spaceAfter=8pt

# Full Name (text SDT)
officecli add onboarding.docx /body --type paragraph \
  --prop text="Full Name: *" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add onboarding.docx /body --type sdt \
  --prop sdtType=text --prop alias="Full Name" --prop tag="full_name" \
  --prop text="Enter full name" --prop lock=sdtlocked

# Employee ID (text SDT)
officecli add onboarding.docx /body --type paragraph \
  --prop text="Employee ID:" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add onboarding.docx /body --type sdt \
  --prop sdtType=text --prop alias="Employee ID" --prop tag="employee_id" \
  --prop text="Assigned by HR" --prop lock=sdtlocked

# Department (dropdown SDT)
officecli add onboarding.docx /body --type paragraph \
  --prop text="Department: *" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add onboarding.docx /body --type sdt \
  --prop sdtType=dropdown --prop alias="Department" --prop tag="dept" \
  --prop items="Engineering,Finance,HR,Legal,Marketing,Operations,Sales" \
  --prop lock=sdtlocked

# Start Date (date SDT)
officecli add onboarding.docx /body --type paragraph \
  --prop text="Start Date: *" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add onboarding.docx /body --type sdt \
  --prop sdtType=date --prop alias="Start Date" --prop tag="start_date" \
  --prop format="yyyy年MM月dd日" --prop lock=sdtlocked

# ── Section 2: Emergency Contact ──────────────────────────────────
officecli add onboarding.docx /body --type paragraph \
  --prop text="Section 2: Emergency Contact" \
  --prop style=Heading2 --prop size=14 --prop bold=true \
  --prop spaceBefore=18pt --prop spaceAfter=8pt

officecli add onboarding.docx /body --type paragraph \
  --prop text="Emergency Contact Name:" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add onboarding.docx /body --type sdt \
  --prop sdtType=text --prop alias="Emergency Contact Name" --prop tag="emergency_name" \
  --prop text="Enter contact name" --prop lock=sdtlocked

officecli add onboarding.docx /body --type paragraph \
  --prop text="Relationship:" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add onboarding.docx /body --type sdt \
  --prop sdtType=text --prop alias="Relationship" --prop tag="emergency_rel" \
  --prop text="e.g. Spouse, Parent, Sibling" --prop lock=sdtlocked

# ── Section 3: Additional Notes ───────────────────────────────────
officecli add onboarding.docx /body --type paragraph \
  --prop text="Section 3: Additional Notes" \
  --prop style=Heading2 --prop size=14 --prop bold=true \
  --prop spaceBefore=18pt --prop spaceAfter=8pt

officecli add onboarding.docx /body --type paragraph \
  --prop text="Additional Notes:" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add onboarding.docx /body --type sdt \
  --prop sdtType=richtext --prop alias="Additional Notes" --prop tag="notes" \
  --prop text="Enter any additional information here"

# Step 4 (LAST): Enable form protection
officecli set onboarding.docx / --prop protection=forms

# Step 5: Close and validate
officecli close onboarding.docx
officecli validate onboarding.docx
```

---
# officecli: v1.0.41

## SDT Content Controls

### Control Types & Syntax

All 5 supported SDT types with verified syntax:

```bash
# TEXT — short text input (name, ID, phone, email)
officecli add form.docx /body --type sdt \
  --prop sdtType=text \
  --prop alias="Full Name" --prop tag="full_name" \
  --prop text="Enter full name" \
  --prop lock=sdtlocked

# RICHTEXT — long text, multi-paragraph (feedback, notes, descriptions)
# NOTE: richtext intentionally omits lock=sdtlocked — richtext fields are for open-ended user input
# and restricting them (sdtlocked still allows editing but is redundant for richtext semantics).
# If you need to prevent deletion of the richtext control, you may add lock=sdtlocked,
# but most forms leave richtext unlocked to allow free multi-paragraph input.
officecli add form.docx /body --type sdt \
  --prop sdtType=richtext \
  --prop alias="Additional Notes" --prop tag="notes" \
  --prop text="Enter comments here"

# DROPDOWN — fixed option list (department, country, status)
officecli add form.docx /body --type sdt \
  --prop sdtType=dropdown \
  --prop alias="Department" --prop tag="dept" \
  --prop items="Engineering,Finance,HR,Legal,Marketing,Operations,Sales" \
  --prop lock=sdtlocked

# COMBOBOX — predefined options + allow custom input
officecli add form.docx /body --type sdt \
  --prop sdtType=combobox \
  --prop alias="Preferred Feature" --prop tag="feature_pref" \
  --prop items="Data Export,API Integration,Reporting,Team Collaboration,Other" \
  --prop lock=sdtlocked

# DATE — Chinese/CJK format (use for Chinese-language forms)
officecli add form.docx /body --type sdt \
  --prop sdtType=date \
  --prop alias="入职日期" --prop tag="start_date" \
  --prop format="yyyy年MM月dd日" \
  --prop lock=sdtlocked

# DATE — International ISO format (use for English or bilingual forms)
officecli add form.docx /body --type sdt \
  --prop sdtType=date \
  --prop alias="Start Date" --prop tag="start_date" \
  --prop format="yyyy-MM-dd" \
  --prop lock=sdtlocked

# DATE — US English format
officecli add form.docx /body --type sdt \
  --prop sdtType=date \
  --prop alias="Contract Date" --prop tag="contract_date" \
  --prop format="MM/dd/yyyy" \
  --prop lock=sdtlocked
```

**Type aliases** (all verified):

| Canonical type | Accepted aliases |
|----------------|-----------------|
| `text` | (default if omitted) |
| `richtext` | `rich` |
| `dropdown` | `dropdownlist` |
| `combobox` | `combo` |
| `date` | `datepicker` |

**WARNING — checkbox SDT does not exist:**

```bash
# WRONG — silently degrades to text; get returns sdtType: text; no checkbox appears in Word
officecli add form.docx /body --type sdt --prop sdtType=checkbox ...

# CORRECT — use Legacy FormField for checkboxes
officecli add form.docx /body --type formfield --prop formfieldtype=checkbox \
  --prop name="agree_terms" --prop checked=false
```

`--prop sdtType=checkbox` is accepted without error but creates a plain text SDT — the checkbox is silently lost. This will cause Hard Rule failure. See [Legacy FormField](#legacy-formfield) for the correct checkbox pattern.

### items — Comma-Separated Format (Required)

```bash
# CORRECT — English commas, no semicolons
--prop items="Engineering,Finance,HR,Marketing,Sales"

# WRONG — semicolons are not supported
--prop items="Engineering;Finance;HR"

# NOTE: Items with internal commas will be split incorrectly — avoid commas in option text
# e.g., "Research & Development" is safe; "R&D, Labs" would split into two items
```

Items are trimmed of leading/trailing spaces automatically. `DisplayText` and `Value` are set to the same string.

### lock Values — Must Be All Lowercase

```bash
# CORRECT (all lowercase)
--prop lock=sdtlocked          # User can fill, cannot delete the control
--prop lock=contentlocked      # Content is read-only, but control can be deleted
--prop lock=sdtcontentlocked   # Both content and control structure are locked (alias: both)
# Omit lock entirely for unlocked (default)

# WRONG — camelCase inputs are NOT accepted
--prop lock=sdtLocked          # Error: Invalid lock value
--prop lock=contentLocked      # Error: Invalid lock value
```

**Note on get output**: When you read back a locked SDT with `get` or `query`, the lock value is returned in camelCase (`sdtLocked`, `contentLocked`, `sdtContentLocked`). This is display-only — the input value must still be lowercase.

**Recommended default**: Use `lock=sdtlocked` for most form fields — users can fill in values but cannot accidentally delete the control structure in Word.

### Lock Reference Table

| lock value (input) | get output | User can edit content | User can delete control |
|--------------------|------------|----------------------|------------------------|
| (omit) / `unlocked` | no lock field | Yes | Yes |
| `sdtlocked` | `sdtLocked` | Yes | No |
| `contentlocked` | `contentLocked` | No | Yes |
| `sdtcontentlocked` / `both` | `sdtContentLocked` | No | No |

### Required Properties for Every SDT

Every SDT control **must** have both `alias` and `tag`. This is a Hard Rule.

```bash
# MISSING alias/tag — Hard Rule violation
officecli add form.docx /body --type sdt --prop sdtType=text --prop text="Enter name"

# CORRECT — both alias and tag present
officecli add form.docx /body --type sdt \
  --prop sdtType=text \
  --prop alias="Full Name" \   # Human-readable label (use document language)
  --prop tag="full_name" \     # Program identifier (use lowercase_underscore)
  --prop text="Enter full name"
```

**Naming conventions**:
- `alias`: User-readable label. Match the document language (Chinese form → Chinese alias, English form → English alias). Example: `"姓名"`, `"Full Name"`, `"Department"`
- `tag`: Program identifier for AI agents and automation. Use lowercase with underscores. Example: `"full_name"`, `"dept"`, `"start_date"`, `"emergency_contact_name"`

### set SDT — What Can Be Modified After Creation

```bash
# These properties CAN be modified after creation
officecli set form.docx '/body/sdt[1]' --prop alias="New Label"
officecli set form.docx '/body/sdt[1]' --prop tag="new_tag"
officecli set form.docx '/body/sdt[1]' --prop lock=sdtlocked
officecli set form.docx '/body/sdt[1]' --prop text="New default text"

# These properties CANNOT be modified — exit 2: UNSUPPORTED
# items, format, sdtType — these are read-only after creation
# To change them: remove the SDT and re-add with new values
```

**Path forms for set/get**:
- `/body/sdt[N]` — by sequential index (1-based)
- `/body/sdt[@sdtId=N]` — by stable sdtId (survives reordering)
- `/body/p[@paraId=XXXXXXXX]/sdt[@sdtId=N]` — inline SDT inside a paragraph

---
# officecli: v1.0.41

## Legacy FormField

Use Legacy FormField **only when you need a checkbox**. For text and dropdown fields, prefer SDT controls — they offer more features (lock, alias, tag).

### When to Use FormField

| Need | Use |
|------|-----|
| Checkbox (agree to terms, conditional selection) | `formfield --prop formfieldtype=checkbox` |
| Simple text input (backward compatibility) | `formfield --prop formfieldtype=text` |
| Fixed dropdown (legacy Word compatibility) | `formfield --prop formfieldtype=dropdown` |

### FormField Syntax (All Types)

```bash
# CHECKBOX — the only real checkbox available in officecli
officecli add form.docx /body --type formfield \
  --prop formfieldtype=checkbox \
  --prop name="agree_terms" \
  --prop checked=false

# TEXT formfield
officecli add form.docx /body --type formfield \
  --prop formfieldtype=text \
  --prop name="emp_name" \
  --prop text="Enter name" \
  --prop maxlength=100

# DROPDOWN formfield
officecli add form.docx /body --type formfield \
  --prop formfieldtype=dropdown \
  --prop name="dept_select" \
  --prop items="Engineering,Sales,Marketing"
```

### Reading and Writing FormField Values

```bash
# Get by index (1-based)
officecli get form.docx '/formfield[1]'

# Get by name (recommended — more stable than index)
officecli get form.docx '/formfield[agree_terms]'

# Set checkbox state
officecli set form.docx '/formfield[agree_terms]' --prop checked=true

# Set text value
officecli set form.docx '/formfield[emp_name]' --prop text="Jane Smith"

# Set dropdown by text name
officecli set form.docx '/formfield[dept_select]' --prop text="Engineering"

# Set dropdown by 0-based index
officecli set form.docx '/formfield[dept_select]' --prop text="1"
```

### FormField Path System

FormField paths are separate from SDT paths:
- FormField: `/formfield[N]` or `/formfield[fieldName]`
- SDT: `/body/sdt[N]` or `/body/sdt[@sdtId=N]`

Both systems coexist in the same document. `protection=forms` locks the document for both SDT and formfield controls.

**get formfield returns**:

| Field | Description |
|-------|-------------|
| `name` | Field identifier |
| `formfieldType` | `text`, `checkbox`, or `dropdown` |
| `checked` | `True`/`False` (checkbox only) |
| `items` | Comma-separated options (dropdown only) |
| `default` | Selected index, 0-based (dropdown only) |
| `editable` | `True`/`False` based on protection state |

---
# officecli: v1.0.41

## Document Protection

### Enabling Form Protection

```bash
# Enable protection — run this LAST after all add commands
officecli set form.docx / --prop protection=forms

# Verify protection is active
officecli get form.docx /
# Output includes:
#   protection: forms
#   protectionEnforced: True
```

### Protection Modes

| Mode | What users can do in Word | CLI behavior |
|------|--------------------------|--------------|
| `forms` | Fill SDT and formfield controls only | `set`/`get` unaffected; `add` requires `--force` |
| `readOnly` | Read only — no editing at all | `set`/`get` unaffected; `add` requires `--force` |
| `comments` | Add comments only | CLI unaffected |
| `trackedChanges` | Edit with tracked changes only | CLI unaffected |
| `none` | Full editing (no protection) | CLI unaffected |

**KEY INSIGHT**: Document protection restricts what **Word users** can do — it does NOT prevent CLI commands from executing. The CLI can always `set`, `get`, and (with `--force`) `add` regardless of protection mode. This means you can programmatically fill a protected form via CLI even though the document is locked to Word users.

**Exception**: `add` commands (adding new controls) fail under any protection mode unless `--force` is specified.

### Protection and Lock Interaction (forms mode)

| lock value | `protection=forms` | User can edit? |
|------------|-------------------|---------------|
| `sdtlocked` | active | Yes — sdtlocked only prevents deleting the control; content is still editable |
| `contentlocked` | active | No — content is locked regardless of protection |
| `sdtcontentlocked` | active | No |
| (no lock) | active | Yes — unlocked SDT in forms mode is fully editable |
| any | `readOnly` | No — readOnly overrides all SDT locks |

### Temporarily Unlocking for Edits

```bash
# Remove protection to add new controls
officecli set form.docx / --prop protection=none

# Make your changes
officecli add form.docx /body --type sdt --prop sdtType=text ...

# Re-enable protection
officecli set form.docx / --prop protection=forms
```

---
# officecli: v1.0.41

## Watermark (Optional Enhancement)

Add watermarks to contracts, SOWs, or draft forms that require status labels.

```bash
# Red CONFIDENTIAL watermark (contracts and SOWs)
officecli add form.docx / --type watermark \
  --prop text="CONFIDENTIAL" --prop color=FF0000 \
  --prop opacity=0.3 --prop rotation=315

# Default DRAFT watermark
officecli add form.docx / --type watermark

# Custom watermark (orange, angled)
officecli add form.docx / --type watermark \
  --prop text="FOR REVIEW" --prop color=FFA500 \
  --prop font="Times New Roman" --prop opacity=0.4 --prop rotation=45

# Update watermark text and color (text, color, font, opacity, rotation are modifiable)
officecli set form.docx /watermark --prop text="APPROVED" --prop color=00AA00

# Check watermark
officecli get form.docx /watermark

# Remove watermark
officecli remove form.docx /watermark
```

**Watermark parent path**: Always use `/` (document root) — not `/body`.

**Modifying size/width/height**: These properties work only in `add`, not `set`. To resize a watermark:
```bash
officecli remove form.docx /watermark
officecli add form.docx / --type watermark --prop text="CONFIDENTIAL" --prop width=500pt --prop height=250pt
```

---
# officecli: v1.0.41

## Reading & Verifying Forms

### Query All SDT Controls

```bash
# List every SDT in the document with all properties
officecli query form.docx "sdt"
# Returns: path, alias, tag, sdtType, lock, editable, items/format (where applicable)
```

### Verify Protection Status

```bash
# Check document-level settings including protection
officecli get form.docx /
# Look for:
#   protection: forms
#   protectionEnforced: True
```

### Inspect Individual Controls

```bash
# Get specific SDT by index
officecli get form.docx '/body/sdt[1]'
# Returns: alias, tag, id, sdtType, lock, editable, and type-specific fields

# Get specific SDT by stable sdtId (survives document reordering)
officecli get form.docx '/body/sdt[@sdtId=3]'

# Get formfield by name
officecli get form.docx '/formfield[agree_terms]'

# Verify editable state of a specific control
# editable: True  → user can fill this field in Word
# editable: False → field is locked (contentlocked or document is readOnly)
```

### Document Verification

```bash
# Full content view
officecli view form.docx text

# Structure overview (shows heading hierarchy and control count)
officecli view form.docx outline

# Validate document structure
officecli validate form.docx
```

**Note on `view text` output for SDT controls**: `dropdown` and `date` controls that have no value selected will appear as `[sdt:alias] ` (empty placeholder) in `view text` output. This is expected — the control exists and is correctly created; it has simply not been filled yet. Use `officecli query form.docx "sdt"` to verify control existence and properties independently of fill state.

---
# officecli: v1.0.41

## Design Principles (Form-Specific)

### Control Type Decision Tree

```
Need user input?
├── Yes, date → sdtType=date (Chinese: format="yyyy年MM月dd日"; International: format="yyyy-MM-dd" or "MM/dd/yyyy")
├── Yes, select from fixed list → sdtType=dropdown
├── Yes, select or type custom → sdtType=combobox
├── Yes, short text (1 line) → sdtType=text
├── Yes, long text (paragraph) → sdtType=richtext
└── Yes, boolean (check/uncheck) → formfield formfieldtype=checkbox
```

### Typography Scale for Forms

> **⚠️ CRITICAL — Spacing Unit**: `spaceBefore`, `spaceAfter`, `spaceLine` 参数默认单位是 **twips**（1/20 pt）。
> 必须明确带 `pt` 后缀才能按 pt 解析：`spaceBefore=18pt` = 18pt，`spaceBefore=18` = 0.9pt。
> 本 skill 所有示例均使用 `pt` 后缀。

| Element | Size | Style | Notes |
|---------|------|-------|-------|
| Form title (H1) | 20pt | Bold | `spaceBefore=0pt`, `spaceAfter=12pt` |
| Section heading (H2) | 14pt | Bold | `spaceBefore=18pt`, `spaceAfter=8pt` |
| Field label | 11pt | Bold | `spaceAfter=4pt` |
| SDT control content | 11pt | Normal | Inherits document default |
| Instructions/notes | 11pt | Italic, color=666666 | `spaceAfter=18pt` |

> **NOTE**: For Chinese/CJK forms, use `docDefaults.font="Microsoft YaHei"` (微软雅黑) instead of Calibri. Setting `docDefaults.font="Calibri"` applies Calibri as the East Asian (CJK) font as well — Calibri lacks Chinese glyphs and will render poorly in Chinese-language documents. Use `Calibri` only for Latin/ASCII content or English-only forms.

### Form Layout Pattern

```
Form Title (H1, 20pt bold)
Instructions (11pt italic, gray)

Section 1: Basic Information (H2, 14pt bold)
  Field Label: * (11pt bold)
  [SDT control]

  Field Label: (11pt bold)
  [SDT control]

Section 2: ... (H2, 14pt bold)
  ...

[set protection=forms — ALWAYS LAST]
```

### Field Ordering Logic

Place fields in the order users naturally think about them:
1. Personal/basic identification (name, ID, email)
2. Role/department/classification
3. Dates (start, end, submission)
4. Supplemental details (notes, preferences, feedback)
5. Confirmation/agreement (checkboxes, signatures)

### Anti-Patterns to Avoid

| Anti-Pattern | Why It Fails | Correct Approach |
|---|---|---|
| `姓名：_______________` underlines | No structured field, invisible to agents | `add --type sdt --prop sdtType=text` |
| `姓名：          ` trailing spaces | Same as underlines | Same correction |
| SDT without alias or tag | Agent cannot identify the field | Always include `--prop alias="X" --prop tag="x"` |
| `--prop sdtType=checkbox` | Silent text degradation — no checkbox appears | `--type formfield --prop formfieldtype=checkbox` |
| protection=forms before all adds | Subsequent add commands fail | All adds first, protection last |
| `--prop items="A;B;C"` semicolons | Items not parsed correctly | `--prop items="A,B,C"` commas only |
| `--prop lock=sdtLocked` camelCase | CLI error: Invalid lock value | `--prop lock=sdtlocked` all lowercase |

---
# officecli: v1.0.41

## Scenario Examples

### Survey Form (All 5 SDT Types)

Demonstrates all five SDT types in a single document:

```bash
officecli create survey.docx
officecli open survey.docx

officecli set survey.docx / \
  --prop title="Customer Satisfaction Survey" \
  --prop docDefaults.font="Calibri" --prop docDefaults.fontSize="12pt"

# Title
officecli add survey.docx /body --type paragraph \
  --prop text="Customer Satisfaction Survey" \
  --prop style=Heading1 --prop size=20 --prop bold=true \
  --prop spaceBefore=0pt --prop spaceAfter=12pt

officecli add survey.docx /body --type paragraph \
  --prop text="Thank you for taking the time to share your feedback. This survey takes approximately 3 minutes." \
  --prop size=11 --prop italic=true --prop color=666666 --prop spaceAfter=18pt

# Section 1: Ratings
officecli add survey.docx /body --type paragraph \
  --prop text="Section 1: Ratings" \
  --prop style=Heading2 --prop size=14 --prop bold=true \
  --prop spaceBefore=18pt --prop spaceAfter=8pt

# NPS (11-item dropdown: 0-10)
officecli add survey.docx /body --type paragraph \
  --prop text="How likely are you to recommend us? (0=Not at all, 10=Extremely likely)" \
  --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add survey.docx /body --type sdt \
  --prop sdtType=dropdown \
  --prop alias="NPS Score" --prop tag="nps_score" \
  --prop items="0,1,2,3,4,5,6,7,8,9,10" \
  --prop lock=sdtlocked

# Usage frequency (dropdown)
officecli add survey.docx /body --type paragraph \
  --prop text="How often do you use our product?" \
  --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add survey.docx /body --type sdt \
  --prop sdtType=dropdown \
  --prop alias="Usage Frequency" --prop tag="usage_freq" \
  --prop items="Daily,Weekly,Monthly,Rarely,First time" \
  --prop lock=sdtlocked

# Section 2: Preferences
officecli add survey.docx /body --type paragraph \
  --prop text="Section 2: Preferences" \
  --prop style=Heading2 --prop size=14 --prop bold=true \
  --prop spaceBefore=18pt --prop spaceAfter=8pt

# Feature preference (combobox — allows custom input beyond preset list)
officecli add survey.docx /body --type paragraph \
  --prop text="Which feature matters most to you? (Select or type your own)" \
  --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add survey.docx /body --type sdt \
  --prop sdtType=combobox \
  --prop alias="Top Feature" --prop tag="top_feature" \
  --prop items="Data Export,API Integration,Reporting,Team Collaboration,Other" \
  --prop lock=sdtlocked

# Contact email (text)
officecli add survey.docx /body --type paragraph \
  --prop text="Your email (optional, for follow-up):" \
  --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add survey.docx /body --type sdt \
  --prop sdtType=text \
  --prop alias="Contact Email" --prop tag="contact_email" \
  --prop text="email@example.com" --prop lock=sdtlocked

# Completion date (date)
officecli add survey.docx /body --type paragraph \
  --prop text="Survey Completion Date:" \
  --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add survey.docx /body --type sdt \
  --prop sdtType=date \
  --prop alias="Completion Date" --prop tag="completion_date" \
  --prop format="MM/dd/yyyy" --prop lock=sdtlocked

# Section 3: Open Feedback
officecli add survey.docx /body --type paragraph \
  --prop text="Section 3: Open Feedback" \
  --prop style=Heading2 --prop size=14 --prop bold=true \
  --prop spaceBefore=18pt --prop spaceAfter=8pt

# Open feedback (richtext — allows long multi-line input)
officecli add survey.docx /body --type paragraph \
  --prop text="What improvements would you suggest?" \
  --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add survey.docx /body --type sdt \
  --prop sdtType=richtext \
  --prop alias="Improvement Suggestions" --prop tag="feedback" \
  --prop text="Please share your thoughts here"

# LAST: Enable protection
officecli set survey.docx / --prop protection=forms

officecli close survey.docx
officecli validate survey.docx
```

### SOW / Contract Form (Fixed Terms + Fillable Fields)

A contract where boilerplate text is locked and only SDT fields are editable:

```bash
officecli create sow.docx
officecli open sow.docx

officecli set sow.docx / \
  --prop title="Statement of Work" \
  --prop docDefaults.font="Calibri" --prop docDefaults.fontSize="12pt"

# Title
officecli add sow.docx /body --type paragraph \
  --prop text="Statement of Work" \
  --prop style=Heading1 --prop size=20 --prop bold=true \
  --prop spaceBefore=0pt --prop spaceAfter=12pt

# Fixed boilerplate (locked after protection=forms)
officecli add sow.docx /body --type paragraph \
  --prop text="This Statement of Work ('SOW') is entered into between the parties identified below and governs the delivery of professional services as described herein." \
  --prop size=11 --prop spaceBefore=12pt --prop spaceAfter=12pt

# Project details section
officecli add sow.docx /body --type paragraph \
  --prop text="1. Project Details" \
  --prop style=Heading2 --prop size=14 --prop bold=true \
  --prop spaceBefore=18pt --prop spaceAfter=8pt

officecli add sow.docx /body --type paragraph \
  --prop text="Project Name:" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add sow.docx /body --type sdt \
  --prop sdtType=text --prop alias="Project Name" --prop tag="project_name" \
  --prop text="Enter project name" --prop lock=sdtlocked

officecli add sow.docx /body --type paragraph \
  --prop text="Client Name:" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add sow.docx /body --type sdt \
  --prop sdtType=text --prop alias="Client Name" --prop tag="client_name" \
  --prop text="Enter client name" --prop lock=sdtlocked

officecli add sow.docx /body --type paragraph \
  --prop text="Contract Start Date:" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add sow.docx /body --type sdt \
  --prop sdtType=date --prop alias="Contract Start Date" --prop tag="contract_start" \
  --prop format="MM/dd/yyyy" --prop lock=sdtlocked

officecli add sow.docx /body --type paragraph \
  --prop text="Contract End Date:" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add sow.docx /body --type sdt \
  --prop sdtType=date --prop alias="Contract End Date" --prop tag="contract_end" \
  --prop format="MM/dd/yyyy" --prop lock=sdtlocked

# Payment section
officecli add sow.docx /body --type paragraph \
  --prop text="2. Payment Terms" \
  --prop style=Heading2 --prop size=14 --prop bold=true \
  --prop spaceBefore=18pt --prop spaceAfter=8pt

# Fixed payment clause (locked after protection)
officecli add sow.docx /body --type paragraph \
  --prop text="Payment shall be made according to the schedule selected below. All invoices are due Net 30 from date of issue." \
  --prop size=11 --prop spaceBefore=12pt --prop spaceAfter=8pt

officecli add sow.docx /body --type paragraph \
  --prop text="Payment Schedule:" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add sow.docx /body --type sdt \
  --prop sdtType=dropdown --prop alias="Payment Schedule" --prop tag="payment_schedule" \
  --prop items="Full Prepayment,50% Upfront / 50% on Completion,Milestone-Based,Net 30 Upon Delivery" \
  --prop lock=sdtlocked

officecli add sow.docx /body --type paragraph \
  --prop text="Total Contract Value (USD):" --prop size=11 --prop bold=true --prop spaceAfter=4pt

officecli add sow.docx /body --type sdt \
  --prop sdtType=text --prop alias="Contract Value" --prop tag="contract_value" \
  --prop text="Enter amount" --prop lock=sdtlocked

# Confidential watermark (optional — remove for non-confidential SOWs)
officecli add sow.docx / --type watermark \
  --prop text="CONFIDENTIAL" --prop color=FF0000 --prop opacity=0.3 --prop rotation=315

# LAST: Enable protection
officecli set sow.docx / --prop protection=forms

officecli close sow.docx
officecli validate sow.docx
```

**Using `--after find:` for precise SDT placement (v1.0.38+):**

```bash
# Insert an SDT immediately after a specific text string
# Useful when the document already has fixed boilerplate text
officecli add sow.docx /body --type sdt \
  --prop sdtType=text --prop alias="Signatory Name" --prop tag="signatory_name" \
  --prop text="Authorized Signatory" --prop lock=sdtlocked \
  --after find:"Client Signature:"
```

---
# officecli: v1.0.41

## QA Checklist (Required)

Run these verification commands after completing every form. Do not skip.

### Hard Rules Verification

```bash
# 1. Verify all SDT controls are present (not underlines)
officecli query form.docx "sdt"
# CHECK: output lists all expected controls with alias, tag, and sdtType
# CHECK: count matches the number of fillable fields you designed

# 2. Verify protection is active
officecli get form.docx /
# CHECK: protection: forms
# CHECK: protectionEnforced: True

# 3. Verify document structure is valid
officecli validate form.docx

# 4. Content review
officecli view form.docx text
officecli view form.docx outline
```

### Pre-Delivery Checklist (Hard Rules)

- [ ] `officecli query form.docx "sdt"` shows all expected controls — no underline substitutes
- [ ] Every SDT in `query sdt` output has both `alias` and `tag` fields
- [ ] `officecli get form.docx /` shows `protection: forms` and `protectionEnforced: True`
- [ ] At least 3 different `sdtType` values appear in `query sdt` output (`text` + `dropdown` + `date` minimum)
- [ ] All `items=` values use English commas, not semicolons
- [ ] All `lock=` values are lowercase (`sdtlocked`, `contentlocked`, `sdtcontentlocked`)
- [ ] `officecli validate form.docx` passes (or any errors are understood/acceptable) — **NOTE**: if the form uses `protection=forms`, validate will always report a `documentProtection` Schema error (Exit 1). This is a known CLI bug and does NOT indicate a problem with the form. Confirm protection is active via `officecli get form.docx /` (`protectionEnforced: True`) instead.
- [ ] No `sdtType: checkbox` appears in `query sdt` — checkbox SDT degrades to text; use formfield instead

### Quality Checks

```bash
# Inspect individual controls (repeat for each key field)
officecli get form.docx '/body/sdt[1]'
# CHECK: editable: True (for fillable fields under forms protection)

# If form has formfields (checkboxes)
officecli get form.docx '/formfield[1]'
# CHECK: formfieldType: checkbox, editable: True
```

---
# officecli: v1.0.41

## Common Pitfalls

| Pitfall | Wrong | Correct | Severity |
|---------|-------|---------|---------|
| checkbox SDT silently degrades | `--prop sdtType=checkbox` | `--type formfield --prop formfieldtype=checkbox` | P0 — Hard Rule fails |
| items semicolons | `--prop items="A;B;C"` | `--prop items="A,B,C"` | P0 — Hard Rule fails |
| lock camelCase | `--prop lock=sdtLocked` | `--prop lock=sdtlocked` | P0 — CLI error |
| protection before all adds | `set protection=forms` then `add sdt` | All `add` first, `set protection=forms` last | P0 — add fails with Exit 1 |
| missing alias or tag | `--prop sdtType=text --prop text="..."` | Add `--prop alias="X" --prop tag="x"` to every SDT | P0 — Hard Rule fails |
| underlines instead of SDT | `姓名：_______________` | `add --type sdt --prop sdtType=text` | P0 — Hard Rule fails |
| unquoted path with [N] in zsh | `get /body/sdt[1]` | `get '/body/sdt[1]'` | P1 — zsh glob error |
| placeholder property | `--prop placeholder="Enter..."` | `--prop text="Enter..."` (placeholder is silently ignored) | P1 — prompt text is lost |
| set items after creation | `set '/body/sdt[N]' --prop items=...` | `remove` the SDT then `add` with new items | P1 — Exit 2: UNSUPPORTED |
| watermark size via set | `set /watermark --prop size=...` | `remove` then `add` with size parameter | P2 — Exit 2: UNSUPPORTED |
| `--prop name=` on SDT | `--prop name="fullname"` | `--prop alias="Full Name" --prop tag="fullname"` | P1 — `name` is alias alias, tag is separate |
| $-signs in --prop text | `--prop text="$50,000"` | Use single quotes: `--prop text='$50,000'` | P1 — shell strips the value |
| items option text with comma | `items="R&D, Labs,Sales"` | Avoid commas in option text; they split items | P1 — option is split wrong |
| date format not validated | `--prop format="invalid"` | Verify format string; errors appear only in Word | P2 — Word shows error on open |

---
# officecli: v1.0.41

## Performance: Resident & Batch Mode

### Resident Mode (Always Recommended)

```bash
officecli open form.docx         # Load once into memory
officecli add form.docx ...      # All commands execute in memory
officecli set form.docx ...
officecli set form.docx / --prop protection=forms
officecli close form.docx        # Write once to disk
```

Always use `open`/`close`. Even a 5-command form build benefits from resident mode — no repeated file I/O.

### Batch Mode (Large Forms)

For forms with many controls, batch mode reduces overhead:

```bash
cat <<'EOF' | officecli batch form.docx
[
  {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Full Name:","bold":true,"size":"11pt","spaceAfter":"4pt"}},
  {"command":"add","parent":"/body","type":"sdt","props":{"sdtType":"text","alias":"Full Name","tag":"full_name","text":"Enter name","lock":"sdtlocked"}},
  {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Department:","bold":true,"size":"11pt","spaceAfter":"4pt"}},
  {"command":"add","parent":"/body","type":"sdt","props":{"sdtType":"dropdown","alias":"Department","tag":"dept","items":"Engineering,Finance,HR,Marketing,Sales","lock":"sdtlocked"}}
]
EOF
```

**Batch notes**:
- Use single-quoted heredoc delimiter `<<'EOF'` to prevent shell variable expansion
- `set / --prop protection=forms` must still be run **after** the batch completes
- Batch supports: `add`, `set`, `get`, `query`, `remove`, `validate`
- If batch fails intermittently, retry or split into smaller chunks (10-15 operations)

---
# officecli: v1.0.41

## Known Issues

| Issue | Workaround |
|-------|-----------|
| `checkbox` SDT silently degrades | `--prop sdtType=checkbox` creates a text SDT. Use `--type formfield --prop formfieldtype=checkbox` instead |
| `placeholder` property ignored | `--prop placeholder=...` is silently ignored. Use `--prop text="..."` to set initial/hint text |
| `add` blocked under protection | `add` fails with Exit 1 under any protection mode. Add all controls before `set protection=forms`, or use `--force` to bypass |
| SDT `items`/`format`/`sdtType` read-only after creation | These properties cannot be changed via `set`. Remove and re-add the SDT to change them |
| Watermark `size`/`width`/`height` not settable | Only configurable at `add` time. To resize: `remove /watermark` then `add` with new dimensions |
| No visual preview for docx | Unlike pptx, there is no `view svg` or `view html`. Use `view text`, `view outline`, `view annotated`. Open in Word for visual check |
| lock error message uses camelCase | Error message shows `contentLocked`, `sdtLocked` (camelCase) but input must be lowercase. The command `--prop lock=sdtlocked` (lowercase) works correctly |
| **`validate` reports Schema error after `protection=forms`** (CLI BUG) | After running `set / --prop protection=forms`, `officecli validate` will always return Exit 1 with `[Schema] unexpected child element documentProtection` — this is a known CLI bug in the schema validator. The protection feature itself works correctly. To confirm protection is active, run `officecli get form.docx /` and verify `protectionEnforced: True`. You can also inspect the XML directly: `officecli docx get --file form.docx --path /documentProtection`. The validate Schema error can be safely ignored when protection is confirmed active. |

---
# officecli: v1.0.41

## Help System

**When unsure about property names or command syntax, run help instead of guessing:**

```bash
officecli docx add              # All addable element types (including sdt, formfield, watermark)
officecli docx set              # All settable elements and their properties
officecli docx get              # All navigable paths
officecli docx query            # Query selector syntax
officecli docx set sdt          # SDT-specific properties
officecli docx add sdt          # SDT add options in detail
```
