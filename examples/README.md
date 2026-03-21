# OfficeCLI Examples

Comprehensive examples demonstrating OfficeCLI capabilities for Word, Excel, and PowerPoint automation.

## 📂 Directory Structure

```
examples/
├── README.md                          # This file
├── word/                              # 📄 Word examples (3 scripts)
│   ├── README.md
│   ├── gen-formulas.sh
│   ├── gen-complex-tables.sh
│   ├── gen-complex-textbox.sh
│   └── outputs/
├── excel/                             # 📊 Excel examples (2 scripts)
│   ├── README.md
│   ├── gen-beautiful-charts.sh
│   ├── gen-charts-demo.sh
│   └── outputs/
└── ppt/                               # 🎨 PowerPoint (3 scripts + 14 style templates)
    ├── README.md
    ├── gen-beautiful-pptx.sh
    ├── gen-animations-pptx.sh
    ├── gen-video-pptx.py
    ├── outputs/
    └── templates/                     # 14 Professional Style Templates ⭐
        ├── README.md
        └── styles/                    # (all with pre-generated PPTs)
            ├── dark--*/               (14 dark styles)
            ├── light--*/              (8 light styles)
            ├── warm--*/               (5 warm styles)
            ├── vivid--*/              (2 vivid styles)
            ├── bw--*/                 (3 black & white)
            └── mixed--*/              (1 mixed style)
```

---

## 🚀 Quick Start

### By Document Type

**Word (.docx):**
```bash
cd word
bash gen-formulas.sh            # LaTeX math formulas
bash gen-complex-tables.sh      # Styled tables
bash gen-complex-textbox.sh     # Formatted text boxes
```

**Excel (.xlsx):**
```bash
cd excel
bash gen-beautiful-charts.sh    # Professional charts
bash gen-charts-demo.sh         # 14+ chart types
```

**PowerPoint (.pptx):**
```bash
cd ppt
bash gen-beautiful-pptx.sh      # Morph transitions
bash gen-animations-pptx.sh     # Animation effects
python gen-video-pptx.py        # Video embedding
```

### Professional Style Templates

```bash
cd ppt/templates/styles/dark--investor-pitch
# View pre-generated PPT
open template.pptx

# Or regenerate
bash build.sh
```

👉 **[Browse all 35 styles →](ppt/templates/)** (all with pre-generated PPTs)

---

## 📚 Documentation by Type

### 📄 [Word Examples →](word/)
**3 scripts demonstrating:**
- Mathematical formulas (LaTeX)
- Complex table creation
- Text styling and formatting

**Key Techniques:**
- Paragraph and run manipulation
- Table structure and styling
- Font and color formatting
- Document structure navigation

---

### 📊 [Excel Examples →](excel/)
**2 scripts demonstrating:**
- Professional chart creation
- Multiple chart types (14+)
- Data visualization

**Key Techniques:**
- Cell value and formula manipulation
- Chart creation and styling
- Data range selection
- Number formatting

---

### 🎨 [PowerPoint Examples →](ppt/)
**3 scripts + 35 professional style templates:**
- Morph transitions
- Animation effects
- Video embedding
- 35 design styles (15 ready-to-use)

**Key Techniques:**
- Slide creation and layout
- Shape positioning and styling
- Transitions and animations
- Media embedding
- Professional design patterns

**Style Categories:**
- 🌑 **Dark** (14) - Tech, corporate, futuristic
- ☀️ **Light** (8) - Clean, professional, product showcases
- 🧡 **Warm** (5) - Friendly, lifestyle, organic brands
- 🌈 **Vivid** (2) - Energetic, youthful marketing
- ⬛ **Black & White** (3) - Minimalist, sophisticated
- 🎨 **Mixed** (1) - Bold architectural designs

---

## 🎓 Learning Path

### Beginner (Start Here)
1. **Word** - [`gen-formulas.sh`](word/gen-formulas.sh)
2. **Excel** - [`gen-charts-demo.sh`](excel/gen-charts-demo.sh)
3. **PowerPoint** - Simple shape creation

**Learn:** Basic commands, file structure, properties

---

### Intermediate
4. **Word** - [`gen-complex-tables.sh`](word/gen-complex-tables.sh)
5. **Excel** - [`gen-beautiful-charts.sh`](excel/gen-beautiful-charts.sh)
6. **PowerPoint** - [`gen-animations-pptx.sh`](ppt/gen-animations-pptx.sh)

**Learn:** Batch operations, styling, advanced properties

---

### Advanced
7. **Style Templates** - Explore [professional styles](ppt/templates/)
8. **PowerPoint** - [`gen-beautiful-pptx.sh`](ppt/gen-beautiful-pptx.sh)
9. **Python Integration** - [`gen-video-pptx.py`](ppt/gen-video-pptx.py)

**Learn:** Morph transitions, complex layouts, design patterns, automation

---

## 🔧 Common Patterns

### Create and Populate

```bash
#!/bin/bash
set -e

FILE="document.docx"
officecli create "$FILE"
officecli add "$FILE" /body --type paragraph --prop text="Hello World"
officecli validate "$FILE"
```

### Batch Operations

```bash
cat << 'EOF' > commands.json
[
  {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Para 1"}},
  {"command":"set","path":"/body/p[1]","props":{"bold":"true","size":"24"}}
]
EOF
officecli batch document.docx < commands.json
```

### Resident Mode (3+ operations)

```bash
officecli open document.docx
officecli add document.docx /body --type paragraph --prop text="Fast operation"
officecli set document.docx /body/p[1] --prop bold=true
officecli close document.docx
```

### Query and Modify

```bash
# Find all Heading1 paragraphs
officecli query report.docx "paragraph[style=Heading1]" --json

# Change their color
officecli set report.docx /body/p[1] --prop color=FF0000
```

---

## 📊 Quick Reference

### Document Types

| Format | Extension | Create | View | Modify |
|--------|-----------|--------|------|--------|
| Word | .docx | ✓ | ✓ | ✓ |
| Excel | .xlsx | ✓ | ✓ | ✓ |
| PowerPoint | .pptx | ✓ | ✓ | ✓ |

### Common Commands

| Command | Purpose | Example |
|---------|---------|---------|
| `create` | Create blank document | `officecli create file.docx` |
| `view` | View content | `officecli view file.docx text` |
| `get` | Get element | `officecli get file.docx /body/p[1]` |
| `set` | Modify element | `officecli set file.docx /body/p[1] --prop bold=true` |
| `add` | Add element | `officecli add file.docx /body --type paragraph` |
| `remove` | Remove element | `officecli remove file.docx /body/p[5]` |
| `query` | CSS-like query | `officecli query file.docx "paragraph[style=Normal]"` |
| `batch` | Multiple operations | `officecli batch file.docx < commands.json` |
| `validate` | Check schema | `officecli validate file.docx` |

### View Modes

| Mode | Description | Usage |
|------|-------------|-------|
| `text` | Plain text | `officecli view file.docx text` |
| `annotated` | Text with formatting | `officecli view file.docx annotated` |
| `outline` | Structure | `officecli view file.docx outline` |
| `stats` | Statistics | `officecli view file.docx stats` |
| `issues` | Problems | `officecli view file.docx issues` |

---

## 💡 Tips

1. **Explore before modifying:**
   ```bash
   officecli view document.docx outline
   officecli get document.docx /body --depth 2
   ```

2. **Use `--json` for automation:**
   ```bash
   officecli query data.xlsx "cell[formula~=SUM]" --json | jq
   ```

3. **Check help for properties:**
   ```bash
   officecli docx set paragraph
   officecli xlsx set cell
   officecli pptx set shape
   ```

4. **Validate after changes:**
   ```bash
   officecli validate document.docx
   ```

5. **Use resident mode for performance:**
   ```bash
   # For 3+ operations on same file
   officecli open file.pptx
   # ... multiple commands ...
   officecli close file.pptx
   ```

6. **Batch for complex operations:**
   - Single open/save cycle
   - Atomic transactions
   - Better performance

---

## 🤝 Contributing Examples

Want to add an example? Follow this structure:

1. **Create script** with clear comments
2. **Test and verify** output
3. **Add to appropriate directory** (word/excel/ppt)
4. **Update directory README**
5. **Submit PR**

**Example format:**
```bash
#!/bin/bash
# Brief description of what this demonstrates
# Key techniques: list them here

set -e

FILE="output.docx"
officecli create "$FILE"
# ... your commands ...
officecli validate "$FILE"
echo "✅ Created: $FILE"
```

---

## 📖 More Resources

- **[SKILL.md](../SKILL.md)** - Complete command reference for AI agents
- **[README.md](../README.md)** - Project overview and installation
- **[API Documentation](../docs/)** - Detailed API reference

---

## 🆘 Getting Help

**Command help:**
```bash
officecli --help
officecli docx --help
officecli docx set --help
officecli pptx set shape
```

**Three-layer help navigation:**
```bash
officecli pptx set              # All settable elements
officecli pptx set shape        # Shape properties
officecli pptx set shape.fill   # Fill property details
```

---

**Happy automating! 🚀**

For questions or issues, visit [GitHub Issues](https://github.com/iOfficeAI/OfficeCLI/issues).
