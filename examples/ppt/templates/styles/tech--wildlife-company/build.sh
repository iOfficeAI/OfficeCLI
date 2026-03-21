#!/bin/bash
set -e

FILENAME="野生动物科技公司.pptx"

echo "Creating PPT: $FILENAME"

# Create PPT file
officecli create "$FILENAME"

# Add slide 1 with background
officecli add "$FILENAME" '/' --type slide --prop layout=blank --prop background=FFF8F0

# Add all actors to slide 1 using batch
cat << 'EOF' | officecli batch "$FILENAME"
[
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "dot-accent1",
      "preset": "ellipse",
      "fill": "FF8C42",
      "opacity": "0.12",
      "x": "28cm",
      "y": "2cm",
      "width": "7cm",
      "height": "7cm"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "dot-accent2",
      "preset": "ellipse",
      "fill": "FFD166",
      "opacity": "0.15",
      "x": "2cm",
      "y": "13cm",
      "width": "5cm",
      "height": "5cm"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "rect-bg",
      "preset": "roundRect",
      "fill": "6AB547",
      "opacity": "0.1",
      "x": "0cm",
      "y": "7cm",
      "width": "8cm",
      "height": "10cm"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "triangle-corner",
      "preset": "triangle",
      "fill": "FF8C42",
      "opacity": "0.12",
      "x": "1cm",
      "y": "1cm",
      "width": "3cm",
      "height": "3cm",
      "rotation": "30"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "leaf-shape",
      "preset": "ellipse",
      "fill": "6AB547",
      "opacity": "0.12",
      "x": "25cm",
      "y": "12cm",
      "width": "4cm",
      "height": "6cm",
      "rotation": "45"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "circle-small",
      "preset": "ellipse",
      "fill": "FFD166",
      "opacity": "0.15",
      "x": "30cm",
      "y": "15cm",
      "width": "2.5cm",
      "height": "2.5cm"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "roundrect-float",
      "preset": "roundRect",
      "fill": "FF8C42",
      "opacity": "0.08",
      "x": "15cm",
      "y": "1cm",
      "width": "5cm",
      "height": "3cm",
      "rotation": "15"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "ellipse-glow",
      "preset": "ellipse",
      "fill": "6AB547",
      "opacity": "0.1",
      "x": "24cm",
      "y": "8cm",
      "width": "6cm",
      "height": "4cm"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "hero-title",
      "text": "野生动物科技公司",
      "font": "PingFang SC",
      "size": "68",
      "bold": "true",
      "color": "4A4A4A",
      "align": "center",
      "valign": "middle",
      "x": "6cm",
      "y": "6cm",
      "width": "22cm",
      "height": "3cm"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "hero-subtitle",
      "text": "WildTech Inc. — 让天性驱动创新",
      "font": "PingFang SC",
      "size": "32",
      "color": "FF8C42",
      "align": "center",
      "valign": "middle",
      "x": "8cm",
      "y": "9.5cm",
      "width": "18cm",
      "height": "1.5cm"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "hero-tagline",
      "text": "当动物王国遇见硅谷",
      "font": "PingFang SC",
      "size": "20",
      "color": "6AB547",
      "align": "center",
      "valign": "middle",
      "x": "11cm",
      "y": "11.5cm",
      "width": "12cm",
      "height": "1cm"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "section-title",
      "text": "",
      "font": "PingFang SC",
      "size": "40",
      "bold": "true",
      "color": "4A4A4A",
      "x": "36cm",
      "y": "0cm",
      "width": "20cm",
      "height": "2cm"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "statement-text",
      "text": "",
      "font": "PingFang SC",
      "size": "52",
      "bold": "true",
      "color": "4A4A4A",
      "align": "center",
      "valign": "middle",
      "x": "36cm",
      "y": "5cm",
      "width": "26cm",
      "height": "3cm"
    }
  },
  {
    "command": "add",
    "parent": "/slide[1]",
    "type": "shape",
    "props": {
      "name": "statement-sub",
      "text": "",
      "font": "PingFang SC",
      "size": "24",
      "color": "6AB547",
      "align": "center",
      "valign": "middle",
      "x": "36cm",
      "y": "10cm",
      "width": "28cm",
      "height": "4cm"
    }
  }
]
EOF

echo "Slide 1 complete (hero)"

# Clone and adjust for slide 2
officecli add "$FILENAME" '/' --from '/slide[1]'
cat slide2.json | officecli batch "$FILENAME"
echo "Slide 2 complete (statement)"

# Clone and adjust for slide 3
officecli add "$FILENAME" '/' --from '/slide[1]'
cat slide3.json | officecli batch "$FILENAME"
echo "Slide 3 complete (pillars)"

# Clone and adjust for slide 4
officecli add "$FILENAME" '/' --from '/slide[1]'
cat slide4.json | officecli batch "$FILENAME"
echo "Slide 4 complete (evidence)"

# Clone and adjust for slide 5
officecli add "$FILENAME" '/' --from '/slide[1]'
cat slide5.json | officecli batch "$FILENAME"
echo "Slide 5 complete (showcase)"

# Clone and adjust for slide 6
officecli add "$FILENAME" '/' --from '/slide[1]'
cat slide6.json | officecli batch "$FILENAME"
echo "Slide 6 complete (quote)"

# Clone and adjust for slide 7
officecli add "$FILENAME" '/' --from '/slide[1]'
cat slide7.json | officecli batch "$FILENAME"
echo "Slide 7 complete (cta)"

# Validate
echo "Validating PPT..."
officecli validate "$FILENAME"

echo "✅ PPT generation complete: $FILENAME"
echo "View outline:"
officecli view "$FILENAME" outline
