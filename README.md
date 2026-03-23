# PPT Maker

Generate `.pptx` presentations from a `content.json` file.

## Usage

```bash
python make_ppt.py [-h] [-i INPUT] [-o OUTPUT]
```
json format as follow

## Available Slide Types

### Title slides

| Type             | Description                                      |
| ---------------- | ------------------------------------------------ |
| `title`          | Accent bar left, centered text                   |
| `title_centered` | Decorative circle behind centered text, white bg |
| `title_split`    | Left half accent-colored, right half white       |

### Section slides

| Type             | Description                                   |
| ---------------- | --------------------------------------------- |
| `section`        | Top/bottom accent lines, centered title       |
| `section_banner` | Full-width accent banner in the middle        |
| `section_corner` | Corner accent decorations, left-aligned title |

### Content slides

| Type               | Description                                          |
| ------------------ | ---------------------------------------------------- |
| `content`          | Sky blue header bar, bullet list                     |
| `content_cards`    | Level-0 bullets become card boxes in a 2-column grid |
| `content_numbered` | Level-0 bullets get numbered circles                 |
| `content_sidebar`  | Left accent sidebar with title, bullets on right     |

### Two-column slides

| Type            | Description                               |
| --------------- | ----------------------------------------- |
| `two_col`       | Standard two columns with divider         |
| `two_col_boxed` | Two columns with colored card backgrounds |

### Special slides

| Type         | Fields                                            | Description                           |
| ------------ | ------------------------------------------------- | ------------------------------------- |
| `quote`      | `text`, `attribution`                             | Large centered quote with attribution |
| `three_col`  | `col1`, `col2`, `col3`, `head1`, `head2`, `head3` | Three equal columns                   |
| `big_number` | `number`, `label`, `description`                  | Hero stat slide with large number     |

## JSON Format

```json
{
  "output": "path/to/output.pptx",
  "slides": [
    {
      "type": "title",
      "title": "Presentation Title",
      "subtitle": "Optional subtitle"
    },
    {
      "type": "content",
      "title": "Slide Title",
      "bullets": [
        [0, "Main bullet"],
        [1, "Sub-bullet"]
      ]
    },
    {
      "type": "two_col",
      "title": "Comparison",
      "left_head": "Left Header",
      "right_head": "Right Header",
      "left": ["Item 1", "Item 2"],
      "right": ["Item A", "Item B"]
    },
    {
      "type": "quote",
      "text": "Some insightful quote here.",
      "attribution": "Author Name"
    },
    {
      "type": "big_number",
      "number": "42",
      "label": "Papers Surveyed",
      "description": "Optional extra context"
    }
  ]
}
```

## Bullet Levels

In `content`-type slides, bullets are `[level, text]` pairs:
- Level `0` — main bullet (▸, bold, larger font)
- Level `1` — sub-bullet (•, normal, smaller font)
