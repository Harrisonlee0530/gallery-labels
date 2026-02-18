# Gallery Labels

Gallery Labels is a Shiny for Python application for creating, managing, and exporting exhibition labels formatted for print (fractional A4 size cards).

The application allows manual entry or bulk upload of artwork metadata and supports export to print-ready PDF (double-spaced, embedded font)

---

## Features

- Create labels via sidebar form inputs
- Bulk upload labels from CSV
- Delete labels using a title-based dropdown selector
- Session-isolated user data (no shared state between users)
- Export labels as:
  - PDF (print-ready, double-spaced typography)

Each label includes:

- Header image (horizontal)
- Title
- Height × Width (cm)
- Medium
- Date
- Optional comments

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/gallery-labels.git
cd gallery-labels
````

### 2. Create the environment

```bash
conda env create -f environment.yml
conda activate gallery-labels
```

### 3. Run the application

```bash
shiny run --reload app.py
```

Open the displayed local URL in your browser.

---

## CSV Upload Format

The CSV file must contain **exactly** the following columns:

```
title,height,width,medium,date,comments
```

### Column Definitions

| Column   | Required | Format     | Description             |
| -------- | -------- | ---------- | ----------------------- |
| title    | Yes      | Text       | Artwork title           |
| height   | Yes      | Numeric    | Height in centimeters   |
| width    | Yes      | Numeric    | Width in centimeters    |
| medium   | Yes      | Text       | Materials or technique  |
| date     | Yes      | YYYY/MM/DD | Date in required format |
| comments | No       | Text       | Optional; may be empty  |

---

## Date Format Requirement

Dates must follow this exact format:

```
YYYY/MM/DD
```

Example:

```
2024/15/03
```

This corresponds to 15 March 2024.

Incorrect formats (e.g., `YYYY/DD/MM`) may fail in certain cases.

---

## Example CSV

```
title,height,width,medium,date,comments
Untitled,30,20,Oil on canvas,2024/15/03,Private collection
Study II,50,40,Ink on paper,2023/01/12,
```

If the `comments` field is empty, it will not appear on the label (labeled in white).

---

## Deployment

Compatible with:

* shinyapps.io
* Posit Connect
* Any Python hosting platform supporting Shiny

Each user session is isolated. Data is not persisted unless explicitly written to external storage.

---

## License

This project is licensed under the MIT License.
See the `LICENSE` file for full details.
