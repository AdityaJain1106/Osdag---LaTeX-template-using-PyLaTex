````markdown
# 📘 PyLaTeX Beam Analysis Report Generator

Automatically generates a **professional engineering PDF report** using:

- Python  
- PyLaTeX  
- pandas  
- TikZ / PGFPlots  
- MiKTeX LaTeX compiler  

The script reads loading data from an **Excel file**, computes the **Shear Force Diagram (SFD)** and **Bending Moment Diagram (BMD)**, recreates the input table in LaTeX, embeds the beam image, and generates a complete engineering report PDF.

---

## 📑 Table of Contents
- [About The Project](#about-the-project)
- [Built With](#built-with)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [Usage](#usage)
- [Project Structure](#project-structure)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)
- [Acknowledgments](#acknowledgments)

---

## 📘 About The Project

### **Automated Engineering PDF Report Generation**

This project solves a common real-world problem for structural engineers:  
➡️ **Generating clean, formatted engineering reports automatically.**

Using PyLaTeX, the script:

### ✔ Reads force/load data from an Excel file  
### ✔ Computes reactions, SFD, BMD  
### ✔ Generates TikZ/PGFPlots diagrams (no images!)  
### ✔ Recreates Excel tables in LaTeX Tabular (selectable text)  
### ✔ Embeds a beam image  
### ✔ Produces a full engineering PDF report with:
- Title page  
- TOC  
- Introduction  
- Beam description  
- Input table  
- SFD/BMD plots  
- Summary  

This tool eliminates repeated manual formatting work and produces consistent engineering documentation.

---

## 🛠 Built With

Major dependencies used in this project:

- **Python 3.10+**
- **PyLaTeX**
- **pandas**
- **numpy**
- **MiKTeX (pdflatex)**
- **TikZ/pgfplots** (inside LaTeX)

---

## 🚀 Getting Started

Follow these steps to get the project running on your local system.

---

### ✅ Prerequisites

#### **Install Python packages**
```bash
pip install pylatex pandas numpy openpyxl
````

#### **Install MiKTeX (for LaTeX PDF generation)**

Download here: [https://miktex.org/download](https://miktex.org/download)

During installation:
✔ ENABLE **"Install missing packages on-the-fly"**

Ensure `pdflatex` is available:

```bash
where pdflatex
```

---

### 🧩 Installation

1. **Fork this repository**

2. **Clone your fork**

   ```bash
   git clone https://github.com/<your-username>/<your-repo>.git
   ```

3. **Add your Excel loading data file**
   Example format:

   | Position | Force |
   | -------- | ----- |
   | 2        | 10    |
   | 5        | 20    |
   | 8        | 15    |

4. **Add your beam image file** (PNG/JPG)

5. Set paths inside `generate_report.py`:

   ```python
   EXCEL_PATH = "yourfile.xlsx"
   BEAM_IMAGE_PATH = "beam.png"
   ```

6. Run the script:

   ```bash
   python generate_report.py
   ```

7. Your `Engineering_Report.pdf` will be generated in the same folder.

---

## 📘 Usage

This project demonstrates:

* How to automate engineering documentation
* How to use PyLaTeX to generate:

  * Tables
  * Figures
  * TikZ plots
* How to compute SFD/BMD for simply supported beams
* How to embed external images and computed diagrams into a PDF

Each time you change the Excel file,
➡️ Run the script again to generate a new updated report.

---

## 📁 Project Structure

```
📦 project-folder
 ┣ 📄 generate_report.py
 ┣ 📄 beam.png
 ┣ 📄 Book1.xlsx
 ┣ 📄 Engineering_Report.pdf
 ┗ 📄 README.md
```

---

## 🗺 Roadmap

* [ ] Add command-line arguments (CLI support)
* [ ] Add support for UDL, moments, distributed loads
* [ ] Add stress diagrams
* [ ] Add multi-page appendix support
* [ ] Add more LaTeX styling templates

---

## 🤝 Contributing

Contributions are welcome and appreciated!

1. Fork the Project
2. Create a Feature Branch

   ```bash
   git checkout -b feature/NewFeature
   ```
3. Commit Your Changes

   ```bash
   git commit -m "Add new feature"
   ```
4. Push the Branch

   ```bash
   git push origin feature/NewFeature
   ```
5. Open a Pull Request

---

## 📜 License

Distributed under the **UNLICENSE** License.
See `LICENSE` file for more information.

---

## 📬 Contact

**Aditya Jain**
GitHub: [https://github.com/AdityaJain1106](https://github.com/AdityaJain1106)
Email: [adijain1106@gmail.com](mailto:adijain1106@gmail.com)

Project Link:
👉 [https://github.com/AdityaJain1106/Osdag---LaTeX-template-using-PyLaTex](https://github.com/AdityaJain1106/Osdag---LaTeX-template-using-PyLaTex)

---

## ⭐ Acknowledgments

Some useful resources used during development:

* PyLaTeX Documentation
* MiKTeX
* TikZ & PGFPlots
* pandas
* numpy
* CTAN Packages
* GitHub Emoji Cheat Sheet
* Flexbox & Grid Cheatsheets
* Open Source Community 💙

