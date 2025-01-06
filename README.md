# Well Penetration Chart Generator

This Python script generates a **Well Penetration Chart** using pie charts to visualize the presence and state of play elements (Reservoir, Seal, Charge, Trap) for wells and plays. It also exports the pie charts as PNG files and updates the input data with the file paths of the exported charts.

---

## 📋 Table of Contents
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Input Data Format](#input-data-format)
- [Output](#output)
- [Example](#example)
- [License](#license)

---

## ✨ Features
- **Dynamic Pie Charts**: Visualizes the state of play elements (Present, Ambiguous, Not Present, Unevaluated) using pie charts.
- **Customizable Input**: Supports customizable column names for wells, plays, and play elements.
- **Export Options**:
  - Exports pie charts as **PNG files** (128x128 pixels).
  - Updates the input file with the paths of the exported pie charts.
  - Supports output in **Excel** or **CSV** format.
- **Flexible Configuration**: Allows customization of chart size, colors, and grid properties.

---

## 📦 Requirements
- Python 3.x
- Libraries:
  - `pandas`
  - `matplotlib`
  - `numpy`

Install the required libraries using:
```bash
pip install pandas matplotlib numpy
```

---

## 🛠 Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/well-penetration-chart.git
   cd well-penetration-chart
   ```
2. Install the required libraries (see [Requirements](#requirements)).

---

## 🚀 Usage
Run the script with the following command:
```bash
python well_penetration_chart.py
```

### Example Command
```python
main(
    input_file=r'C:\plays.xlsm',
    well_col='Well Name',
    reservoir_col='Reservoir',
    seal_col='Seal',
    charge_col='Charge',
    trap_col='Trap',
    play_col='Play',
    age_col='Age',
    cell_width=2.0,
    cell_height=1.0,
    header_cell_width=1.5,
    header_cell_height=1.5,
    vertical_header_width=0.8,
    grid_line_thickness=0.5,
    grid_line_color='#DCDCDC',  # light grey
    export_dir=r'C:\output_directory',  # Directory to save exported pie charts
    export_pie_charts=True,  # Export pie charts as PNG files
    export_updated_file=True,  # Export updated file with PieChartPath column
    output_file_format='csv'  # File format for updated file ('excel' or 'csv')
)
```

---

## 📂 Input Data Format
The input file should be an **Excel file** with the following columns (default column names can be customized):
- **Well Name**: Name of the well.
- **Reservoir**: State of the reservoir (e.g., Present, Ambiguous, Not Present, Unevaluated).
- **Seal**: State of the seal.
- **Charge**: State of the charge.
- **Trap**: State of the trap.
- **Play**: Name of the play.
- **Age**: Age of the well (used for sorting).
- **Color** (optional): Color for each play (used for visualization).

---

## 📤 Output
1. **Pie Charts**:
   - Exported as PNG files (128x128 pixels) to the specified `export_dir`.
   - File names follow the format: `{Play}_{Well}.png`.

2. **Updated Data File**:
   - A new column `PieChartPath` is added to the input data, containing the full paths of the exported pie charts.
   - The updated data is saved in the specified format (`Excel` or `CSV`).

---

## 🖼 Example
### Input Data
| Well Name | Reservoir   | Seal        | Charge      | Trap        | Play     | Age | Color  |
|-----------|-------------|-------------|-------------|-------------|----------|-----|--------|
| Well-1    | Present     | Ambiguous   | Not Present | Unevaluated | Play-A   | 10  | #FAFFB0|
| Well-2    | Ambiguous   | Present     | Present     | Not Present | Play-B   | 5   | #EBAC99|

### Output
- **Pie Charts**:
  - `Play-A_Well-1.png`
  - `Play-B_Well-2.png`
- **Updated Data File**:
  - A new column `PieChartPath` is added with the paths to the exported PNG files.

---

## 📄 License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments
- Built with ❤️ using `pandas`, `matplotlib`, and `numpy`.
- Inspired by the need for clear visualization of well penetration data.

---

For questions or feedback, please open an issue or contact the author.
