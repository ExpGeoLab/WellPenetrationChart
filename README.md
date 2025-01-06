# Well Penetration Chart Generator

**Drilled Hole Analysis (DHA):** A systematic review of petroleum system elements (Reservoir, Seal, Trap, and Charge) within drilled wells, DHA enables a structured and detailed understanding of subsurface features and potential. The play elements can be described as either: ‘present’, ‘uncertain’, ‘void or absent’ and ‘unassessed’.

**Purpose:** To identify the presence or absence of key elements and calibrate play maps. Can be used to help constrain the play maps.

**Display:** Individual well evaluations are visualized as pie charts or wagon wheels, The overall results can be used to display as well penetration charts, offering interpretable visuals to identify trends and outliers across regions.

This Python script generates a **Well Penetration Chart** using pie charts to visualize the presence and state of play elements (Reservoir, Seal, Charge, Trap) for wells and plays. It also exports the pie charts as PNG files and updates the input data with the file paths of the exported charts.
![WellPenetrationChart](https://github.com/user-attachments/assets/d7996288-a296-4384-904a-e52ba76bc51a)

---

## 📋 Table of Contents
- [Features](#-features)
- [Q-GIS Integration](-integrating-exported-pie-charts-into-qgis)
- [Requirements](#-requirements)
- [Installation](#-installation)
- [Usage](#-usage)
- [Input Data Format](#-input-data-format)
- [Output](#-output)
- [Example](#-example)
- [License](#-license)

---

## ✨ Features
- **Dynamic Pie Charts**: Visualizes the state of play elements (Present, Ambiguous, Not Present, Unevaluated) using pie charts.
- **Customizable Input**: Supports customizable column names for wells, plays, and play elements.
- **Export Options**:
  - Exports pie charts as **PNG files** (128x128 pixels). *Enabled by default (export_pie_charts=True)*.
  - Updates the input file with the paths of the exported pie charts.
  - Supports output in **Excel** or **CSV** format. *Enabled by default (export_updated_file=True)*.
- **Flexible Configuration**: Allows customization of chart size, colors, and grid properties.
- **QGIS Integration**: The exported pie charts and their file paths can be seamlessly integrated into QGIS. The `PieChartPath` column in the updated CSV or Excel file can be used to replace point markers with the corresponding raster images (pie charts) in QGIS, enabling dynamic visualization of well penetration data on maps.
![Q-GIS](https://github.com/user-attachments/assets/3985e14a-c9c9-4654-af7e-a958fad9108b)

### 🗺️ **Integrating Exported Pie Charts into QGIS**

To replace traditional point markers with the exported pie charts in QGIS, follow these steps:

#### 1. **Load the Updated CSV/Excel File**
   - Open QGIS and go to **Layer > Add Layer > Add Delimited Text Layer** (for CSV) or **Add Vector Layer** (for Excel).
   - Load the updated file containing the `PieChartPath` column.

#### 2. **Style the Layer with Raster Markers**
   - Right-click the layer in the **Layers Panel** and select **Properties**.
   - Go to the **Symbology** tab.
   - Change the symbology type to **Raster Image Marker**.
   - Click the **Data-Driven Override** button (ε) next to the **Image Path** field.
   - Choose **Field Type** and select the `PieChartPath` column.

#### 3. **Adjust Marker Size**
   - Set the **Size** of the raster markers to fit your map (e.g., 10 mm).
   - Optionally, use a **data-defined override** for dynamic sizing based on an attribute (e.g., well size or production).

#### 4. **Visualize and Customize**
   - The pie charts will now replace the point markers on the map.
   - Customize the layer further (e.g., labels, transparency) as needed.

#### 5. **Save and Share**
   - Save your QGIS project to retain the styling.
   - Share the project or export the map for presentations or reports.


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
- **Age**: Age of the well (used for sorting - Mandatory).
- **Color** (optional): Color for each play (used for visualization).
![Excel](https://github.com/user-attachments/assets/59149a1f-64b0-4f86-8217-9aa4160ee968)

---
### **How to Use the Input [Excel File](plays.xlsm)**

To integrate the provided Excel function and VBA macro into your workbook, follow these steps:

---

#### **1. Add the International Reference Chronostratigraphic Chart**
- Create a sheet named `References` in your workbook.
- Populate the sheet with the following columns:
  - **Column C**: Stage Names (e.g., `References!$C$2:$C$104`).
  - **Column D**: Age Values (e.g., `References!$D$2:$D$104`).
  - **Column E**: HEX Color Codes (optional, if you want to store colors directly).

---

#### **2. Add the Excel Function**
- In your main worksheet, use the following formula to get the **stage name** or **HEX color code** based on the entered age:
  ```excel
  =INDEX(References!$C$2:$C$104, MATCH(MIN(ABS(References!$D$2:$D$104 - G2)), ABS(References!$D$2:$D$104 - G2), 0))
  ```
  - Replace `G2` with the cell containing the entered age.
  - To get the **HEX color code** instead of the stage name, replace `References!$C$2:$C$104` with the column containing the HEX codes (e.g., `References!$E$2:$E$104`).

---

#### **3. Add the VBA Macro**
- Open the VBA editor (`Alt + F11`).
- Locate the worksheet where you want the macro to run (e.g., `Sheet1`).
- Copy and paste the provided VBA code into the worksheet's code window.
- Save the workbook as a **Macro-Enabled Workbook** (`.xlsm`).

---

#### **4. Use the Combined Functionality**
1. **Enter an Age**:
   - In your main worksheet, enter an age value in a cell (e.g., `G2`).

2. **Get the Stage Name or HEX Code**:
   - Use the Excel function to retrieve the corresponding stage name or HEX color code.

3. **Apply the HEX Code**:
   - If you retrieve the HEX code, enter it into another cell (e.g., `#FF5733`).
   - The VBA macro will automatically apply the corresponding color to the cell's background.

---

### **Example Workflow**
1. **Input Age**: Enter `65` in cell `G2`.
2. **Get Stage Name**:
   - Use the formula to get the stage name (e.g., `Paleocene`).
3. **Get HEX Code**:
   - Use the formula to get the HEX code (e.g., `#A3FF99`).
4. **Apply Color**:
   - Enter the HEX code into a cell, and the VBA macro will apply the color.

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
This project is licensed under the MIT License. See the [LICENSE](LICENSE.txt) file for details.

---

## 🙏 Acknowledgments
- Built with ❤️ using `pandas`, `matplotlib`, and `numpy`.
- Inspired by the need for clear visualization of well penetration data.

---

For questions or feedback, please open an issue or contact the author.
