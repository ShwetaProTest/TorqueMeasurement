# Torque Comparison Tool
This tool allows users to compare torque measurements from various files and view the data in a comprehensive plot.

# Table of Contents
* Features
* Installation
* Usage
* Contributing
* License

# Features
* Product Selection: Users can select specific products to compare.
* File Selection: Users can select specific data files to compare.
* Directional Filter: Filter files based on measurement direction (clockwise, anti-clockwise, sweep, or behavior).
* Dynamic Plotting: View and compare the data from the selected files in a plot.
* Interactive Data View: Click on specific data points in the table to highlight them on the plot.

# Installation
1. Ensure you have Python and pip installed.
2. Clone this repository:
```bash
git clone https://github.com/ShwetaProTest/TorqueFlaskapp.git
cd TorqueFlaskapp
```

3. Install the required packages:
```bash
pip install -r requirements.txt
```

4. Run the application:
```bash
python torquetransformation.py
```

# Usage
1. Navigate to the application in your web browser (default is http://127.0.0.1:5000/ if running locally).
2. Use the dropdowns to select the product, direction, and measurement data files.
3. Click "Add Data" to add the file to the comparison.
4. After selecting all the desired files, click "View Plot" to visualize the comparison.
5. Additional data is displayed below the plot.

# Contributing
Pull requests are welcome! For major changes, please open an issue first to discuss what you'd like to change. Please make sure to update the tests as appropriate.
