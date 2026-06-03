# Grading System Application

## Overview
The Grading System application is an automated and flexible Python-based solution for calculating and assigning student grades. It was developed as a term project for **DS-221: Inferential Statistics and Applied Probability** at the Ghulam Ishaq Khan Institute of Engineering Sciences and Technology (GIKI).

## Features
- **Data Handling:** Upload and validate student data seamlessly via Excel files (requires `Quizzes`, `Assignments`, `Midterm`, `Finals`, and `Project` columns).
- **Customizable Weightages:** Dynamically set percentage weightages and total marks for each assessment component.
- **Absolute Grading:** Assign grades based on predefined HEC guidelines or customize your own post-adjustment thresholds.
- **Relative Grading:** Grade on a curve using statistical distributions (mean and standard deviation) with options for custom standard deviation multipliers.
- **Data Visualization:** Generate intuitive bar charts for grade frequency distributions and bell curves to visualize statistical spreads.
- **Export Results:** Save the final processed dataset, including the original inputs, calculated final scores, and assigned grades, to CSV or Excel formats.

## Methodology & Algorithms
1. **Data Ingestion:** Upload an `.xlsx` file containing raw scores. The system automatically handles invalid entries.
2. **Normalization:** The system calculates the final score for each student by normalizing component scores and applying user-defined weightages to ensure they sum perfectly to 100%.
3. **Grading Scheme Application:** - *Absolute:* Compares final scores against fixed boundaries (e.g., 90% = A, 80% = B+).
   - *Relative:* Calculates the overall mean and standard deviation, assigning grades based on statistical ranges (Z-scores).
4. **Adjustment & Refinement:** Educators can adjust the strict grade boundaries or tweak the standard deviation multipliers to refine the grading curve.
5. **Analytics & Export:** The system generates visual analytics and exports a comprehensive output file for record-keeping.

## Technologies Used
- **Python**
- **Pandas** (Data manipulation and export)
- **Matplotlib / SciPy / NumPy** (Statistical calculations, normal distributions, and visualizations)
- **OpenPyXL** (Excel file handling)
- **Tkinter** (Graphical User Interface)

## Contributors
- **Mohammad Musa Ali** (2023330)
- **Sarosh Ishaq** (2023640)
- **Muhammad Umer** (2023538)

**Instructor:** Sir Sajid Ali
