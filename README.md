
# Data Analysis and Visualization Toolkit

This repository contains Python scripts for analyzing and visualizing data related to cell dimensions, widths, lengths, and other parameters. The toolkit processes data from multiple directories, generates statistical summaries, and creates visualizations such as histograms, PDFs, CDFs, and error bar plots. The results are saved in Excel files for further analysis.

## Repository Contents

### `data_analasys.py`
- Processes data from multiple directories.
- Generates histograms, PDFs, and CDFs for width, length, theta values, and cell dimensions.
- Saves processed data and visualizations into an Excel file.
- Handles special cases for specific directories with custom filters and bin sizes.

### `summery_with_error_bars.py`
- Combines data from multiple directories to generate summary statistics.
- Creates error bar plots for visualizing the mean and standard deviation of parameters across directories.
- Saves the summary data and plots into an Excel file.

### `run_analasys_and_summery.py`
- Automates the execution of `data_analasys.py` and `summery_with_error_bars.py`.
- Ensures seamless integration of data processing and summary generation.
- Provides a single entry point for running the entire analysis pipeline.

## Features

### Data Processing
- Filters data based on specified thresholds.
- Computes histograms, PDFs, and CDFs for various datasets.
- Calculates statistical metrics such as mean and standard deviation.

### Visualization
- Generates combination charts (bar charts for histograms and line charts for PDFs and CDFs).
- Creates error bar plots for summary statistics.

### Excel Integration
- Saves processed data, visualizations, and summaries into Excel files.
- Uses the `openpyxl` library for creating and formatting Excel sheets.

### Customizable
- Supports multiple directories with different configurations.
- Allows customization of filters, bin sizes, and output file names.

## Requirements

The scripts require the following Python libraries:
- `numpy`  
- `pandas`  
- `openpyxl`  
- `scipy`  
- `matplotlib`  

## Usage

### Input Data

The scripts process CSV files located in specified directories. Each directory should contain the following files:
- `hist_euclidean_width.csv`: Data for width analysis.
- `hist_euclidean_length.csv`: Data for length analysis.
- `hist_theta_values.csv`: Data for theta analysis.
- `cell_dimensions.csv`: Data for cell dimensions (length and width).

### Running the Scripts

#### `data_analasys.py`
Processes data and generates visualizations for each directory.  
The results are saved in an Excel file: `data_analasys.xlsx`.

#### `summery_with_error_bars.py`
Generates summary statistics and error bar plots.  
The results are saved in an Excel file: `summary_with_error_bars.xlsx`.

#### `run_analasys_and_summery.py`
Runs both `data_analasys.py` and `summery_with_error_bars.py` in sequence.

## Customization

### File names

you need to add the directory names you wish to analise in to the list in `data_analasys.py`
- `directory_names = ["Epsilon_3.116_Q_25_G_1.3",
                   "Epsilon_3.116_Q_25_G_1.4",...`
### Filters and Bins

You can customize the filters and bin sizes for each directory by modifying the `process_data_to_excel` function in `data_analasys.py`:

- `filter_value_w`: Filter threshold for width data.
- `filter_value_l`: Filter threshold for length data.
- `bins_w`, `bins_l`, `bins_t`, `bins_cell_dimensions`: Number of bins for histograms.

## Output

### `data_analasys.py`

#### Processed Data
- Original data, filtered data, bins, histograms, PDFs, CDFs, and statistical summaries.

#### Charts
- Combination charts for histograms, PDFs, and CDFs.

### `summery_with_error_bars.py`

#### Summary Data
- Mean and standard deviation for each parameter across directories.

#### Error Bar Plots
- Visualizations of the mean and standard deviation.
