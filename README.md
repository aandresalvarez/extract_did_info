# De-identification Information Extractor and Analyzer

## Table of Contents
1. [Introduction](#introduction)
2. [Features](#features)
3. [Installation](#installation)
4. [Usage](#usage)
5. [Project Structure](#project-structure)
6. [Configuration](#configuration)
7. [Advanced Usage](#advanced-usage)
8. [Troubleshooting](#troubleshooting)
9. [Contributing](#contributing)
10. [License](#license)

## Introduction

The De-identification Information Extractor and Analyzer is a powerful tool designed to streamline the process of managing and analyzing de-identification (DID) strategies across multiple datasets and templates. This project is particularly useful for data scientists, privacy officers, and researchers working with sensitive data in healthcare, finance, or any field requiring robust data anonymization.

By automating the extraction and analysis of de-identification operations, this tool helps ensure consistency and compliance across different data sources, ultimately enhancing data privacy and reducing the risk of unintended information disclosure.

## Features

- **YAML Configuration Parsing**: Efficiently extract de-identification information from complex YAML configuration files.
- **Cross-Dataset Analysis**: Compare and contrast de-identification operations across multiple datasets to identify inconsistencies.
- **Consistency Scoring**: Utilize advanced algorithms to calculate consistency scores, providing a quantitative measure of de-identification uniformity.
- **Excel Report Generation**: Automatically generate comprehensive Excel reports including:
  - A catalog of all distinct de-identification operations
  - Detailed per-table analysis of differences and consistency scores
  - Overall consistency score for the entire de-identification strategy
- **Modular Architecture**: Easily extensible codebase allowing for the addition of new features and customizations.

## Installation

### Prerequisites

- Python 3.9+
- Poetry (for dependency management)
- Git (for version control and installation)

### Step-by-step Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/deid-info-extractor.git
   cd deid-info-extractor
   ```

2. Install dependencies using Poetry:
   ```bash
   poetry install
   ```

3. Activate the virtual environment:
   ```bash
   poetry shell
   ```

4. Verify the installation:
   ```bash
   python -c "import deid_extractor; print('Installation successful!')"
   ```

## Usage

### Basic Usage

To run a complete analysis:

```bash
python main.py
```

This command will:
1. Load the YAML configuration from `dby_project.yaml`
2. Extract de-identification information
3. Perform a comprehensive analysis
4. Generate an Excel report (`deid_info_analysis.xlsx`)

### Customizing the Analysis

You can customize the analysis by modifying the `constants.py` file. This allows you to adjust:

- File paths for input and output
- Dataset column name mappings
- De-identification operation descriptions
- Excel styling options

## Project Structure

```
deid-info-extractor/
│
├── constants.py         # Global constants and configurations
├── deid_analyzer.py     # Core analysis logic
├── deid_extractor.py    # YAML extraction functionality
├── excel_exporter.py    # Excel report generation
├── main.py              # Main execution script
├── yaml_handler.py      # YAML file handling utilities
│
├── tests/               # Unit and integration tests
├── docs/                # Additional documentation
└── examples/            # Example configurations and outputs
```

## Configuration

The primary configuration file is `dby_project.yaml`. This YAML file defines the de-identification templates and operations for different datasets. 

Example structure:
```yaml
dataset_name:
  tables_to_deid:
    - table_id: example_table
      col_deid_operations:
        - col_id: sensitive_column
          op_name: anonymization_method
```

Refer to `examples/sample_config.yaml` for a complete example.

## Advanced Usage

### Custom De-identification Operations

To add custom de-identification operations:

1. Define the operation in `constants.py`:
   ```python
   DEID_OPERATIONS['new_operation'] = "Description of the new operation"
   ```

2. Implement the logic in `deid_analyzer.py`

3. Update the YAML configuration to use the new operation

### Extending the Analysis

The modular structure allows for easy extension. To add new analysis features:

1. Create a new module (e.g., `advanced_analytics.py`)
2. Import and use it in `main.py`

## Troubleshooting

Common issues and their solutions:

- **YAML parsing errors**: Ensure your YAML is properly formatted. Use a YAML validator if necessary.
- **Missing dependencies**: Run `poetry install` to ensure all dependencies are up to date.
- **Unexpected results**: Check the log files in the `logs/` directory for detailed execution information.

For more help, please open an issue on the GitHub repository.

## Contributing

We welcome contributions! Please follow these steps:

1. Fork the repository
2. Create a new branch (`git checkout -b feature/AmazingFeature`)
3. Make your changes
4. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
5. Push to the branch (`git push origin feature/AmazingFeature`)
6. Open a Pull Request

Please ensure your code adheres to our coding standards and includes appropriate tests.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

For additional information, please refer to the [documentation](docs/). If you find this project useful, please consider giving it a star on GitHub!