# Email Export Parser

A tool to analyze email exports, identify commercial emails, and generate deletion instructions.

## Overview

The Email Export Parser helps you analyze email exports (such as those from Gmail or other providers), identify commercial/promotional emails, and generate reports and instructions for managing or deleting them. This tool is particularly useful for:

- Cleaning up your inbox by identifying promotional emails
- Generating step-by-step instructions for unsubscribing from mailing lists
- Analyzing your email usage patterns

## Features

- **Email Analysis**: Scans through email export directories to identify commercial emails
- **Report Generation**: Creates CSV reports of commercial emails with details
- **Deletion Instructions**: Generates step-by-step instructions for managing identified emails
- **Multilingual Support**: Provides instructions in English, French, or both
- **Dual Interface**: Offers both a command-line interface (CLI) and a graphical user interface (GUI)

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/email-export-parser.git
   cd email-export-parser
   ```

2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

### Graphical User Interface (GUI)

To use the GUI, simply run the script without any arguments:

```
python main.py
```

The GUI allows you to:
- Select the email export directory
- Choose an output directory for reports
- Select the language for deletion instructions
- Monitor progress with a status log and progress bar

### Command Line Interface (CLI)

For command-line usage:

```
python main.py [export_directory] [output_directory] [language]
```

Arguments:
- `export_directory`: Path to the directory containing email export files
- `output_directory` (optional): Path where reports will be saved (defaults to export directory)
- `language` (optional): Language for deletion instructions (`en`, `fr`, or `both`, defaults to `en`)

Example:
```
python main.py ~/Downloads/Takeout/Mail ~/Documents/EmailReports en
```

## Email Export Format

This tool is designed to work with standard email export formats, particularly those from Gmail's Google Takeout service. The export should contain .eml or .mbox files.

## Output Files

The tool generates two main output files:

1. `commercial_emails_report.csv`: A detailed CSV report of all identified commercial emails
2. `deletion_instructions.txt`: Step-by-step instructions for managing the identified emails

## Requirements

- Python 3.6 or higher
- Dependencies listed in `requirements.txt`

## License

[MIT License](LICENSE)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Created by

This tool was created by [Jérémy Pastouret](https://github.com/jenovateurs) and made with ❤️ by [Les E-novateurs](https://les-enovateurs.com).

![Les E-novateurs](assets/les-enovateurs-logo.webp)

### Follow Us
- [LinkedIn](https://www.linkedin.com/company/les-enovateurs)
- [Mastodon](https://mastodon.social/@enovateurs_media)

