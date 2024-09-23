# PST to Text Converter

## Description

This project provides a Python script to convert Microsoft Outlook PST (Personal Storage Table) files into plain text format. It's designed to process both small and large PST files efficiently, making the email content accessible for further analysis or integration with other systems.

## Features

- Converts PST files to plain text, preserving email metadata and content
- Handles large PST files with efficient memory usage
- Provides detailed logging of the conversion process
- Generates output files with names correlating to input PST files

## Prerequisites

Before you begin, ensure you have met the following requirements:

- Python 3.6 or higher

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/pst-to-text-converter.git
   cd pst-to-text-converter
   ```

2. Install the required Python packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

To use the PST to Text Converter, follow these steps:

1. Place your PST file in the project directory or note its path.

2. Run the script with Python:
   ```
   python pst_to_text.py
   ```

3. When prompted, enter the path to your PST file and the desired output folder.

4. The script will process the PST file and create a text file containing all extracted emails.

## Configuration

You can modify the following variables in the script to adjust its behavior:

- Logging level: Change `logging.basicConfig(level=logging.INFO, ...)` to adjust verbosity

## Contributing

Contributions to the PST to Text Converter are welcome. To contribute:

1. Fork the repository.
2. Create a new branch: `git checkout -b <branch_name>`.
3. Make your changes and commit them: `git commit -m '<commit_message>'`
4. Push to the original branch: `git push origin <project_name>/<location>`
5. Create the pull request.

Alternatively, see the GitHub documentation on [creating a pull request](https://help.github.com/articles/creating-a-pull-request/).

## License

This project uses the following license: [MIT License](https://opensource.org/licenses/MIT).

## Contact

If you want to contact the maintainer of this project, please email <roger.l.jiahao@gmail.com>.

## Acknowledgements

- [pypff](https://github.com/libyal/libpff) - Library for accessing PST files
- [html2text](https://github.com/Alir3z4/html2text) - Library for converting HTML to plain text