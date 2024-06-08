# Excel Remit Processor

This Python project processes a given remit Excel file using a helper and template file. The final template is then renamed to the remit file with the current date's name.

## Project Structure

The project consists of the following files:

- **helper.xlsx**: The helper Excel file.
- **Template.xlsx**: The template Excel file.
- **Input Remit.xlsx**: The input remit Excel file to be processed.
- **script.bat**: Batch script to open and save the helper file.
- **script-2.bat**: Batch script to generate the final remit file.

## Requirements

Ensure you have the following Python modules installed:

- `os`
- `openpyxl`
- `pandas`
- `datetime`
- `time`
- `tqdm`
- `threading`

You can install the required modules using `pip`:

```bash
pip install openpyxl pandas tqdm datetime time tqdm threading
```
## Usage

Follow these steps to run the project:

1. **Run** script.bat:

This will open the helper.xlsx file.
Simply save the file by pressing Ctrl + S and then close it.

2. **Run** script-2.bat:

This will process the input remit file and generate the final remit file.
The final remit file will be automatically saved in the current date folder with the name **Weekly Remit <Current Date>.xls.**

## Important Note
The Final Remit.xls file may show a warning. You need to manually save the Remit.xlsx file in XLS format again to ensure compatibility.

## Example

Here's a quick example of how to use the batch scripts:

- Double-click script.bat.
- Press Ctrl + S to save the helper.xlsx file and close it.
- Double-click script-2.bat.
- Navigate to the current date folder to find the Final Remit.xls file.
- Manually save the Final Remit.xls file as Remit.xlsx in XLS format if any warning appears.
## Contributing
If you'd like to contribute to this project, feel free to submit a pull request or open an issue on GitHub.