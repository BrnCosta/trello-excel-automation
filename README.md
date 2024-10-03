# Excel Automation to Trello Board

## Overview

This project automates the process of sending content from an Excel worksheet to a Trello board via email. The script reads the data from a specified range in an Excel sheet, processes it, and then sends each record as an email to a Trello-specific email address.

It's designed to be easy to use, requiring minimal setup beyond filling in the necessary input arguments, such as the path to the workbook and the email credentials.

## How It Works

1. **Read Data from Excel**: The script extracts information from a defined selection of cells in the Excel worksheet.
2. **Prepare Emails**: The script creates individual emails with a subject and body from the extracted data.
3. **Send Emails to Trello**: Each email is sent to a designated Trello email address, which can be used to create Trello cards.

### Key Features:
- Easy way to automate the creation of Trello cards from Excel data.
- Customizable email content based on Excel cell values.
- Error handling to avoid sending emails if any issue occurs with the input data.

## Prerequisites

- **Python 3.8+**
- Install the required libraries using the `requirements.txt` file:
  ```bash
  pip install -r requirements.txt
  ```

## Arguments

- `-e`, `--email`: The sender's email address (used for authentication).
- `-p`, `--password`: The sender's email password (used to log into the email account).
- `-t`, `--trello_email`: The Trello email address where emails will be sent.
- `-b`, `--workbook`: The file path to the Excel workbook.
- `-w`, `--worksheet`: The worksheet name in the Excel file.
- `-s`, `--selection`: The cell range to read data from (e.g., `A1:F10`).

## Example Usage

```bash
python send_to_trello.py -e "your_email@gmail.com" -p "your_password" \
-t "trello_board@cards.trello.com" -b "~/path/to/excel.xlsx" \
-w "Sheet1" -s "A1:C5"
```

## Windows Users

For Windows users, an executable file is available for download. You can find the `.exe` file in the [GitHub Releases](https://github.com/) section of this repository. This allows you to run the tool without needing to install Python or additional libraries.

Simply download the `.exe` file, run it, and provide the required inputs. This version is a ready-to-use option for those unfamiliar with Python or who prefer a simpler setup.

## Notes

- Ensure your email provider allows SMTP connections.
- Make sure Trello's email feature is set up correctly for your board.
- The email body will contain the content of the cells in the selected range, formatted with `|` between cell values in the same row.