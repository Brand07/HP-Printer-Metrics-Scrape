# HP Printer Scrape

This project is designed to scrape data from HP printers and store the information in an Excel file. The data includes page count, toner level, and maintenance level for each printer.

## Prerequisites

- Python 3.8 or higher
- `uv` for project management

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/Brand07/hp-printer-scrape.git
    cd hp-printer-scrape
    ```

2. Install `uv`:
    ```sh
    pip install uv
    ```

3. Install project dependencies:
    ```sh
    uv sync
    ```

## Usage

1. Run the script to scrape printer data:
    ```sh
    uv run python main.py
    ```

2. The script will generate or update the `Printer_Metrics.xlsx` file with the latest data.

Example Output:

![Screenshot 2023-10-23 114533](https://github.com/Brand07/HP-Printer-Metrics-Scrape/assets/81128304/580b4bf4-c41e-464d-a54d-453c353698ac)

## Project Structure

- `HP_Printer_Scrape.py`: Main script to scrape printer data and save it to an Excel file.
- `Printer_Metrics.xlsx`: Excel file where the scraped data is stored.
- `uv.lock`: Lock file for `uv` project management.

## Configuration

- Update the `printers` dictionary in `HP_Printer_Scrape.py` with the IP addresses of your printers.
- Modify the tags for scraping if necessary.

### Other Requirements
This script requires that you disable *Encrypt All Web Communication (not including IPP)* on the printer. This can be done from the internal webpage under the Networking Tab >> Mgmt. Protocals.
This allows the script to request the webpage, otherwise the script would fail to reach the webpage.

Do this at your own risk as you are disabling a security feature on your printer
![254900750-857f363f-b7dd-4d55-934e-ab152758a916](https://github.com/Brand07/HP-Printer-Metrics-Scrape/assets/81128304/b7b84f02-1c4e-43da-9cf7-1059ed8ba43b)

## Contributing

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Commit your changes (`git commit -am 'Add new feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Create a new Pull Request.

## License

This project is licensed under the MIT License.









