from storage_json import load_json, save_json
from drives_scanner import scan_valid_drives_to_df, scan_valid_drives_to_dict

DOCUMENTATION_PATH = "drives_documentation.json"
SPREADSHEET_NAME = "external_drives_docu"  # TODO: rename to "external drives documentation"


def scan_drives_and_update_documentation_json():
    """Scan all valid drives and their valid directories and update the documentation JSON file."""
    drives_documentation_json = load_json(DOCUMENTATION_PATH)
    current_external_drives_properties = scan_valid_drives_to_dict()
    drives_documentation_json.update(current_external_drives_properties)
    save_json(DOCUMENTATION_PATH, drives_documentation_json)


def main():
    """
    Main function to run the application.

    This function will:
      - Load blacklist from Spreadsheet
      - Scan valid drives
      - Update the Spreadsheet with the scanned drives
      - Format the Spreadsheet with a conditional formatting rule
      - Return the updated Spreadsheet (optional)
    """
    from storage_gspread import SpreadsheetDocu, SpreadsheetFormatter

    spread_utils = SpreadsheetDocu(SPREADSHEET_NAME)
    blacklist_drives, blacklist_directories = spread_utils.load_blacklist()
    drives_df = scan_valid_drives_to_df(blacklist_drives, blacklist_directories)
    spread_utils.update_online_spreadsheet(drives_df)
    spread_formatter = SpreadsheetFormatter(spread_utils)
    spread_formatter.format_drives_column()
    # return spread_utils.fetch_online_docu()


if __name__ == "__main__":
    pass
    main()

