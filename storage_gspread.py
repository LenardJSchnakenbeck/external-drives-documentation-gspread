import gspread
import gspread_formatting
import pandas as pd
from colorsys import hsv_to_rgb
from gspread_formatting import ConditionalFormatRule, GridRange, BooleanRule, BooleanCondition, CellFormat, \
    textFormat, Color, get_conditional_format_rules

"""Google Sheets integration for external drives documentation.

This module provides functionality to interact with Google Sheets for managing
and formatting external drives documentation. It includes:

- SpreadsheetDocu: Class for spreadsheet operations and data management
- SpreadsheetFormatter: Handles spreadsheet formatting and styling
- Helper functions for color generation and text formatting
"""

def generate_distinct_colors(n):
    """
    Generates n colors with different hues (1/n, 2/n, ..., n/n) and alternating saturation (0.5 & 0.35) and value/brightness (0.6 & 0.75).
    """
    return [
        hsv_to_rgb(i / n, 0.35 + 0.15 * ((i + 1) % 2), 0.6 + 0.15 * (i % 2))
        for i in range(n)
    ]


class SpreadsheetDocu:
    """
    Handles interactions with a Google Spreadsheet.

    Attributes:
        spreadsheet (str): Name of the spreadsheet.
        gc (gspread.Client): Google Sheets API client.
        sh (gspread.Spreadsheet): Opened spreadsheet object.
        worksheet (gspread.Worksheet): First worksheet (drives documentation).
        drives_col (str): Column name for drive names.
        dir_col (str): Column name for project names.
    """
    def __init__(self, spreadsheet_name: str):
        """
        Initialize the SpreadsheetDocu object.

        Args:
            spreadsheet_name (str): Name of the Google Spreadsheet document.
        """
        self.spreadsheet = spreadsheet_name
        self.gc = gspread.service_account()
        self.sh = self.gc.open(self.spreadsheet)
        self.worksheet = self.sh.get_worksheet(0)

        self.drives_col = "drive-name"
        self.dir_col = "project-name"

    def fetch_online_docu(self) -> pd.DataFrame:
        """fetch online drives documentation and return as DataFrame"""
        return pd.DataFrame(self.worksheet.get_all_records())

    def load_blacklist(self) -> tuple[set[str], set[str]]:
        """
        load blacklisted drives and directories from Google Spreadsheet and return as tuple of sets.
        The blacklisted are fetched from the second Worksheet (index 1) with two columns: "blacklist drives" and "blacklist folders".

        Returns:
            tuple[set[str], set[str]]: tuple of sets of blacklisted drives and directories
        """
        worksheet_blacklist = self.sh.get_worksheet(1)
        df = pd.DataFrame(worksheet_blacklist.get_all_records())
        df = df.replace("", pd.NA)  # different length of columns are filled with ""
        bl_drives = df["blacklist drives"].dropna().to_list()
        bl_directories = df["blacklist folders"].dropna().to_list()
        return set(bl_drives), set(bl_directories)

    def apply_blacklist_on_df(self, docu: pd.DataFrame, bl_drives: set, bl_directories: set) -> pd.DataFrame:
        """
        scans df for entries that are in the blacklists and deletes them.

        Args:
            docu (pd.DataFrame): DataFrame to be filtered
            bl_drives (set): Set of blacklisted drives
            bl_directories (set): Set of blacklisted directories

        Returns:
            pd.DataFrame: Filtered DataFrame
        """
        return docu[
            (~docu[self.drives_col].isin(bl_drives)) &
            (~docu[self.dir_col].isin(bl_directories))
        ]

    def apply_blacklist_online(self, bl_drives: set, bl_directories: set):
        """apply blacklist on online spreadsheet"""
        df = self.fetch_online_docu()
        df_filtered = self.apply_blacklist_on_df(df, bl_drives, bl_directories)
        if not df.equals(df_filtered):
            self.update_online_spreadsheet(df_filtered)

    def _update_docu_offline(self, downloaded_docu: pd.DataFrame, connected_drives_docu: pd.DataFrame) -> pd.DataFrame:
        """
        updates the downloaded_docu with the connected_devices_docu

        Args:
            downloaded_docu:
            connected_drives_docu:

        Returns:

        """
        if self.drives_col in connected_drives_docu.keys():
            connected_drives = connected_drives_docu[self.drives_col].unique()
            downloaded_docu = downloaded_docu[
                ~downloaded_docu[self.drives_col].isin(connected_drives)
            ]
        return pd.concat([downloaded_docu, connected_drives_docu], ignore_index=True)

    def _upload_docu(self, docu):
        self.worksheet.clear()
        self.worksheet.update([docu.columns.values.tolist()] + docu.values.tolist())

    def update_online_spreadsheet(self, connected_drives_docu):
        """
        Update the online spreadsheet with the connected drives' documentation.

        Args:
            connected_drives_docu (pd.DataFrame): A DataFrame containing the properties of the connected drives.

        Returns:
            None
        """
        downloaded_docu = self.fetch_online_docu()
        updated_docu = self._update_docu_offline(downloaded_docu, connected_drives_docu)
        blacklist_drives, blacklist_directories = self.load_blacklist()
        updated_docu = self.apply_blacklist_on_df(updated_docu, blacklist_drives, blacklist_directories)
        self._upload_docu(updated_docu)


class SpreadsheetFormatter:
    """
    Handles formatting of a Google Spreadsheet.

    Attributes:
        spread (SpreadsheetDocu): SpreadsheetDocu object to format.
    """
    def __init__(self, spread: SpreadsheetDocu):
        """
        Initializes SpreadsheetFormatter object.

        Args:
            spread (SpreadsheetDocu): SpreadsheetDocu object to format.

        Returns:
            None
        """
        self.spread = spread

    def get_column_id_from_column_position(self, position: int) -> str:
        """
        convert column index to spreadsheet column id (e.g. 0 -> 'A', 1 -> 'B', 27 -> 'AB', etc.)

        Args:
            position:

        Returns:
            str: spreadsheet column id
        """
        if position > 25:
            return (self.get_column_id_from_column_position(position // 26 -1)
                    + self.get_column_id_from_column_position(position % 26))
        return chr(ord('A') + position)

    @staticmethod
    def create_conditional_formatting_rule_text_eq(worksheet, text_match, color, cell_range) -> ConditionalFormatRule:
        """
        Create a conditional formatting rule for a Google Spreadsheet.

        The rule sets the background color (and bold text) for every cell
        in the given `cell_range` whose value exactly matches `text_match`.
        """
        return ConditionalFormatRule(
            ranges=[GridRange.from_a1_range(cell_range, worksheet)],
            booleanRule=BooleanRule(
                condition=BooleanCondition('TEXT_EQ', [text_match]),
                format=CellFormat(textFormat=textFormat(bold=True), backgroundColor=Color(*color))
            )
        )

    @staticmethod
    def remove_rules_by_column(rules: gspread_formatting.ConditionalFormatRules, column_position: int) -> gspread_formatting.ConditionalFormatRules:
        """
        Removes all conditional formatting rules from the given list that apply to the given column position.

        Args:
            rules (ConditionalFormatRules): List of conditional formatting rules.
            column_position (int): Column position (index) to remove rules from.

        Returns:
             ConditionalFormatRules: List of conditional formatting rules with rules applying to the given column position removed.
        """
        for rule in rules[:]:
            if rule.ranges[0].startColumnIndex == column_position \
                    and rule.ranges[0].endColumnIndex == column_position + 1:
                rules.remove(rule)
        return rules

    def color_unique_cells_by_column(self, column_name):
        """
        Colors unique cells in the given column of the spreadsheet.
        Excludes first row (header).

        Args:
            column_name (str): Name of the column to color unique cells in.
        """
        rules = get_conditional_format_rules(self.spread.worksheet)
        df = self.spread.fetch_online_docu()
        column_values = df[column_name].unique()
        colors = generate_distinct_colors(len(column_values))
        column_position = df.columns.get_loc(column_name)
        column_id_gspread = self.get_column_id_from_column_position(column_position)
        range_of_column_gspread = f"{column_id_gspread}2:{column_id_gspread}{max(990, len(df))}"

        rules = self.remove_rules_by_column(rules, column_position)
        for value, color in zip(column_values, colors):
            rules.append(
                self.create_conditional_formatting_rule_text_eq(self.spread.worksheet, value, color, range_of_column_gspread)
            )
        rules.save()

    def format_drives_column(self):
        self.color_unique_cells_by_column(self.spread.drives_col)


if __name__ == "__main__":
    spread_docu_utils = SpreadsheetDocu("external_drives_docu")
    spread_docu_formatter = SpreadsheetFormatter(spread_docu_utils)

    blacklist_drives, blacklist_directories = spread_docu_utils.load_blacklist()
    spread_docu_utils.apply_blacklist_online(blacklist_drives, blacklist_directories)
    spread_docu_formatter.color_unique_cells_by_column(spread_docu_utils.drives_col)

