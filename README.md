import csv
import pandas as pd
from collections import defaultdict
from typing import Dict, List, Tuple


class FixTag:
    def __init__(self, tag_number: str, tag_name: str, is_enum: bool = False):
        self.tag_number = tag_number
        self.tag_name = tag_name
        self.is_enum = is_enum
        self.enum_values = {}

    def add_enum_value(self, value: str, description: str):
        if self.is_enum:
            self.enum_values[value] = description

    def __str__(self):
        return f"{self.tag_name} ({self.tag_number})"


class FixLogParser:
    def __init__(self, fix_tag_file: str):
        self.fix_tags = self.parse_fix_tags(fix_tag_file)
        self.tag_values = {}
        self.standard_values = {}

    def parse_fix_tags(self, fix_tag_file: str) -> Dict[str, FixTag]:
        fix_tags_df = pd.read_excel(fix_tag_file, sheet_name='FIX_TAG')
        fix_tags = {}
        for tag, tag_name, is_enum in zip(fix_tags_df['Tag'], fix_tags_df['Tag Name'], fix_tags_df['Enum']):
            fix_tag = FixTag(tag, tag_name, is_enum)
            fix_tags[tag] = fix_tag
        return fix_tags

    def parse_fix_log(self, log_lines: List[str]):
        for line in log_lines:
            if '8=' in line:
                for field in line.strip().split('\x01'):
                    if '=' in field:
                        if field.count('=') == 1:
                            tag, value = field.split('=')
                            fix_tag = self.fix_tags.get(tag)
                            if fix_tag:
                                self.tag_values[tag] = value
                                fix_tag.add_enum_value(value, value)
                            else:
                                self.tag_values[tag] = value

    def generate_excel_report(self, output_file: str):
        with open(output_file, 'w', newline='') as fix_log_csv:
            fieldnames = [str(fix_tag) for fix_tag in self.fix_tags.values()]
            writer = csv.DictWriter(fix_log_csv, fieldnames=fieldnames)
            writer.writeheader()

            for tag, value in self.tag_values.items():
                fix_tag = self.fix_tags.get(tag)
                if fix_tag:
                    if fix_tag.is_enum:
                        value = fix_tag.enum_values.get(value, value)
                    row = {str(fix_tag): value for fix_tag in self.fix_tags.values()}
                    writer.writerow(row)
                else:
                    row = {'unknown tag': tag, 'value': value}
                    writer.writerow(row)


if __name__ == '__main__':
    fix_tag_file = 'fix_tags.xlsx'
    fix_log_file = 'fixlog.txt'
    output_file = 'fixlog.csv'

    with open(fix_log_file, 'r') as file:
        log_lines = file.readlines()

    parser = FixLogParser(fix_tag_file)
    parser.parse_fix_log(log_lines)
    parser.generate_excel_report(output_file)
