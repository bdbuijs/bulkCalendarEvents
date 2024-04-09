"""
Bulk Calendar Events

Creates and opens a template Excel sheet, in order to save all events added to it to a .ics file
"""

__author__ = "Bram Buijs"
__version__ = "0.1.0"
__licence__ = "MPL 2.0"
__copyright__ = """Copyright 2024 Bram Buijs

Licensed under the Mozilla Public License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.mozilla.org/en-US/MPL/2.0/

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License."""

import argparse
from datetime import datetime, timedelta
import logging
import pathlib
import subprocess


import icalendar
from icalendar import Calendar, Event
import openpyxl


def main():
    """Main processing function"""
    args = parse_args()
    setup_logging(args)

    if args.file is None:
        # create file
        args.file = create_event_file()

    events = load_events(args.file)

    export_ics(events, args.output)

    delete_temporary_files()

    subprocess.call(["open", "-R", str(args.output)])


def create_event_file() -> pathlib.Path:
    """Creates empty event template file and opens it in Microsoft Excel"""
    filename = pathlib.Path(__file__).parent / "events.xlsx"
    logging.info(f"Creating events sheet at {filename}...")
    headers = ('Title', 'Description', 'Date', 'Start time', 'End time', 'Location')
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(headers)
    workbook.save(filename)
    logging.info("Opening events file...")
    subprocess.call(['open', '-a', '/Applications/Microsoft Excel.app', f'{str(filename)}'])
    _ = input("Add events to the file, save it and press RETURN to continue...")
    return filename


def load_events(filename: pathlib.Path) -> [Event]:
    """Loads events from .xlsx file"""
    events = []
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    rows = sheet.rows
    next(rows)  # skip header
    for title_cell, description_cell, date_cell, start_time_cell, end_time_cell, location_cell in rows:
        title = title_cell.value
        description = description_cell.value
        location = str(location_cell.value)
        date = date_cell.value
        start_time = start_time_cell.value
        end_time = end_time_cell.value
        start_dt = datetime.combine(date, start_time)
        end_dt = datetime.combine(date, end_time)
        events.append(create_calendar_event(title, description, start_dt, end_dt, location))

    return events


def delete_temporary_files():
    """Deletes all temporary files created"""
    events_file = pathlib.Path(__file__).parent / "events.xlsx"
    if events_file.exists():
        events_file.unlink()


def create_calendar_event(title: str, description: str, start_dt: datetime, end_dt: datetime, location=None):
    """Returns calendar event with given name, start, end and location"""
    ev = Event()
    ev.add('summary', title)
    ev.add('description', description)
    ev.add('dtstart', start_dt)
    ev.add('dtend', end_dt)
    ev.add('dtstamp', datetime.now())
    if location is not None:
        ev['location'] = icalendar.vText(location)
    return ev


def export_ics(events: [Event], filename: pathlib.Path):
    """Export events to .ics file"""
    logging.info("Saving calendar to .ics file...")
    cal = Calendar()
    for ev in events:
        cal.add_component(ev)
    with open(filename, 'wb') as f:
        f.write(cal.to_ical())
    logging.info(f"Saved events to {filename}")


def parse_args():
    """Parses script arguments"""

    def ics_file(s: str) -> pathlib.Path:
        """Check that path provided to -f arg is a .xlsx file"""
        p = pathlib.Path(s)
        if (ext := p.suffix.lower()) != ".ics":
            raise argparse.ArgumentTypeError(f"File must be a Calendar file (.ics), but got {ext}")
        return p

    def xlsx_file(s: str) -> pathlib.Path:
        """Check that path provided to -f arg is a .xlsx file"""
        p = pathlib.Path(s)
        if (ext := p.suffix.lower()) != ".xlsx":
            raise argparse.ArgumentTypeError(f"File must be an Excel spreadsheet (.xlsx), but got {ext}")
        return p

    parser = argparse.ArgumentParser(description=__doc__, prog="bulkcal")

    parser.add_argument(
        "--version",
        action="version",
        version=f"Bulk Calendar Events {__version__}")

    parser.add_argument(
        "--debug",
        action="store_true",
        help="Run script with DEBUG level logging (default is INFO)")

    parser.add_argument(
        "--file", "-f",
        type=xlsx_file,
        required=False,
        help="Specify pre-saved template file to load (optional)")

    parser.add_argument(
        "--output", "-o",
        type=ics_file,
        required=False,
        default=pathlib.Path(__file__).parent / "output.ics",
        help=f"Specify output filename (defaults to 'output.ics' in script folder)")

    args = parser.parse_args()
    return args


def setup_logging(args):
    """Sets up logging and returns logging level"""
    if args.debug:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(level=logging.INFO)
    logging_level = logging.getLogger().getEffectiveLevel()
    logging.info(f"Running with logging level {logging_level}")
    logging.debug(f"Script args: {args}")
    return logging_level


if __name__ == '__main__':
    raise SystemExit(main())
