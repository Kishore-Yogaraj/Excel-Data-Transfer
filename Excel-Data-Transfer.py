import openpyxl
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Load source workbook
source_workbook = openpyxl.load_workbook('sourcefile.xlsm')
source_sheet = source_workbook.active

# Load destination workbook
destination_workbook = openpyxl.load_workbook('destination.xlsx')
destination_sheet = destination_workbook.active

# Read values from source cells
value_a1 = source_sheet['AB496'].value
value_a2 = source_sheet['AB504'].value
value_a3 = source_sheet['AB498'].value
value_a4 = source_sheet['AB501'].value
value_a5 = source_sheet['AB502'].value

# Write values to destination cells
destination_sheet['BB2'] = value_a1
destination_sheet['BB4'] = value_a2
destination_sheet['BB6'] = value_a3
destination_sheet['BB8'] = value_a4
destination_sheet['BB10'] = value_a5

# Save changes to destination workbook
destination_workbook.save('destination.xlsx')
# Save changes to destination workbook
destination_workbook.save('destination.xlsx')


class ExcelHandler(FileSystemEventHandler):  # Defining a custom class ExcelHandler that inherits from the FileSystemEventHandler class
    def on_modified(self, event):  # Overriding the on_modified method of the parent class
        try:
            while True:  # Entering an infinite loop
                time.sleep(1)  # Sleeping for 1 second
        except KeyboardInterrupt:  # Handling the KeyboardInterrupt exception
            observer.stop()

        observer.join()


# Create watchdog observer
observer = Observer()
event_handler = ExcelHandler()
observer.schedule(event_handler, path='.', recursive=False) # Scheduling the event_handler to be called whenever a file is modified in the current directory
observer.start()
