__author__ = "David Marienburg"
__version__ = "1.0"
__LastUpdate__ = "4/9/2019"

import pandas as pd
import numpy as np

from datetime import datetime
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

class LongStayersReport:
    def __init__(self):
        file = askopenfilename(
            title="Open the raw Current Long Stayers report",
            initialdir="//tproserver/Reports/Monthly Reports/"
        )
        self.entries = pd.read_excel(file)
        self.entries_copy = self.entries.copy()
        self.today = datetime.today()
        self.save_output()

    def create_los_columns(self):
        # Fill blank values in the Entry Exit Exit Date column with todays date
        self.entries_copy["Entry Exit Exit Date"].fillna(self.today, inplace=True)

        # Create length of stay columns using the np.timedelta64 method
        self.entries_copy["LOS Years"] = (
            self.entries_copy["Entry Exit Exit Date"] - self.entries_copy["Entry Exit Entry Date"]
        )/np.timedelta64(1, "Y")
        self.entries_copy["LOS Days"] = (
            self.entries_copy["Entry Exit Exit Date"] - self.entries_copy["Entry Exit Entry Date"]
        )/np.timedelta64(1, "D")

        # Use the pandas.groupby method to find the total length of stay for
        # each participant
        grouped = self.entries_copy[
            ["Client Uid", "LOS Years", "LOS Days"]
        ].groupby(by="Client Uid").sum()

        # Return only rows where the total length of stay is longer than 364 days
        return grouped[grouped["LOS Days"] > 364]

    def show_current_location(self):
        los_data = self.create_los_columns()

        # Slice the self.entries dataframe so that it only shows the current
        # participants who are currently entered into a shelter
        current_stayers = self.entries[self.entries["Entry Exit Exit Date"].isna()]

        # Merge the current stayers and los_data dataframes then return the
        # using an inner merge.s
        output = current_stayers[
            [
                "Client Uid",
                "Client First Name",
                "Client Last Name",
                "Entry Exit Provider Id",
                "Entry Exit Entry Date"
            ]
        ].merge(
            los_data.reset_index(),
            on="Client Uid",
            how="inner"
        ).sort_values(by="LOS Years", ascending=False)

        return output

    def save_output(self):
        final_data = self.show_current_location()

        writer = pd.ExcelWriter(
            asksaveasfilename(
                    title="Save the Current Long Stayers report",
                    defaultextension=".xlsx",
                    initialfile="Current Long Stayers(Processed).xlsx",
                    initialdir="//tproserver/Report/Monthly Reports/"
            ),
            engine="xlsxwriter"
        )
        final_data.to_excel(writer, sheet_name="Current Long Stayers", index=False)
        writer.save()

if __name__ == "__main__":
    LongStayersReport()
