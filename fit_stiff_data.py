import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import re
import sys

from matplotlib.widgets import SpanSelector
from matplotlib.widgets import Button

filename = "4 Punkt Biegeversuch-Stab_4.xlsx"

translate = {"Prüfzeit" : "Test_time",
             "Standardkraft" : "Force",
             "Verformung": "Deformation"}

row_skip = [2,3]

class RunStiffSelector:
    def __init__(self, filename, translate, row_skip):
        self.xlxs = pd.ExcelFile(filename)

        self.data = {}
        
        num_loaded = 0

        for name in reversed(self.xlxs.sheet_names):
            if not re.match("^[A-Z]\\d+$", name): continue

            sample = pd.read_excel(self.xlxs, sheet_name=name, header=1, usecols=translate.keys(), skiprows=row_skip)

            sample = sample.rename(translate, axis=1)

            self.data[name] = sample

        print("Stiffness data loaded with %d samples" % len(self.data))

        self.fig, self.ax = plt.subplots(1, figsize = (8, 6))
        self.ax.set_title("Welcome to the stiffness selector, press enter to begin")

        self.fig.canvas.mpl_connect('key_press_event', self.on_press)


    def run(self):
        plt.show()

    def update_plot(self):
        sample_name, df = self.data.popitem()
        
        self.current_name = sample_name
        self.current_x = df["Deformation"].to_numpy()
        self.current_y = df["Force"].to_numpy()

        self.ax.plot(self.current_x, self.current_y)
        self.ax.relim()

    def on_press(self, event):
        sys.stdout.flush()
        if event.key == "enter":
            self.update_plot()

if __name__ == "__main__":
    stiff = RunStiffSelector(filename, translate, row_skip)
    stiff.run()