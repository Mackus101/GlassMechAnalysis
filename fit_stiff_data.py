import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import re
import sys
import time
from scipy import stats

from matplotlib.widgets import SpanSelector
from matplotlib.widgets import Button

filename = "4 Punkt Biegeversuch-Stab_4.xlsx"

translate = {"Prüfzeit" : "Test_time",
             "Standardkraft" : "Force",
             "Verformung": "Deformation"}

row_skip = [2,3]

output_prefix = "Stiffness\\Rod_stiffness_"
output_suffix = ".csv"

class RunStiffSelector:
    def __init__(self, filename, translate, row_skip):
        self.xlxs = pd.ExcelFile(filename)

        self.data = {}

        self.output = pd.DataFrame()

        time_str = time.strftime("%Y%m%d-%H%M")
        self.OUTPUT_FILENAME = output_prefix + time_str + output_suffix

        self.current_x = np.empty(0)
        self.current_y = np.empty(0)

        self.fit_region_x = np.empty(0)
        self.fit_region_y = np.empty(0)

        for name in reversed(self.xlxs.sheet_names):
            if not re.match("^[A-Z]\\d+$", name): continue

            sample = pd.read_excel(self.xlxs, sheet_name=name, header=1, usecols=translate.keys(), skiprows=row_skip)

            sample = sample.rename(translate, axis=1)

            self.data[name] = sample

        self.NUM_SAMPLES = len(self.data)

        print("Stiffness data loaded with %d samples" % self.NUM_SAMPLES)

        self.fig, (self.ax_top, self.ax_bot) = plt.subplots(2, figsize = (8, 6))
        self.fig.suptitle("Welcome to the stiffness selector, press enter to begin")

        self.fig.canvas.mpl_connect('key_press_event', self.on_press)
        self.fig.canvas.mpl_connect('close_event', self.on_close)

        self.span = SpanSelector(self.ax_top, self.on_select, "horizontal", useblit=True, props=dict(alpha=0.5, facecolor="tab:blue"), interactive=True, drag_from_anywhere=True)


    def run(self):
        plt.show()

    def update_plot(self):
        try:
            row_format = {
                "Sample"                    :[self.current_name],
                "Stiff_slope"               :[self.reg.slope],
                "Stiff_intercept"           :[self.reg.intercept],
                "Stiff_rvalue"              :[self.reg.rvalue],
                "Stiff_rsquared"            :[self.reg.rvalue**2],
                "Stiff_stderr"              :[self.reg.stderr],
                "Stiff_intercept_stderr"    :[self.reg.intercept_stderr]
            }
            new_row = pd.DataFrame(row_format)
            self.output = pd.concat([self.output, new_row], ignore_index=True)
        except AttributeError:
            pass

        self.ax_top.cla()
        self.ax_bot.cla()
        self.span.set_visible(False)

        if not self.data:
            self.fig.suptitle("Curve fitting complete, shutting down")
            self.fig.canvas.draw_idle()
            plt.pause(5)
            plt.close()

        sample_name, df = self.data.popitem()
        
        self.current_name = sample_name
        self.current_x = df["Deformation"].to_numpy()
        self.current_y = df["Force"].to_numpy()

        self.ax_top.plot(self.current_x, self.current_y)
        self.fig.suptitle("Sample: %s" % sample_name)
        self.fig.canvas.draw_idle()

    def on_press(self, event):
        sys.stdout.flush()
        if event.key == "enter":
            self.update_plot()
        elif event.key == "escape":
            self.ax_bot.cla()
            self.fig.canvas.draw_idle()

    def on_select(self, xmin, xmax):
        self.span.set_visible(True)
        indmin, indmax = np.searchsorted(self.current_x, (xmin, xmax))
        indmax = min(len(self.current_x) - 1, indmax)

        self.fit_region_x = self.current_x[indmin:indmax]
        self.fit_region_y = self.current_y[indmin:indmax]


        if len(self.fit_region_x) >= 2:
            self.reg = stats.linregress(self.fit_region_x, self.fit_region_y)
            

            self.ax_bot.cla()
            self.ax_top.cla()

            self.ax_top.plot(self.current_x, self.current_y)
            self.ax_top.plot(self.fit_region_x, self.reg.intercept + self.reg.slope * self.fit_region_x, "r")

            self.ax_bot.scatter(self.fit_region_x, self.fit_region_y, marker='.')
            self.ax_bot.plot(self.fit_region_x, self.reg.intercept + self.reg.slope * self.fit_region_x, "r")

            self.ax_bot.set_xlim(self.fit_region_x[0], self.fit_region_x[-1])
            self.ax_bot.set_ylim(self.fit_region_y.min(), self.fit_region_y.max())
            self.fig.canvas.draw_idle()

    def on_close(self, event):
        print("Figure closed with %d out of %d curves fitted\n" % (len(self.output.index), self.NUM_SAMPLES))
        self.save_and_exit()

    def save_and_exit(self):
        while True:
            Prompt = input("Do you wish to save the data to \"%s\"? Answer y or n\n" % self.OUTPUT_FILENAME)
            if Prompt in ['y', 'yes']:
                self.output.to_csv(self.OUTPUT_FILENAME, index=False)
                break
            elif Prompt in ['n', 'no']:
                break
        sys.exit(0)


if __name__ == "__main__":
    stiff = RunStiffSelector(filename, translate, row_skip)
    stiff.run()