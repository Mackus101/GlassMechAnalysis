import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import re
import sys
import time
from scipy import stats

from matplotlib.widgets import SpanSelector
from matplotlib.widgets import Button

filename = "Collated Filament Test Data.xlsx"

translate = {"Prüfzeit" : "test_time",
             "Standardkraft" : "force",
             "Standardweg": "deformation"}

row_skip = [2,3]

## Figure Formatting

output_file_fmt = "Stiffness\\Rod_stiffness_{}.csv"

figure_file_fmt= "Stiff_Figs\\Stiff_line_{}.png"

top_y_title = "Force (N)"
top_x_title = "Deflection (mm)"

bot_y_title = "Selected Force (N)"
bot_x_title = "Selected Deflection (mm)"


suptitle_fmt = "Sample: {}"
curve_stats_fmt = "Slope = {:.2f}\t Intercept = {:.2f}\t $r^{{2}}$ = {:.2f}"

class RunStiffSelector:
    def __init__(self, filename, translate, row_skip):
        self.xlxs = pd.ExcelFile(filename)

        self.data = {}

        self.output = pd.DataFrame()

        self.OUTPUT_FILENAME = output_file_fmt.format(time.strftime("%Y%m%d-%H%M"))

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

        print("Stiffness data loaded with %d samples\n" % self.NUM_SAMPLES)

        self.fig, (self.ax_top, self.ax_bot) = plt.subplots(2, figsize = (8, 6))
        self.fig.suptitle("Welcome to the stiffness selector, press enter to begin")
        self.fig.tight_layout()

        self.fig.canvas.mpl_connect('key_press_event', self.on_press)
        self.fig.canvas.mpl_connect('close_event', self.on_close)

        self.span = SpanSelector(self.ax_top, self.on_select, "horizontal", useblit=True, props=dict(alpha=0.5, facecolor="tab:blue"), interactive=True, drag_from_anywhere=True)


    def run(self):
        plt.show()

    def update_plot(self):
        try:
            row_format = {
                "sample"                    :[self.current_name],
                "stiff_slope"               :[self.reg.slope],
                "stiff_intercept"           :[self.reg.intercept],
                "stiff_rvalue"              :[self.reg.rvalue],
                "stiff_rsquared"            :[self.reg.rvalue**2],
                "stiff_stderr"              :[self.reg.stderr],
                "stiff_intercept_stderr"    :[self.reg.intercept_stderr]
            }
            new_row = pd.DataFrame(row_format)
            self.output = pd.concat([self.output, new_row], ignore_index=True)

            fig, ax = plt.subplots(1)
            ax.plot("deformation", "force", data=self.current_data)
            ax.plot(self.fit_region_x, self.reg.intercept + self.reg.slope * self.fit_region_x, "--r")
            
            fig.suptitle(suptitle_fmt.format(self.current_name))
            ax.set_title(curve_stats_fmt.format(self.reg.slope, self.reg.intercept, self.reg.rvalue**2))
            ax.set_ylabel(top_y_title)
            ax.set_xlabel(top_x_title)
            fig.savefig(figure_file_fmt.format(self.current_name))
            plt.close(fig)
        except AttributeError:
            pass

        self.ax_top.cla()
        self.ax_bot.cla()
        self.span.set_visible(False)

        if not self.data:
            self.fig.suptitle("Curve fitting complete, shutting down")
            self.fig.canvas.draw_idle()
            plt.pause(3)
            plt.close("all")
            return

        self.current_name, self.current_data = self.data.popitem()


        df = self.current_data.iloc[0:self.current_data["force"].idxmax(), :]
        
        self.current_x = df["deformation"].to_numpy()
        self.current_y = df["force"].to_numpy()

        self.ax_top.plot(self.current_x, self.current_y)

        self.fig.suptitle(suptitle_fmt.format(self.current_name))
        self.ax_top.set_title(curve_stats_fmt.format(np.nan, np.nan, np.nan))
        self.ax_top.set_ylabel(top_y_title)
        self.ax_top.set_xlabel(top_x_title)
        self.ax_bot.set_ylabel(bot_y_title)
        self.ax_bot.set_xlabel(bot_x_title)

        self.fig.tight_layout()
        self.fig.canvas.draw_idle()

    def on_press(self, event):
        sys.stdout.flush()
        if event.key == "enter":
            if self.span.get_visible():
                self.update_plot()
            else:
                print("No curve selected, please select part of curve before moving on")
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
            self.ax_top.plot(self.fit_region_x, self.reg.intercept + self.reg.slope * self.fit_region_x, "--r")


            self.ax_bot.scatter(self.fit_region_x, self.fit_region_y, marker='.')
            self.ax_bot.plot(self.fit_region_x, self.reg.intercept + self.reg.slope * self.fit_region_x, "--r")

            self.ax_bot.set_xlim(self.fit_region_x[0], self.fit_region_x[-1])
            self.ax_bot.set_ylim(self.fit_region_y.min(), self.fit_region_y.max())

            self.ax_top.set_title(curve_stats_fmt.format(self.reg.slope, self.reg.intercept, self.reg.rvalue**2))

            self.ax_top.set_ylabel(top_y_title)
            self.ax_top.set_xlabel(top_x_title)
            self.ax_bot.set_ylabel(bot_y_title)
            self.ax_bot.set_xlabel(bot_x_title)

            self.fig.canvas.draw_idle()

    def on_close(self, event):
        print("\nFigure closed with %d out of %d curves fitted!\n" % (len(self.output.index), self.NUM_SAMPLES))
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