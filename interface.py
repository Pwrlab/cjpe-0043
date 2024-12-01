import tkinter as ttk
from tkinter import filedialog, messagebox
import subprocess


class Application(ttk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Refilling Tool")
        # self.geometry("400x200")
        self.cmd = 'refilling.exe'
        self.create_widgets()

    def create_widgets(self):
        self.input_label = ttk.Label(self, text="Input:")
        self.input_entry = ttk.Entry(self)
        self.input_button = ttk.Button(self, text="Select File", command=lambda: self.select_file(self.input_entry))
        self.output_label = ttk.Label(self, text="Output:")
        self.output_entry = ttk.Entry(self)
        self.output_button = ttk.Button(self, text="Select Folder",
                                        command=lambda: self.select_folder(self.output_entry))

        self.run_button = ttk.Button(self, text="RUN!", command=lambda: self.run())
        self.method_label = ttk.Label(self, text="Method:")

        self.refilling_button_state = ttk.IntVar()
        self.exp_button_state = ttk.IntVar()
        self.line_button_state = ttk.IntVar()
        self.et_button_state = ttk.IntVar()
        self.verbose_button_state = ttk.IntVar()

        self.refilling_button = ttk.Checkbutton(self, text="refilling", variable=self.refilling_button_state, onvalue=1,
                                                offvalue=0)
        self.exp_button = ttk.Checkbutton(self, text="exp", variable=self.exp_button_state, onvalue=1, offvalue=0)
        self.line_button = ttk.Checkbutton(self, text="line", variable=self.line_button_state, onvalue=1, offvalue=0)
        self.et_button = ttk.Checkbutton(self, text="et", variable=self.et_button_state, onvalue=1, offvalue=0)
        self.verbose_button = ttk.Checkbutton(self, text="output fitting data", variable=self.verbose_button_state,
                                              onvalue=1, offvalue=0)

        self.verbose_label = ttk.Label(self, text="output option:")

        self.input_label.grid(row=0, column=0)
        self.input_entry.grid(row=0, column=1)
        self.input_button.grid(row=0, column=2)
        self.output_label.grid(row=1, column=0)
        self.output_entry.grid(row=1, column=1)
        self.output_button.grid(row=1, column=2)

        self.method_label.grid(row=3, column=0)
        self.refilling_button.grid(row=4, column=0)
        self.exp_button.grid(row=4, column=1)
        self.line_button.grid(row=4, column=2)
        self.et_button.grid(row=4, column=3)
        self.verbose_label.grid(row=5, column=0)
        self.verbose_button.grid(row=6, column=0, columnspan=2)
        self.run_button.grid(row=7, column=0, columnspan=3)
        create_tool_tip(self.output_entry, "选择结果输出的文件地址.")

    def select_file(self, entry):
        filename = filedialog.askopenfilename()
        entry.delete(0, ttk.END)
        entry.insert(0, filename)

    def select_folder(self, entry):
        folder_name = filedialog.askdirectory()
        entry.delete(0, ttk.END)
        entry.insert(0, folder_name)

    def run(self):
        # self.cmd = 'python main.py'
        self.cmd = 'refilling.exe'
        self.generate_command()
        # self.cmd = 'python main.py'

        process = subprocess.Popen(self.cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        print(process.stdout.read().decode('gbk'))
        # 实时打印输出

    def generate_command(self):
        if self.input_entry.get() != '':
            input_file = self.input_entry.get()
            # 判断是不是有效的xlsx
            if not input_file.endswith('.xlsx'):
                messagebox.showerror("Error", "Please input the correct file.")
                return False
            self.cmd += f" -i {self.input_entry.get()}"
        if self.output_entry.get() != '':
            self.cmd += f" -o {self.output_entry.get()}"
        if self.refilling_button_state.get():
            print(self.refilling_button_state.get())
            self.cmd += f" -r"
        if self.exp_button_state.get():
            self.cmd += f" -e"
        if self.line_button_state.get():
            self.cmd += f" -l"
        if self.et_button_state.get():
            self.cmd += f" -t"
        if self.verbose_button_state.get():
            self.cmd += f" -v"


class ToolTip(object):
    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def show_tip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 27
        y = y + self.widget.winfo_rooty() + 27
        self.tipwindow = tw = ttk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))

        label = ttk.Label(tw, text=self.text, justify=ttk.LEFT,
                          background="#ffffe0", relief=ttk.SOLID, borderwidth=1,
                          font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()


def create_tool_tip(widget, text):
    toolTip = ToolTip(widget)

    def enter(event):
        toolTip.show_tip(text)

    def leave(event):
        toolTip.hide_tip()

    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)


if __name__ == '__main__':
    a = Application()
    a.mainloop()
