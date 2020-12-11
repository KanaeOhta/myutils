import os
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import tkinter.filedialog as dialog


class ConverterWindow(ttk.Frame):
    
    def __init__(self, master):
        super().__init__(master)
        self.set_style()
        self.create_variables()
        self.create_ui()


    def set_style(self):
        s = ttk.Style()
        s.configure('new.TFrame', background='#ffffff')

    def create_variables(self):
        self.file_path = tk.StringVar()


    def create_ui(self):
        self.create_main_frame()
        self.create_option_frame()
        self.create_footer_frame()

    def create_main_frame(self):
        frame = ttk.Frame(self.master)
        frame.grid(row=0, column=0, pady=20, padx=20, sticky=(tk.W, tk.E, tk.N, tk.S))
        filename_label = ttk.Label(frame, text='File name: ')
        filename_label.grid(row=0, column=0, sticky=tk.E)
        path_entry = ttk.Entry(frame, textvariable=self.file_path, width=70)
        path_entry.grid(row=0, column=1, sticky=tk.E)
        open_button = ttk.Button(frame, text='Open', command=self.open)
        open_button.grid(row=0, column=2, sticky=tk.E)
     

    def create_option_frame(self):
        self.option_frame = ttk.Frame(self.master)
        self.option_frame.grid(row=1, column=0, pady=5, padx=20, sticky=(tk.W, tk.E, tk.N, tk.S))
      

    def create_footer_frame(self):
        frame = ttk.Frame(self.master)
        frame.grid(row=3, column=0, pady=20, padx=20, sticky=(tk.W, tk.E, tk.N, tk.S))
        close_button = ttk.Button(frame, text='Close', command=self.close)
        close_button.pack(side=tk.RIGHT)
        convert_button = ttk.Button(frame, text='Convert', command=self.convert)
        convert_button.pack(side=tk.RIGHT, padx=5)
        option_button = ttk.Button(frame, text='Option >>')
        option_button.pack(side=tk.RIGHT, padx=5)


    def convert(self):
        if not (convert_file := self.file_path.get()):
            messagebox.showerror('Error', 'No file is selected.')
            return
       

    def open(self):
        initialdir = os.path.abspath(os.path.dirname(__file__))
        filetypes = [('Data file', '*.xlsx;*.json')]
        file_path = dialog.askopenfilename(filetypes=filetypes,
            initialdir=initialdir)
        if file_path:
            self.file_path.set(file_path)


    def close(self, event=None):
        self.quit()


def main():
    application = tk.Tk()
    application.withdraw()
    application.title('Converter')
    application.option_add('*tearOff', False)
    window = ConverterWindow(application)
    application.protocol('WM_DELETE_WINDOW', window.close)
    application.deiconify()
    application.mainloop()


if __name__ == '__main__':
    main()