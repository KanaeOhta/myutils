import os
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import tkinter.filedialog as dialog

from jsonexcel import ToExcel, FromExcel, file_check, ExtensionError


class ConverterWindow(ttk.Frame):
    
    def __init__(self, master):
        super().__init__(master)
        self.create_variables()
        self.create_ui()


    # def set_style(self):
    #     s = ttk.Style()
    #     s.configure('new.TFrame', background='#ffffff')


    def create_variables(self):
        self.file_path = tk.StringVar()
        self.edit_key = tk.StringVar()
        # self.key_array = tk.StringVar()
    #     self.file_path.trace_add('write', self.path_entered)
        

    # def path_entered(self, *args):
    #     """ _name is automatically set unique name."""
    #     target = None
    #     if args[0] == self.file_path._name:
    #         target = self.file_path.get()
       
        
    def create_ui(self):
        self.create_base_frame()
        self.create_tab_toexcel()
        self.create_tab_fromexcel()
        

    def create_base_frame(self):
        frame = ttk.Frame(self.master)
        frame.pack(fill=tk.BOTH, expand=True)
        note = ttk.Notebook(frame)
        note.pack(fill=tk.BOTH, expand=True, pady=20, padx=20)
        self.tab_toexcel = ttk.Frame(note)
        self.tab_fromexcel = ttk.Frame(note)
        note.add(self.tab_toexcel, text='ToExce')
        note.add(self.tab_fromexcel, text='ToExce')
        close_button = ttk.Button(frame, text='Close', command=self.close)
        close_button.pack(side=tk.RIGHT, pady=10, padx=20)


    def create_tab_toexcel(self):
        main_frame = ttk.Frame(self.tab_toexcel)
        main_frame.pack(fill=tk.BOTH, expand=True, pady=20, padx=20)
        filename_label = ttk.Label(main_frame, text='File name ')
        filename_label.grid(row=0, column=0, sticky=tk.E)
        path_entry = ttk.Entry(main_frame, textvariable=self.file_path, width=70, state='readonly')
        path_entry.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E))
        open_button = ttk.Button(main_frame, text='Open', command=self.open_json)
        open_button.grid(row=0, column=3, sticky=tk.E)

        label_frame  = ttk.LabelFrame(main_frame, text=' Select keys, if you want export selected data. ')
        label_frame.grid(row=1, column=0, columnspan=4, pady=30, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.key_box = tk.Listbox(label_frame, selectmode='multiple')
        self.key_box.grid(row=0, column=0, pady=20, padx=(20, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scroll = ttk.Scrollbar(label_frame, orient=tk.VERTICAL, command=self.key_box.yview)
        self.key_box['yscrollcommand'] = y_scroll.set
        y_scroll.grid(row=0, column=1, pady=20, padx=(0, 20), sticky=(tk.N, tk.S, tk.W))

        self.convert_button = ttk.Button(main_frame, text='Convert', command=self.to_excel, state=tk.DISABLED)
        self.convert_button.grid(row=2, column=2)
        self.deselect_button = ttk.Button(main_frame, text='Deselect', command=self.deselect, state=tk.DISABLED)
        self.deselect_button.grid(row=2, column=3)

        label_frame.columnconfigure(0, weight=1)
        label_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)


    def create_tab_fromexcel(self):
        main_frame = ttk.Frame(self.tab_fromexcel)
        main_frame.pack(fill=tk.BOTH, expand=True, pady=20, padx=20)
        filename_label = ttk.Label(main_frame, text='File name ')
        filename_label.grid(row=0, column=0, sticky=tk.E)
        path_entry = ttk.Entry(main_frame, textvariable=self.file_path, width=70, state='readonly')
        path_entry.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E))
        open_button = ttk.Button(main_frame, text='Open', command='')
        open_button.grid(row=0, column=3, sticky=tk.E)

        self.convert_button = ttk.Button(main_frame, text='Convert', command=self.to_excel, state=tk.DISABLED)
        self.convert_button.grid(row=2, column=2)
        self.deselect_button = ttk.Button(main_frame, text='Deselect', command=self.deselect, state=tk.DISABLED)
        self.deselect_button.grid(row=2, column=3)

        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)

        label_frame  = ttk.LabelFrame(main_frame, text=' Edit keys, if you need. ')
        label_frame.grid(row=1, column=0, columnspan=4, pady=30, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.keys_box = tk.Listbox(label_frame)
        self.keys_box.grid(row=0, rowspan=3, column=0, pady=20, padx=(20, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scroll_keys = ttk.Scrollbar(label_frame, orient=tk.VERTICAL, command=self.keys_box.yview)
        self.keys_box['yscrollcommand'] = y_scroll_keys.set
        y_scroll_keys.grid(row=0, rowspan=3, column=1, pady=20, padx=(0, 10), sticky=(tk.N, tk.S, tk.E))

        self.edit_entry = ttk.Entry(label_frame, textvariable=self.edit_key, width=20)
        self.edit_entry.grid(row=0, column=2, pady=(20, 0), padx=(20, 0), sticky=(tk.W, tk.N, tk.E))
        ok_button = ttk.Button(label_frame, text='OK', width=4, command='')
        ok_button.grid(row=0, column=3, pady=(19, 0), sticky=(tk.N, tk.E))
        allow_label = ttk.Label(label_frame, text='↓')
        allow_label.grid(row=1, column=2, pady=2, padx=(30, 20), sticky=tk.N)
        self.edited_box = tk.Listbox(label_frame)
        self.edited_box.grid(row=2, column=2, columnspan=2, pady=(5, 20), padx=(20, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scroll_edited = ttk.Scrollbar(label_frame, orient=tk.VERTICAL, command=self.edited_box.yview)
        self.edited_box['yscrollcommand'] = y_scroll_edited.set
        y_scroll_edited.grid(row=2, column=4, pady=(5, 20), padx=(0, 20), sticky=(tk.N, tk.S, tk.W))

        label_frame.columnconfigure(0, weight=1)
        label_frame.columnconfigure(2, weight=1)
        label_frame.rowconfigure(2, weight=1)


    def select_file(self, ext):
        initialdir = os.path.abspath(os.path.dirname(__file__))
        filetypes = [('Data file', f'*.{ext}')]
        target_file = dialog.askopenfilename(
                filetypes=filetypes, initialdir=initialdir)
        self.file_path.set(target_file)

        
    def set_converter(self, converter_class):
        converter = None
        if target_file := self.file_path.get():
            try:
                converter = converter_class(target_file)
            # Whether the selected file exists is checked by tkinter doalog.  
            except ExtensionError as e:
                messagebox.showerror('Error', e)
                self.file_path.set('')
        return converter   

    
    def switch_button_state(self, state):
        self.convert_button['state'] = state
        self.deselect_button['state'] = state


    def to_excel(self, event=None):
        if selected_keys := [self.key_box.get(x) for x in self.key_box.curselection()]:
            self.converter.partial_convert(*selected_keys)
        else:
            self.converter.convert()
        messagebox.showinfo('Info', 'Complete!')
        
  
    def open_json(self, event=None):
        self.switch_button_state(tk.DISABLED)
        self.select_file('json')
        self.key_box.delete(0, tk.END)
        if converter := self.set_converter(ToExcel):
            self.converter = converter
            self.converter.set_sheet_format()
            exclude_keys = set(f'{sh_name}-0' for sh_name \
                in self.converter.sheet_format.values())
            display_keys = sorted(key for key in self.converter.sheet_format.keys() \
                if key not in exclude_keys)
            self.key_box.insert(tk.END, *display_keys)
            self.switch_button_state(tk.NORMAL)

    
    def deselect(self, event=None):
        self.key_box.selection_clear(0, tk.END)

        
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