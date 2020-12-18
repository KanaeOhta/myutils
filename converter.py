from collections import namedtuple
import errno
import os
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import tkinter.filedialog as dialog

from jsonexcel import ToExcel, FromExcel, file_check, ExtensionError


TOEXCEL = 'ToExcel'
FROMEXCEL = 'FromExcel'


Key = namedtuple('Key', 'nested real')


class ConverterWindow(ttk.Frame):
    
    def __init__(self, master):
        super().__init__(master)
        self.create_variables()
        self.create_ui()


    def create_variables(self):
        self.indent_list = [0, 2, 4]
        self.json_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.edit_key = tk.StringVar()
        self.indent = tk.StringVar()
        self.now_selected = None
       
       
    def create_ui(self):
        self.create_base_frame()
        self.create_tab_toexcel()
        self.create_tab_fromexcel()
        

    def create_base_frame(self):
        frame = ttk.Frame(self.master)
        frame.pack(fill=tk.BOTH, expand=True)
        self.note = ttk.Notebook(frame)
        self.note.pack(fill=tk.BOTH, expand=True, pady=20, padx=20)
        self.tab_toexcel = ttk.Frame(self.note)
        self.tab_fromexcel = ttk.Frame(self.note)
        self.note.add(self.tab_toexcel, text=TOEXCEL)
        self.note.add(self.tab_fromexcel, text=FROMEXCEL)
        close_button = ttk.Button(frame, text='Close', command=self.close)
        close_button.pack(side=tk.RIGHT, pady=10, padx=20)


    def create_tab_top_frame(self, main_frame, textvariable):
        """Create a frame having an entry for a file path.
        """
        top_frame = ttk.Frame(main_frame)
        filename_label = ttk.Label(top_frame, text='File name ')
        filename_label.grid(row=0, column=0, sticky=tk.E)
        path_entry = ttk.Entry(top_frame, textvariable=textvariable, width=70, state='readonly')
        path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        open_button = ttk.Button(top_frame, text='Open', command=self.open)
        open_button.grid(row=0, column=2, sticky=tk.E)
        top_frame.columnconfigure(1, weight=1)
        return top_frame


    def create_tab_toexcel(self):
        main_frame = ttk.Frame(self.tab_toexcel)
        main_frame.pack(fill=tk.BOTH, expand=True, pady=20, padx=20)
        top_frame = self.create_tab_top_frame(main_frame, self.json_path)
        top_frame.grid(row=0, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S))
  
        label_frame = ttk.LabelFrame(main_frame, text=' Select keys, if you want export selected data. ')
        label_frame.grid(row=1, column=0, columnspan=4, pady=30, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.key_box = tk.Listbox(label_frame, selectmode='multiple')
        self.key_box.grid(row=0, column=0, pady=20, padx=(20, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scroll = ttk.Scrollbar(label_frame, orient=tk.VERTICAL, command=self.key_box.yview)
        self.key_box['yscrollcommand'] = y_scroll.set
        y_scroll.grid(row=0, column=1, pady=20, padx=(0, 20), sticky=(tk.N, tk.S, tk.W))

        self.toexcel_convert_button = ttk.Button(main_frame, text='Convert',
            command=self.convert, state=tk.DISABLED)
        self.toexcel_convert_button.grid(row=2, column=2)
        self.toexcel_deselect_button = ttk.Button(main_frame, text='Deselect', 
            command=self.deselect, state=tk.DISABLED)
        self.toexcel_deselect_button.grid(row=2, column=3)
        
        label_frame.columnconfigure(0, weight=1)
        label_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)


    def create_tab_fromexcel(self):
        main_frame = ttk.Frame(self.tab_fromexcel)
        main_frame.pack(fill=tk.BOTH, expand=True, pady=20, padx=20)
        top_frame = self.create_tab_top_frame(main_frame, self.excel_path)
        top_frame.grid(row=0, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.fromexcel_convert_button = ttk.Button(main_frame, text='Convert', 
            command=self.convert, state=tk.DISABLED)
        self.fromexcel_convert_button.grid(row=2, column=2)
        self.fromexcel_deselect_button = ttk.Button(main_frame, text='Deselect', 
            command=self.deselect, state=tk.DISABLED)
        self.fromexcel_deselect_button.grid(row=2, column=3)

        label_frame  = ttk.LabelFrame(main_frame, text=' Edit keys, if you need. ')
        label_frame.grid(row=1, column=0, columnspan=4, pady=30, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.keys_box = tk.Listbox(label_frame)
        self.keys_box.grid(row=0, rowspan=3, column=0, pady=20, padx=(20, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scroll_keys = ttk.Scrollbar(label_frame, orient=tk.VERTICAL, command=self.keys_box.yview)
        self.keys_box['yscrollcommand'] = y_scroll_keys.set
        y_scroll_keys.grid(row=0, rowspan=3, column=1, pady=20, padx=(0, 10), sticky=(tk.N, tk.S, tk.E))
        self.keys_box.bind('<<ListboxSelect>>', self.keys_box_click)

        self.edit_entry = ttk.Entry(label_frame, textvariable=self.edit_key, width=20)
        self.edit_entry.grid(row=0, column=2, pady=(20, 0), padx=(20, 0), sticky=(tk.W, tk.N, tk.E))
        ok_button = ttk.Button(label_frame, text='OK', width=4, command=self.edit)
        ok_button.grid(row=0, column=3, pady=(19, 0), sticky=(tk.N, tk.E))
        allow_label = ttk.Label(label_frame, text='↓')
        allow_label.grid(row=1, column=2, pady=2, padx=(30, 20), sticky=tk.N)
        self.edited_box = tk.Listbox(label_frame, selectmode='multiple')
        self.edited_box.grid(row=2, column=2, columnspan=2, pady=(5, 20), padx=(20, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scroll_edited = ttk.Scrollbar(label_frame, orient=tk.VERTICAL, command=self.edited_box.yview)
        self.edited_box['yscrollcommand'] = y_scroll_edited.set
        y_scroll_edited.grid(row=2, column=4, pady=(5, 20), padx=(0, 20), sticky=(tk.N, tk.S, tk.W))

        label_frame.columnconfigure(0, weight=1)
        label_frame.columnconfigure(2, weight=1)
        label_frame.rowconfigure(2, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)


    def open(self, event=None):
        tab_name = self.get_current_tab_name()
        if tab_name == TOEXCEL:
            self.open_json()
        else:
            self.open_excel()

    
    def convert(self, event=None):
        """Called when convert button is clicked.
        """
        tab_name = self.get_current_tab_name() 
        if tab_name == TOEXCEL:
            self.to_excel()
        else:
            self.from_excel()
        

    def deselect(self, event=None):
        """Called when deselect button is clicked.
        """
        tab_name = self.get_current_tab_name() 
        if tab_name == TOEXCEL:
            self.key_box.selection_clear(0, tk.END)
        else:
            self.edited_box_deselect()
            

    def close(self, event=None):
        self.quit()


    def keys_box_click(self, event=None):
        """Called when item is selected in keys_box(ListBox).
        """
        if indexes := self.keys_box.curselection():
            item = self.keys_box.get(indexes[0])
            self.now_selected = item.split('  ')[0]
            key = self.key_table[self.now_selected]
            self.edit_key.set(key.real)
    
    
    def edit(self, event=None):
        """Called when ok button on FromExcel tab is clicked.
        """
        if edited := self.edit_key.get():
            if self.now_selected in self.edited_keys:
                messagebox.showerror('Error', f'{self.now_selected} was already edited.')
            else:
                original_key = self.key_table[self.now_selected] 
                if original_key.real != edited:
                    self.edited_keys[self.now_selected] = Key(original_key.nested, edited)
                    self.edited_box.insert(tk.END, f'{self.now_selected}  {edited}')
        self.edit_entry.delete(0, tk.END)
        self.keys_box.select_clear(0, tk.END)


    def select_file(self, ext, string_var):
        initialdir = os.path.abspath(os.path.dirname(__file__))
        filetypes = [('Data file', f'*.{ext}')]
        target_file = dialog.askopenfilename(
                filetypes=filetypes, initialdir=initialdir)
        string_var.set(target_file)

        
    def set_converter(self, string_var, class_):
        converter = None
        if target_file := string_var.get():
            try:
                converter = class_(target_file)
            # Whether the selected file exists is checked by tkinter doalog.  
            except ExtensionError as e:
                messagebox.showerror('Error', e)
                string_var.set('')
            except IOError as e:
                if e.errno == errno.EACCES:
                    messagebox.showerror('Error', 'Please close the excel file.')
                    string_var.set('')
        return converter  
       

    def get_current_tab_name(self):
        """Return the tab name which is selected now.
        """
        current_tub = self.note.select()
        tab_name = self.note.tab(current_tub)['text']
        return tab_name


    def switch_button_state(self, state):
        tab_name = self.get_current_tab_name()
        if tab_name == TOEXCEL:
            self.toexcel_convert_button['state'] = state
            self.toexcel_deselect_button['state'] = state
        else:
            self.fromexcel_convert_button['state'] = state
            self.fromexcel_deselect_button['state'] = state


    def to_excel(self):
        """Called when convert button on toexcel tab is clicked.
        """
        if selected_keys := [self.key_box.get(x) for x in self.key_box.curselection()]:
            self.converter.partial_convert(*selected_keys)
        else:
            self.converter.convert()
        messagebox.showinfo('Info', 'Complete!')
        

    def from_excel(self):
        """Called when convert button on fromexcel tab is clicked.
        """
        replacement = {key.nested: key.real for key in self.edited_keys.values()} \
            if self.edited_keys else None
        self.converter.convert(replacement=replacement)
        messagebox.showinfo('Info', 'Complete!')


    def open_json(self):
        """Called when open button on toexcel tab is clicked.
        """
        self.switch_button_state(tk.DISABLED)
        self.select_file('json', self.json_path)
        self.key_box.delete(0, tk.END)
        if converter := self.set_converter(self.json_path, ToExcel):
            self.converter = converter
            self.converter.set_sheet_format()
            exclude_keys = set(f'{sh_name}-0' for sh_name \
                in self.converter.sheet_format.values())
            display_keys = sorted(key for key in self.converter.sheet_format.keys() \
                if key not in exclude_keys)
            self.key_box.insert(tk.END, *display_keys)
            self.switch_button_state(tk.NORMAL)


    def open_excel(self):
        """Called when open button on fromexcel tab is clicked.
        """
        self.select_file('xlsx', self.excel_path)
        self.keys_box.delete(0, tk.END)
        self.edited_box.delete(0, tk.END)
        self.edited_keys = {}
        if converter := self.set_converter(self.excel_path, FromExcel):
            self.converter = converter
            converter.set_sheets()
            key_table = set()
            for sh in converter.sheets:
                key_table.update({Key(nested_key, real_key) for sh_keys 
                    in sh.keys for nested_key, real_key in converter.separate(sh_keys)})
            len_table = len(key_table)
            zeros = len(str(len_table))
            self.key_table = {str(i).zfill(zeros): key for i, key in enumerate(sorted(key_table), 1)}
            display_keys = [f'{idx}  {key.nested}' for idx, key in self.key_table.items()]
            self.keys_box.insert(tk.END, *display_keys)
            self.switch_button_state(tk.NORMAL)


    def edited_box_deselect(self):
        """Called when deselect button on fromexcel tab is clicked.
        """
        delete_items = []
        for index in sorted(self.edited_box.curselection(), reverse=True):
            selected = self.edited_box.get(index)
            # index number is append to delete_item list
            delete_items.append(selected.split('  ')[0])
            self.edited_box.delete(index)
        self.edited_keys = {k: v for k, v in self.edited_keys.items() \
            if k not in delete_items}
        print(self.edited_keys)

        
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