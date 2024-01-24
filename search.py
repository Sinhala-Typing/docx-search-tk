import os
import json
import tkinter as tk
from tkinter import ttk
from docx import Document
from concurrent.futures import ThreadPoolExecutor
import logging
from datetime import datetime
from tkinter import font

if not os.path.isdir('logs'):
    os.mkdir('logs')

log_format = '(%(asctime)s) [%(levelname)s] %(message)s'
log_file_name = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + '.log'
log_file_path = os.path.join(os.getcwd(), "logs", log_file_name)

logging.basicConfig(level=logging.DEBUG, format=log_format, handlers=[
    logging.FileHandler(log_file_path, encoding='utf-8'),
    logging.StreamHandler()
])

logger = logging.getLogger(__name__)

class DocxSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Docx Search GUI")
        self.root.resizable(width=False, height=False)
        
        title_font = font.Font(family="Helvetica", size=16, weight="bold")
        
        self.title_label = ttk.Label(root, text="Docx-Search GUI", font=title_font)
        self.title_label.grid(row=2, column=0, padx=3, pady=3, columnspan=4)
        
        self.target_word_label = ttk.Label(root, text="Query:")
        self.target_word_label.grid(row=3, column=0, padx=4, pady=4, columnspan=1)

        self.target_word_entry = ttk.Entry(root, width=45)  # Increase the width of the entry widget
        self.target_word_entry.grid(row=3, column=1, padx=4, pady=4, columnspan=2)
        
        self.clear_button = ttk.Button(root, text="X", command=self.clear_entries, width=5)
        self.clear_button.grid(row=3, column=3, columnspan=1, pady=4, padx=4)

        self.search_button = ttk.Button(root, text="Search", command=self.search, width=62)
        self.search_button.grid(row=4, column=0, columnspan=4, pady=4, padx=4)

        self.result_label = ttk.Label(root, text="Made by @hirushaadi")
        self.result_label.grid(row=5, column=0, columnspan=4, pady=4)

        # Add scrollbars to the Listbox
        self.found_files_listbox = tk.Listbox(root, selectmode=tk.SINGLE, exportselection=0, height=10, width=55)
        self.found_files_listbox.grid(row=6, column=0, columnspan=3, pady=4, padx=4, sticky="nsew")

        self.scrollbar_y = ttk.Scrollbar(root, orient="vertical", command=self.found_files_listbox.yview)
        self.scrollbar_y.grid(row=6, column=3, sticky="ns", columnspan=1, rowspan=1)
        self.found_files_listbox.configure(yscrollcommand=self.scrollbar_y.set)

        self.scrollbar_x = ttk.Scrollbar(root, orient="horizontal", command=self.found_files_listbox.xview)
        self.scrollbar_x.grid(row=7, column=0, columnspan=3, sticky="ew")
        self.found_files_listbox.configure(xscrollcommand=self.scrollbar_x.set)

        self.found_files_listbox.bind('<Double-Button-1>', self.open_selected_file)

    def clear_entries(self):
        self.target_word_entry.delete(0, tk.END)
        self.found_files_listbox.delete(0, tk.END)
    
    def check(self, fpath, target):
        try:
            doc = Document(fpath)
            for paragraph in doc.paragraphs:
                if target in paragraph.text:
                    return True
            return False
        except Exception as e:
            logger.error("Error processing %s: %s" % (fpath, e))
            return False

    def process_file(self, file):
        fname, target = file
        fpath = os.path.join(os.getcwd(), fname)
        if self.check(fpath, target):
            logger.info("'%s' found in %s" % (target, fname))
            return fpath
        else:
            logger.debug("'%s' not found in %s" % (target, fname))
            return None

    def load_config_json(self, file_list, target_word):
        config_file_path = os.path.join(os.getcwd(), 'config.json')

        if os.path.exists(config_file_path):
            logger.debug(f"Found config file at: {config_file_path}")
            with open(config_file_path, 'r') as config_file:
                config_data = json.load(config_file)

                if 'dirs' in config_data and isinstance(config_data['dirs'], list):
                    logger.debug(f"Found {len(config_data['dirs'])} directories in 'dirs'")
                    for directory in config_data['dirs']:
                        logger.debug(f"Traversing through: '{directory}'")
                        directory_path = os.path.abspath(directory)

                        for entry in os.scandir(directory_path):
                            if entry.is_file() and entry.name.endswith(".docx"):
                                file_list.append((entry.path, target_word))
                            elif entry.is_dir():
                                for root, _, files in os.walk(entry.path):
                                    for fname in files:
                                        if fname.endswith(".docx"):
                                            file_list.append((os.path.join(root, fname), target_word))
                else:
                    logger.debug("Error in 'dirs' key of config file")

    def main(self, target_dir=None, target_word=None):
        if target_word is None or target_word == '':
            raise ValueError("target_word cannot be None or an empty string. Please pass in a valid value.")

        if target_dir is None:
            target_dir = os.getcwd()

        file_list = []
        self.load_config_json(file_list, target_word)
        logger.debug(f"Discovered {len(file_list)} files.")

        found_files = []
        with ThreadPoolExecutor() as executor:
            for result in executor.map(self.process_file, file_list):
                if result is not None:
                    found_files.append(result)

        return found_files

    def docx_search(self, target_dir=None, target_word=None):
        return self.main(target_dir=target_dir, target_word=target_word)

    def search(self):
        target_word = self.target_word_entry.get()
        found_files = self.docx_search(target_word=target_word)
        self.result_label.config(text=f"Found '{target_word}' in {len(found_files)} files")

        self.found_files_listbox.delete(0, tk.END)
        for file_path in found_files:
            self.found_files_listbox.insert(tk.END, file_path)

    def open_selected_file(self, event):
        selected_index = self.found_files_listbox.curselection()
        if selected_index:
            selected_file = self.found_files_listbox.get(selected_index)
            os.startfile(selected_file)

if __name__ == "__main__":
    root = tk.Tk()
    app = DocxSearchApp(root)
    root.mainloop()
