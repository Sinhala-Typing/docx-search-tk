import os
import logging
import json
from docx import Document
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import tkinter as tk
from tkinter import ttk

# Configure logger
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

def check(fpath, target):
    try:
        doc = Document(fpath)
        for paragraph in doc.paragraphs:
            if target in paragraph.text:
                return True
        return False
    except Exception as e:
        logger.error("Error processing %s: %s" % (fpath, e))
        return False

def process_file(file):
    fname, target = file
    fpath = os.path.join(os.getcwd(), fname)
    if check(fpath, target):
        logger.info("'%s' found in %s" % (target, fname))
        return fpath
    else:
        logger.debug("'%s' not found in %s" % (target, fname))
        return None

def load_config_json(file_list, target_word):
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

def main(target_dir=None, target_word=None):
    if target_word is None or target_word == '':
        raise ValueError("target_word cannot be None or an empty string. Please pass in a valid value.")

    if target_dir is None:
        target_dir = os.getcwd()

    file_list = []
    load_config_json(file_list, target_word)
    logger.debug(f"Discovered {len(file_list)} files.")
    
    found_files = []
    with ThreadPoolExecutor() as executor:
        for result in executor.map(process_file, file_list):
            if result is not None:
                found_files.append(result)

    return found_files

def docx_search(target_dir=None, target_word=None):
    return main(target_dir=target_dir, target_word=target_word)


class DocxSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Docx Search App")

        self.target_word_label = ttk.Label(root, text="Enter Target Word:")
        self.target_word_label.grid(row=0, column=0, padx=10, pady=10)

        self.target_word_entry = ttk.Entry(root)
        self.target_word_entry.grid(row=0, column=1, padx=10, pady=10)

        self.search_button = ttk.Button(root, text="Search", command=self.search)
        self.search_button.grid(row=1, column=0, columnspan=2, pady=10)

        self.result_label = ttk.Label(root, text="")
        self.result_label.grid(row=2, column=0, columnspan=2, pady=10)

        self.found_files_listbox = tk.Listbox(root)
        self.found_files_listbox.grid(row=3, column=0, columnspan=2, pady=10)
    
    def search(self):
        target_word = self.target_word_entry.get()
        found_files = docx_search(target_word=target_word)
        self.result_label.config(text=f"Found '{target_word}' in {len(found_files)} files")

        # Clear previous entries
        self.found_files_listbox.delete(0, tk.END)
        # Add found files to the listbox
        for file_path in found_files:
            self.found_files_listbox.insert(tk.END, file_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = DocxSearchApp(root)
    root.mainloop()
