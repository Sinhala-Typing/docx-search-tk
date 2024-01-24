import os
import logging
from docx import Document
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

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

def __check(fpath, target):
    try:
        doc = Document(fpath)
        for paragraph in doc.paragraphs:
            if target in paragraph.text:
                return True
        return False
    except Exception as e:
        logger.error("Error processing %s: %s" % (fpath, e))
        return False

def __process_file(file):
    fname, target = file
    fpath = os.path.join(os.getcwd(), fname)
    if __check(fpath, target):
        logger.info("'%s' found in %s" % (target, fname))
        return fname
    else:
        logger.debug("'%s' not found in %s" % (target, fname))
        return None

def perform_search(target_dir=None, target_word=None):
    if target_word is None or target_word == '':
        raise ValueError("target_word cannot be None or an empty string. Please pass in a valid value.")

    if target_dir is None:
        target_dir = os.getcwd()

    file_list = [(fname, target_word) for fname in os.listdir(target_dir) if fname.endswith(".docx")]

    results = []
    with ThreadPoolExecutor() as executor:
        results = list(executor.map(__process_file, file_list))

    return [result for result in results if result is not None]

class GUIFrontend:
    def __init__(self, master):
        self.master = master
        self.master.title("Document Search App")

        self.target_word_label = tk.Label(master, text="Target Word:")
        self.target_word_label.pack()

        self.target_word_entry = tk.Entry(master)
        self.target_word_entry.pack()

        self.browse_button = tk.Button(master, text="Browse", command=self.browse_directory)
        self.browse_button.pack()

        self.search_button = tk.Button(master, text="Search", command=self.search_documents)
        self.search_button.pack()

        self.result_listbox = tk.Listbox(master)
        self.result_listbox.pack()
        self.result_listbox.bind("<Double-Button-1>", self.open_file)

        # Initialize target_directory
        self.target_directory = None

    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.target_directory = directory

    def search_documents(self):
        target_word = self.target_word_entry.get()
        if not target_word:
            tk.messagebox.showerror("Error", "Target word cannot be empty.")
            return

        if self.target_directory is None:
            self.target_directory = os.getcwd()

        results = perform_search(target_dir=self.target_directory, target_word=target_word)
        self.update_result_list(results)

    def update_result_list(self, results):
        self.result_listbox.delete(0, tk.END)
        for result in results:
            self.result_listbox.insert(tk.END, result)

    def open_file(self, event):
        selected_item = self.result_listbox.curselection()
        if selected_item:
            file_name = self.result_listbox.get(selected_item)
            file_path = os.path.join(self.target_directory, file_name)
            os.startfile(file_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = GUIFrontend(root)
    root.mainloop()
