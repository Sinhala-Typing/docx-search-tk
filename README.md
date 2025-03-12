# .docx Search

<p align="center">
  <img src="https://github.com/user-attachments/assets/ac9d8010-4ec6-4928-b7e0-d1782a134726" />
</p>

The `docx-search` Python script is a tool designed to search for a specified word within Microsoft Word (.docx) documents in a given directory. It utilizes the `python-docx` library for handling Word documents and implements multi-threading to improve search efficiency.

<p align="center">
  <img src="https://github.com/Sinhala-Typing/docx-search-tk/assets/36286877/049dca93-ce24-4eb4-998c-cd34cbeeb94f" />
</p>


**Features:**

- **Word Search:** The script searches for a specified target word within the paragraphs of each Word document in the provided directory.
- **Logging:** Detailed logging is implemented, capturing information about the search process, including the presence or absence of the target word in each document.
- **Multi-threading:** The script utilizes the `concurrent.futures.ThreadPoolExecutor` to concurrently process multiple Word documents, improving overall search performance.

- **Graphical User Interface**

  - Check out the below demonstration:

https://github.com/Sinhala-Typing/docx-search-tk/assets/36286877/a93352a4-28d3-4f8c-8f5c-39e9c883dd91



**Getting Started:**

1. **Requirements:**

   - Python 3.x
   - Install required Python packages using

   ```
   pip install -r requirements.txt
   ```

2. **Usage:**

   - Run the script from the command line:
     ```
     python search.py
     ```

3. **Logging:**

   - Logs are saved in the 'logs' directory with filenames in the format 'YYYY-MM-DD_HH-MM-SS.log.'

4. **Output:**

   - The script outputs information about the presence or absence of the target word in each processed document.
