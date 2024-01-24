# docx-search

<p align="center">
  <img src="https://github.com/hirusha-adi/docx-search/assets/36286877/3c2d9933-8460-43ba-b8b3-773940cc3cb4" />
</p>

**Description:**

The `docx-search` Python script is a tool designed to search for a specified word within Microsoft Word (.docx) documents in a given directory. It utilizes the `python-docx` library for handling Word documents and implements multi-threading to improve search efficiency.

**Features:**

- **Word Search:** The script searches for a specified target word within the paragraphs of each Word document in the provided directory.
- **Logging:** Detailed logging is implemented, capturing information about the search process, including the presence or absence of the target word in each document.
- **Multi-threading:** The script utilizes the `concurrent.futures.ThreadPoolExecutor` to concurrently process multiple Word documents, improving overall search performance.

- **Graphical User Interface**

  - Check out the below demonstration:

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

**Contributing:**

- Contributions are welcome! Feel free to fork the repository, make improvements, and create a pull request.

**License:**

- This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

**Acknowledgments:**

- This readme.md and the docstrings were generated with ChatGPT, a language model developed by OpenAI.
