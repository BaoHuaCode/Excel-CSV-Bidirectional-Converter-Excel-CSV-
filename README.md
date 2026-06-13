# Excel-CSV Bidirectional Data Pipeline

A high-performance, memory-efficient Python utility tool designed for seamless, lossless bidirectional 
conversion between Excel files (`.xlsx`) and CSV files. 

This project goes beyond simple format changes by implementing industrial-grade data streaming,
dynamic memory grouping, and an intelligent type-recovery pipeline.

 Key Features

High Performance & Low Memory Footprint**: Utilizes `openpyxl`'s streaming readers (`iter_rows`) 
and writers (`append`) to handle large datasets efficiently without running into memory overflow (OOM)
issues.
Dynamic Memory Grouping**: When reconstructing Excel workbooks from scattered CSV files, the script
uses an in-memory dictionary-grouping algorithm (`setdefault`) to automatically bind sheets back to 
their original parent workbooks.
Lossless Data-Type Recovery**: Solves the common CSV limitation where all data degrades into text 
strings. Built with a Pythonic "EAFP" (`try-except`) detection pipeline to automatically differentiate 
and restore `int`, `float`, and `str` types natively, preventing Excel's annoying "Number Stored as 
Text" green-triangle warnings.
Preserved Workbook Structures**: Supports multi-sheet conversions natively. Splitting an Excel
file preserves sheet names in the CSV filenames, and merging them back perfectly retains the 
multi-sheet architecture.

Project Structure

```text
├── excel_to_csv.py       # Pipeline Step 1: Batch splits Excel workbooks into CSV sheets
├── csv_to_excel.py       # Pipeline Step 2: Batches, groups, cleans, and merges CSVs back into multi-sheet Excel files
└── csv/                  # Central repository for streaming mid-tier CSV data
```

Tech Stack & Key Methods Used

`pathlib.Path`**: Modern object-oriented filesystem paths handling.
`openpyxl`**: Advanced Excel manipulation (`sheet.iter_rows(values_only=True)` & `sheet.append()`).
csv`**: Streamlined native CSV writing and parsing.
Only-Once Flag Pattern**: Implemented stateful switches to elegantly handle `openpyxl`'s default template sheets during runtime generation.

What I Learned From This Project

This project was built during a deep-dive review of Python automation practices. Through it, I mastered:
1. Moving away from heavy `list()` evaluation to **lazy/stream evaluation** for performance optimization.
2. Leveraging **Python Exception Handling (`try-except ValueError`)** as a functional business logic filter rather than just error trapping.
3. Managing data state and collection grouping natively using dictionary memory slots.

