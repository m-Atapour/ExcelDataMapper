DataMapper: Data Cleansing and Mapping Tool
DataMapper is a specialized desktop application built using Python (Pandas and Tkinter) designed to streamline the process of identifying, suggesting fixes for, and mapping inconsistent or missing data points within large spreadsheet files (CSV/Excel).

It focuses primarily on standardizing the values within a designated "Fix Column" by leveraging pattern recognition and external standardization maps.

Key Features
Error Identification:

Finds rows where the value in the "Fix Column" is empty (null).

Finds rows where the value in the "Fix Column" is one of the user-defined invalid unique values.

Intelligent Suggestion Engine:

External Consolidation Map: Users can load a reference file (the mapping) containing standard target values and a list of common input variants (aliases).

Internal Token Search: When an error is found, the system searches designated "Search Columns" in the same row for any word tokens that match standardized values existing in the Fix Column or within the loaded Map.

Prioritized Suggestions: Suggestions are ranked based on the source (e.g., direct map matches are prioritized) and relevance, offering the user the best possible fix immediately.

Interactive Review and Correction:

Results are displayed in a clean, filterable table (Treeview).

Users can review original data and quickly select suggestions via a Context Menu (Right-Click) or manually edit the "Final Selection" column.

Data Persistence and Output:

The application processes the selected rows, applying the corrections from the "Final Selection" column back to the original DataFrame's "Fix Column."

The corrected file is then saved to a new CSV or Excel file, ensuring data standardization for downstream analysis.

Technical Stack
Core: Python

Data Handling: Pandas (for efficient file loading, processing, and column manipulation)

GUI: Tkinter / ttk (for a responsive, native desktop interface)

Standardization: Regular Expressions (re) and comprehensive Persian/Arabic character normalization.
