# Student Management System

The **Student Management System** is a Python-based desktop application for managing student admissions, fee payments, and data records. It provides functionality to handle student data efficiently, including generating receipts, managing fee status, and more.

## Features

- **Student Admission Management**:
  - Add new student details.
  - Generate unique admission IDs.

- **Fee Management**:
  - Record fee payments.
  - Generate and save payment receipts.

- **Data Management**:
  - Handle student lists across different batches.
  - Filter and sort student data.

- **Firebase Integration**:
  - Sync data files (Excel) to Firebase Storage.

- **Report Generation**:
  - Generate printable receipts for admissions and fee payments.

## Requirements

### Libraries Used

- `tkinter` (Standard Python library for GUI)
- `pandas` (Data manipulation)
- `openpyxl` (Excel file handling)
- `xlrd` (Read older Excel formats)
- `Pillow` (Image processing)
- `firebase-admin` (Firebase integration)
- `shutil`, `os`, `datetime`, `urllib` (Standard Python libraries)

### Installation

Create a `requirements.txt` file:
```plaintext
pandas
openpyxl
xlrd
Pillow
firebase-admin
