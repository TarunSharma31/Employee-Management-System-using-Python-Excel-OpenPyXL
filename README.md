# Employee Management System using Python & Excel

A complete **Employee Management System** developed using Python and Excel (`openpyxl`) for storing and managing employee records.
This project performs employee record management, salary calculation, temporary/permanent deletion, recovery of deleted data, and PDF salary slip generation.

---

# 📌 Project Description

The **Employee Management System** is a console-based Python application that stores employee records in Excel files and allows users to perform various operations such as:

* Insert employee records
* View employee data
* Search employee records
* Update records
* Delete records
* Recover deleted records
* Calculate employee salary
* Generate salary reports
* Create PDF employee salary cards

The project uses Excel files as a lightweight database system.

---

# 🎯 Why This Project Was Made

This project was created for:

* Learning Python file handling
* Understanding Excel automation using Python
* Practicing CRUD operations
* Learning employee management systems
* Working with real-world business records
* Understanding salary calculations
* Learning PDF generation using Python

This project is useful for:

* BCA Students
* MCA Students
* Python beginners
* Office management systems
* Academic projects

---

# 🚀 Features

## ✅ Employee Record Management

* Add new employee records
* Auto-generate employee IDs
* Display all employee records
* Update employee information
* Search employee details

---

## ✅ Search Features

* Search by Employee ID
* Search by Department
* Search by City

---

## ✅ Delete & Recovery System

* Temporary delete
* Permanent delete
* Recover deleted employee records
* Recover all deleted records

---

## ✅ Salary Management

* Salary calculation
* HRA calculation
* DA calculation
* Income Tax deduction
* Net salary generation

---

## ✅ Report Generation

* Employee salary card
* PDF salary report generation

---

# 🛠️ Technologies Used

* Python
* OpenPyXL
* Excel Workbook (.xlsx)
* ReportLab
* Console-based Interface

---

# 📚 Libraries Used

```python
import openpyxl as op
import os
import sys
import msvcrt as msv
```

Additional Library for PDF Generation:

```python
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
```

---

# 📂 Project Files

```text
employeeOPXL.py
myemployee.xlsx
```

---

# 📖 Excel Sheets Used

| Sheet Name              | Purpose                            |
| ----------------------- | ---------------------------------- |
| MainFile / Active Sheet | Main employee records              |
| deletedEmp              | Stores temporarily deleted records |
| EmpSalary               | Stores salary calculation data     |

---

# 📋 Employee Data Stored

Each employee record contains:

```text
ID
NAME
DEPARTMENT
D.O.B.
CITY
PHONE NO.
SALARY
```

---

# ⚙️ Operations Performed

## 1. Insert Employee Records

* Automatically generates employee ID
* Stores employee details into Excel sheet

---

## 2. Show Employee Records

* Displays all employee data in tabular format

---

## 3. Search Employee Data

Search employee records using:

* Employee ID
* Department
* City

---

## 4. Update Employee Records

Update:

* Name
* Department
* Date of Joining
* City
* Phone Number
* Salary
* Entire record

---

## 5. Delete Employee Records

### Temporary Delete

* Moves records to `deletedEmp` sheet

### Permanent Delete

* Deletes records permanently

---

## 6. Recover Deleted Records

* Recover single employee
* Recover all deleted employees

---

## 7. Salary Calculation

The system calculates:

| Salary Component | Formula                |
| ---------------- | ---------------------- |
| HRA              | 20% of Basic Salary    |
| DA               | 60% of Basic Salary    |
| Income Tax       | 15% deduction          |
| Net Salary       | Basic + HRA + DA - Tax |

---

## 8. Generate Employee Salary Card

Displays:

* Employee information
* Salary details
* Net salary

---

## 9. PDF Generation

Generate professional employee salary PDF reports using ReportLab.

---

# ▶️ How to Run the Project

## Step 1: Install Python

Download Python:

[https://www.python.org/downloads/](https://www.python.org/downloads/)

---

## Step 2: Install Required Libraries

Open terminal or command prompt:

```bash
pip install openpyxl reportlab
```

---

## Step 3: Run the Program

```bash
python employeeOPXL.py
```

---

# 🖥️ Main Menu

```text
1 -> Insert records in the file
2 -> Show the records in the file
3 -> Search data in the file
4 -> Update the records of the file
5 -> Delete any record in the file
6 -> Recover deleted records
7 -> Salary calculator of employees
8 -> Get full record of any employee
```

---

# 🔍 Sample Record

```text
ID        : 1001
Name      : Tarun
Department: IT
City      : Jaipur
Salary    : 50000
```

---

# 📄 Sample Salary Calculation

```text
Basic Salary : ₹50000
HRA          : ₹10000
DA           : ₹30000
Income Tax   : ₹13500

Net Salary   : ₹76500
```

---

# 📈 Learning Concepts Covered

This project covers:

* Python File Handling
* Excel Automation
* CRUD Operations
* Workbook & Worksheet Handling
* Search Algorithms
* Update/Delete Operations
* Salary Calculation Logic
* PDF Generation
* Console-based Application Design

---

# 🔮 Future Improvements

Possible future enhancements:

* GUI version using Tkinter
* Database integration using MySQL
* Login authentication system
* Attendance management
* Employee dashboard
* Export reports to PDF automatically
* Email salary slips
* Web-based employee portal

---

# 👨‍💻 Author

Developed for educational and learning purposes using Python and Excel automation concepts.

---

# 📜 License

This project is free to use for educational and academic purposes.
