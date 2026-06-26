# 📅 Automated Timetable Generation System

An Excel VBA-based automated timetable generation system that automatically creates and validates **Class, Faculty, Lab, and Room** timetables. The project uses VBA macros to generate conflict-free schedules while providing real-time error detection to ensure timetable accuracy.

---

## 📂 Project Structure
  Project Folder/
  │
  ├── Master_D.xlsm # Main Excel workbook containing all data, worksheets, and VBA macros
  └── README.md # Project documentation

> **Note:** All worksheets, master data, and VBA code are contained within **`Master_D.xlsm`**. The `README.md` file is located outside the workbook in the project folder.

---

## 🚀 Features

### Automated Timetable Generation
- Generates Class, Faculty, Lab, and Room timetables using Excel VBA macros.
- Automatically updates all timetable worksheets from the master data.

### Dynamic Scheduling Engine
- Implements a **20-row Arithmetic Progression (AP) scanning algorithm** for timetable generation.
- Ensures accurate scheduling while preventing duplicate or conflicting allocations.
- Maintains synchronization across all timetable worksheets.

### Real-Time Error Detection
- Displays popup alerts for:
  - Missing class allocations
  - Header mismatches
  - Faculty scheduling conflicts
  - Lab scheduling conflicts
  - Room scheduling conflicts

---

## 🛠️ Technologies Used

- Microsoft Excel
- VBA (Visual Basic for Applications)
- Excel Macros

---

## ⚙️ How to Use

1. Open **`Master_D.xlsm`**.
2. Enable macros when prompted.
3. Update the required master data.
4. Run the timetable generation macro.
5. Review the generated Class, Faculty, Lab, and Room timetables.
6. Resolve any popup error messages if conflicts are detected.

---

## 📌 Key Highlights

- Single workbook solution
- Fully automated timetable generation
- Conflict-free scheduling
- Dynamic validation using a 20-row AP scanning rule
- Real-time error detection
- Easy to maintain and extend

---

## 📈 Future Enhancements

- Interactive dashboard
- Faculty workload optimization
- PDF export functionality
- Multi-department support
- One-click timetable generation

---

## 👨‍💻 Author

Developed as an Excel VBA automation project to simplify academic timetable management by automating schedule generation, validating data, and detecting conflicts in real time.
