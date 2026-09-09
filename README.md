# Report Management System

Desktop application developed in Python for managing and consulting reports within a workplace environment.

## Description

The project allows users to register materials, manage relationships between requesters and leaders, create reports, and later consult, filter, edit, and export them to Excel.

Current features include:

- Material catalog management.
- Requester and leader management.
- Report creation.
- Mandatory leader assignment for every report.
- Requester assignment for EPP materials.
- Editing existing reports.
- Report search and filtering.
- Exporting results to Excel.
- Data persistence using SQLite.
- Version control using Git.

## Technologies

- Python
- Tkinter / ttk
- SQLAlchemy
- SQLite
- Tkcalendar
- OpenPyXL
- Git / GitHub

## AI-Assisted Development

This project was developed using a **vibe coding** approach, with an iterative collaboration between the developer and an artificial intelligence tool.

AI was used as a development assistant for:

- Generating and modifying code.
- Analyzing errors and unexpected behavior.
- Exploring implementation and design alternatives.
- Explaining programming concepts.
- Refactoring and adapting existing code.
- Rapidly iterating on new features.

The system was built progressively based on real workplace requirements and continuously adjusted through testing and iteration.

The term *vibe coding* is used here to describe the AI-assisted development process. The project was not generated entirely automatically; the developer directed the requirements, evaluated the results, tested the application, and made development decisions throughout the process.

## Project Structure

```text
.
├── main.py          # Graphical interface and main application flow
├── reportes.py      # Business logic and report operations
├── models.py        # SQLAlchemy data models
├── database.py      # Database connection and configuration
├── exportar.py      # Excel export functionality
├── .gitignore       # Files excluded from Git
└── README.md        # Project documentation
