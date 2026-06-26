# 🏭 NMDC HR Analytics Dashboard

<p align="center">
  <img src="NMDC LOGO.jpg" alt="NMDC Logo" width="180"/>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Status-Active-brightgreen" />
  <img src="https://img.shields.io/badge/Frontend-HTML%20%7C%20CSS%20%7C%20JS-blue" />
  <img src="https://img.shields.io/badge/Backend-Python-yellow" />
  <img src="https://img.shields.io/badge/License-MIT-lightgrey" />
</p>

---

## 📌 Overview

The **NMDC HR Analytics Dashboard** is a full-stack web application built for **National Mineral Development Corporation (NMDC)**. It provides a streamlined interface for managing employee records, grade-wise department summaries, and secure role-based login.

---

## ✨ Features

- 🔐 **Secure Login** — Role-based authentication via `nmdc_login.html`
- 📊 **Employee Grade & Department Summary** — Excel-based reporting (`Employee_Grade_Dept_Summary.xlsx`)
- 🗂️ **Frontend Dashboard** — Intuitive UI built with HTML, CSS, and JavaScript
- ⚙️ **Python Backend** — REST API / server-side logic
- 🚀 **One-click Startup** — Launch entire stack with `start-all.bat`

---

## 🗂️ Project Structure

```
NMDC-PROJECT/
├── frontend/               # HTML, CSS, JS files for UI
├── backend/                # Python backend / API
├── nmdc_login.html         # Login page
├── start-all.bat           # Script to start frontend + backend together
├── Employee_Grade_Dept_Summary.xlsx   # Employee summary report
├── test data 01.05.26.xlsx            # Sample test data
└── NMDC LOGO.jpg           # Official NMDC logo
```

---

## 🚀 Getting Started

### Prerequisites

- Python 3.8+
- A modern web browser (Chrome, Edge, Firefox)

### Installation

```bash
# 1. Clone the repository
git clone https://github.com/sujalsahu5082/NMDC-PROJECT.git
cd NMDC-PROJECT

# 2. Install Python dependencies
pip install -r requirements.txt

# 3. Start the application
start-all.bat        # Windows
# or manually start backend and open frontend in browser
```

---

## 🖥️ Usage

1. Run `start-all.bat` to launch both frontend and backend.
2. Open your browser and go to `http://localhost:PORT` (or open `nmdc_login.html`).
3. Log in with your credentials.
4. Access employee records, grade summaries, and department data.

---

## 🛠️ Tech Stack

| Layer     | Technology          |
|-----------|---------------------|
| Frontend  | HTML5, CSS3, JavaScript |
| Backend   | Python              |
| Data      | Microsoft Excel (.xlsx) |
| Launcher  | Windows Batch Script |

---

## 📸 Screenshots

> *(Add screenshots of the login page and dashboard here)*

---

## 🤝 Contributing

Contributions are welcome! Please follow these steps:

1. Fork this repository
2. Create a new branch: `git checkout -b feature/your-feature`
3. Commit your changes: `git commit -m "Add your feature"`
4. Push to the branch: `git push origin feature/your-feature`
5. Open a Pull Request

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).

---

## 👤 Author

**Sujal Sahu**  
GitHub: [@sujalsahu5082](https://github.com/sujalsahu5082)

---

<p align="center">Made with ❤️ for NMDC</p>
