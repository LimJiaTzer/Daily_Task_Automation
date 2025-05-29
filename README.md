# 🛠️ Daily Task Automation

A Python script designed to automate your mundane Windows-based daily tasks and workflows with just one click. It seamlessly connects multiple processes like Excel macro execution, file management, application control, and web automation into a streamlined pipeline—boosting productivity and reducing manual effort.  
![image](https://github.com/user-attachments/assets/e56490c2-af48-4170-847d-df8d679c04a3)

---

## 🚀 Key Features

- ✅ **Run Existing Excel Macros**  
  Automate complex tasks already built into your `.xlsm` or `.xlsb` files.
- 📂 **Intelligent File Management**  
  Automatically find, move, copy, or delete files based on specific criteria (name patterns, location). Handles archiving and cleanup.
- 🖥️ **Application & Website Management**  
  Programmatically open, focus, or close desktop applications (like Outlook, SAP Logon) and web pages.
- 🔁 **SAP Reinitialization**  
  Overcome SAP timeout issues and reinitiates the application to maintain smooth operation.
- 🌛 **Time defined tasks**  
  Automatically detect time and run tasks after the time indicated
- 🌍 **Web Automation with Selenium**  
  Automate browser tasks like exporting data from Power BI and integrating it directly into your Excel workflows.
- 🐞 **Robust Debugging Framework**  
  Debug with confidence—comprehensive error handling and logging included.

---

## 🧩 Design Philosophy

- 🎨 **User-Friendly Interface:**  
  Interactive command-line prompts with clear instructions and visual feedback using `colorama` and `cowsay`.
- 🔧 **Extensible Task Engine**  
  Modular architecture makes it easy to add new tasks and customize workflows.
- 📝 **Well-documented Functions**  
  Well-documented functions for easier understanding and usage for new task creation.

---

## 📌 Use Cases

- Daily Excel report generation  
- Automated data downloads and processing  
- Routine file archiving  
- Power BI export → Excel macro automation pipeline
- Export data from RR
- Automatically run end-of-day tasks
- and many more...

---

## 🛠️ Project Setup & Task Creation Guide

### 📦 Download and Setup

Start by downloading or cloning the repository. The following Python files are essential:

- `config.py` – Stores configuration variables and settings.
- `helper.py` – Contains utility functions used across multiple tasks.
- `Main.py` – Main execution script.
- `sample.py` – A reference script showing how to create and execute a new task.

### 🧩 Creating a New Task

To create a new task:

1. **Use `sample.py` or `sample_for_webaccess_and_selenium.py` as a Template:**  
   Refer to `sample.py` for a working example. It demonstrates the general structure and usage of existing basic functions.  
   Refer to `sample_for_webaccess_and_selenium` for a working example for selenium usage in automating web-related tasks.

3. **Leverage Existing Functions:**  
   Most actions you'll need are already implemented in `helper.py`. Simply import and use them as needed.

4. **Add Custom Logic (if necessary):**  
   If your task requires additional functionality not found in `helper.py`:
   - Write your custom functions.
   - Add them to `helper.py` to maintain consistency and reusability across tasks.

> **Tip:** Keeping all helper functions inside `helper.py` helps maintain code structure, improves generality, and makes debugging easier.

### 📁 File Structure Overview
project-root/  
│  
├── config.py                       # Configuration and settings  
├── helper.py                       # Reusable utility functions  
├── Main.py                         # Entry point of the application  
├── sample.py                       # Example script for task creation  
└── sample_for_webaccess_and_selenium.py # Example script for web access and Selenium tasks



