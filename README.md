# 🐍 Tally Excel to XML Converter (Optimized CLI Tool)

## 🚀 Project Overview
This Python-based Command Line Tool is designed to automate the **Bank Statement to Tally Import** process. It converts Excel data into **Tally ERP 9** compatible XML format, allowing for bulk voucher import in seconds.

**The Optimization Story:** Originally built using standard libraries (resulting in a 70MB+ file), this tool was re-engineered using **manual spec pruning and UPX compression**, reducing the final size to just **~24 MB** while maintaining the power of Pandas.

---

## ✨ Features
* **Optimized Performance:** Lightweight (~24 MB) and loads instantly, stripping away unnecessary dependencies.
* **Smart Validation:** Automatically validates Data, Date formats, and Ledgers before generating XML to prevent Tally import errors.
* **XML Safety:** Handles special characters safely to ensure a crash-free import process.
* **Bulk Processing:** Converts hundreds of entries from Excel to XML in a single click.

---

## 🛠️ How to Use

**⚠️ IMPORTANT:** This tool works with a specific Excel Template.

1.  **Download the Project ZIP:** [Download ZIP File](https://drive.google.com/file/d/1ZzvTSKlvTN9Rorp1lS4rP9usfp1et6yp/view?usp=sharing)
    *(Contains the .EXE file and the Excel Template)*
2.  **Open `Tally_Import_Template.xlsx`:**
    * Go to the **"Instructions"** sheet. **READ IT CAREFULLY.**
    * Paste your data into the **"Data"** sheet as per the format.
    * **Close the Excel file.**
3.  **Run the `.exe` file.**
4.  Paste the **Excel File Path** and **Output Folder Path** when prompted.
5.  Import the generated `Tally_Import.xml` in Tally via **Gateway of Tally > Import > Vouchers**.

---

## 💻 Tech Stack & Dependencies
This project is built using:
* **Python 3.13**
* **`pandas` & `openpyxl`**: For robust data handling and Excel reading.
* **`PyInstaller` & `UPX`**: Used for advanced compilation and binary compression.

---

## 🧑‍💻 Connect with the Developer

I am a **CA Finalist** with a focus on leveraging **Python and Automation** to drive efficiency in finance, tax, and audit domains.

| Platform | Link |
| :--- | :--- |
| **LinkedIn** | [Connect with Jaidev Duggal](https://www.linkedin.com/in/jaidev-duggal) |
| **GitHub** | [@JaidevDuggal](https://github.com/JaidevDuggal) |
| **Email** | jaidevduggal249@gmail.com |

**Looking to collaborate on Finance Tech & Automation projects!**

---

## ⚖️ License
This project is licensed under the **MIT License**.
