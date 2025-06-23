# 📊 Teams Policy Matrix GUI by Edgar Avellan

A PowerShell-based graphical tool that connects to Microsoft Teams, retrieves all available policy types, and enables side-by-side **viewing** and **drift comparison** of policies in an interactive Windows Forms interface.

---

## 🧠 What is This?

`TeamsPolicyMatrixGui.ps1` is a self-contained PowerShell script designed to:

* Connect to Microsoft Teams using `Connect-MicrosoftTeams`
* Pull and display over **55+ policy types**
* Allow **multi-selection** for viewing
* Show side-by-side policy comparisons
* Perform **drift detection** by highlighting differences between two selected policies
* Output logs, loading status, and summaries in real-time

---

## 🚀 Features

| Feature            | Description                                                   |
| ------------------ | ------------------------------------------------------------- |
| 🔌 Connect Button  | Authenticates with Microsoft Teams tenant                     |
| 📥 Load Button     | Retrieves and lists all policies with identity tagging        |
| 🪟 View Button     | Opens a side-by-side viewer for up to 2 selected policies     |
| 🆚 Compare Button  | Highlights property differences between 2 selected policies   |
| 🖥️ GUI Interface  | Easy-to-use Windows Forms interface                           |
| 📋 Summary Logging | Console-style feedback for successful vs. failed policy loads |

---

## 🛠️ Requirements

* Windows with PowerShell 5+
* `MicrosoftTeams` PowerShell module installed

  ```powershell
  Install-Module MicrosoftTeams -Force -AllowClobber
  ```
* Proper permissions to access Teams policies via Graph

---

## 🧾 How to Use

1. **Run the script** using PowerShell:

   ```powershell
   .\TeamsPolicyMatrixGui.ps1
   ```
2. Click **"Connect to Teams"**
3. Click **"Load All Policies"**
4. Select policies to **View** or **Compare**
5. Watch real-time logs, status, and summary

---

## ✏️ Customization Ideas

* Export to CSV or JSON
* Add filtering for loaded vs. failed policies
* Support for differential export (drift-only)
* Group policies by category

---

## 👨‍💻 Author

**Edgar Avellan**
🏷️ Microsoft Teams Automation & Security Engineer
🛠️ Building tools to visualize, compare, and teach real-world M365 skills.

---

## 📁 Repository Structure

```
/TeamsPolicyMatrixGui
|-- TeamsPolicyMatrixGui.ps1   # Main GUI script
|-- README.md                  # This file
```

---

## 📸 Preview

> (Add screenshots of the GUI, policy listbox, and comparison output here for visual aid)

---

## 📬 Feedback

Feel free to fork, submit issues, or suggest improvements. Collaboration welcome!

---

## 📜 License

MIT License
