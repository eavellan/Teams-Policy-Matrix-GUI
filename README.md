# 📊 Teams Policy Matrix GUI  
*Built with ❤️ in PowerShell by Edgar Avellan*

---

## 👨‍💻 Purpose  
This PowerShell-based GUI app allows Microsoft Teams admins to **connect**, **load**, **view**, and **compare** all available Teams policies across your tenant — offering a real-time look at policy configurations with drift detection between any two selected policies.

---

## ✨ Features (Current Version)

### 🔐 Connect to Teams  
- Uses the `MicrosoftTeams` module to authenticate and establish a connection.

### 🔄 Load All Policies  
- Automatically pulls and parses over **50+ Teams policy types**, separating loaded vs failed types.
- Shows a full **📊 Summary Report** with success/fail counts and types.

### 📋 View Selected  
- Select up to **two policies** and view them **side-by-side** in a popup.
- Displays property names and values for each.

### 🔍 Compare Selected  
- Select **exactly two policies** and compare them line by line.
- Matching settings appear in **🟢 green**
- Drifted/different settings appear in **🔴 red**
- Clear visual diff for easy auditing and compliance checks.

---

## 🧠 Tech Stack

- **PowerShell (WinForms)**
- `MicrosoftTeams` module
- RichTextBox UI elements for formatted text output

---

## 🛣️ Coming Soon (Planned Features)

- 📝 **Export Feature**  
  Export side-by-side comparisons or policy data to `.csv` or `.json` format.

- ⬆️ **Upload Feature**  
  Compare uploaded `.csv` exports to check for drift over time (historical baseline vs current state).

- 🔎 **Search/Filter**  
  Add filtering by policy type or name to quickly narrow results.

- 🧠 **Drift Intelligence**  
  Smarter diff logic for nested objects or JSON blobs inside properties.

---

## 🚀 How to Use

1. **Run as Admin**: Ensure PowerShell is running with administrative privileges.
2. **Connect to Teams** using the GUI button.
3. **Load All Policies** from your tenant.
4. **Select** one or two from the list:
    - Use “View Selected” for side-by-side view.
    - Use “Compare Selected” to see drift in color-coded output.

---

## 📎 Example Output

```
🟢 ✅ Matching settings will appear in green.
🔴 ❌ Differences will be highlighted in red.

✔ AllowUserCreateUpdateRemoveConnectors = True
✘ AllowUserCreateUpdateRemoveTabs:
   ↪ [Policy A] True
   ↪ [Policy B] False
```
![Teams_Policy_Matrix](https://github.com/user-attachments/assets/b57fcc95-2502-4227-b8b7-6dac05a97a4a)

![Policy_Drift](https://github.com/user-attachments/assets/5ac5b244-e4de-4d18-b033-2f85825fb630)

![Policy_Drift2](https://github.com/user-attachments/assets/0fcbcf51-26a6-4496-9b06-6d3dbc8f63bf)



---

## 👑 Author  
**Edgar Avellan**  
PowerShell enthusiast | Microsoft 365 Architect | Security + Automation Nerd  

Connect with me on [LinkedIn](https://www.linkedin.com/in/edgaravellan)  

---

## 📂 File

- `TeamsPolicyMatrix_GUI.ps1` — Main PowerShell GUI script  
- `README.md` — This file
