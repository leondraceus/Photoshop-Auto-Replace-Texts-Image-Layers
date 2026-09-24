# Photoshop-Auto-Replace-Texts-Image-Layers
A simple tool for Adobe Photoshop that batch-replaces texts, image layers from CSV data, saving hours of manual copy-paste content I have spent for 2 years without creating this after the recent event with VIP Cards for 2000+ VIP guests lmao.

---

## ✨ **Features**

- **Automated Text & Image Replacement:** Swaps out Text and Smart Object/Image layers directly from CSV columns.
- **Layer Export Control:** Export directly to **PNG**, **JPEG**, **PSD** or **TIFF** with options to flatten or merged layer file.
- **Dynamic File Naming:** Define custom output naming rules (e.g., `[DocName]_[Name]_[ID].png`).

## 🚀 **Installation**
### 1. Download the Script:
Download the latest CSV_Auto_Replace.jsx file from the [Release](https://github.com/leondraceus/Photoshop-Auto-Replace-Texts-Image-Layers/releases/new) page.

### 2. Copy to Photoshop Scripts Folder:
Place the .jsx file into your Photoshop installation directory:

- **Windows:** C:\Program Files\Adobe\Adobe Photoshop [Version]\Presets\Scripts\

- **macOS:** /Applications/Adobe Photoshop [Version]/Presets/Scripts/

### 3. Restart Photoshop:
Once installed, access the script in Photoshop from File > Scripts > CSV_Auto_Replace.

### Note:
Alternatively, you can run the script without moving it by opening Photoshop and navigating to File > Scripts > Browse... and selecting **CSV_Auto_Replace.jsx**.

## 🛠️ **Usage Guide**
### 1. Prepare Your PSD Template:
- Save a Copy of your main PSD file as Template
- Rename target *text* and *image* layers in your Photoshop document so they match your CSV header names exactly.

### 2. Prepare Your CSV Data File [Template](https://docs.google.com/spreadsheets/d/1vvMicyKHjj3Gf_dk8k1lj1bAyhH0zdDlACngVipnRfE/edit?gid=0#gid=0)
Create a UTF-8 encoded .csv file with column names matching your PSD layer names.
