# 📱 Smart Document Extractor: Android APK Conversion Guide

This guide details how `document_extractor.py` was adapted and how to compile it into a **standalone Android APK (`.apk`)**.

---

## 🏗️ 1. Why Desktop Python Scripts Require Mobile Adaptation

Converting a desktop Python script like `document_extractor.py` into a standalone Android APK requires solving three core architectural differences:

| Desktop (`document_extractor.py`) | Mobile Android Adaptation (`android/`) |
| :--- | :--- |
| **GUI Framework**: Desktop `tkinter` | **Kivy Mobile UI** (`android/main.py`) with touch-friendly layout, responsive buttons, and live log view. |
| **Tesseract OCR**: Subprocess calling desktop `tesseract.exe` in `C:\Program Files` | **Cross-platform OCR bridge** (`extractor_core.py`) checking Android internal app paths, Google ML Kit, or bundled ARM64 tesseract. |
| **PDF Rendering**: Desktop `pymupdf` (C/C++ MuPDF) | **Universal PDF Engine**: Supports `pymupdf`, `pypdfium2`, and Android's native OS `android.graphics.pdf.PdfRenderer`. |
| **File Dialogs**: `filedialog.askopenfilename()` | **Mobile File Picker**: Android Storage Access Framework (`FileChooserListView` / Android storage paths). |

All adapted code is located in the `android/` directory.

---

## 🚀 2. Three Ways to Build the Standalone `.apk`

Because `python-for-android` (Buildozer) cross-compiles Python and native C-extensions for ARM64 using Linux toolchains (Clang, Autotools, NDK), it **cannot run natively on Windows Command Prompt**.

Choose one of the 3 simple methods below:

---

### Option A: 1-Click Cloud Build with GitHub Actions (Recommended, 100% Free)
*No Linux, no Android NDK, and no SDK setup needed on your PC.*

1. Push this repository to GitHub (or upload the files to a GitHub repo).
2. The pre-configured workflow at `.github/workflows/build_apk.yml` is already in place.
3. On GitHub, navigate to **Actions** -> **Build Android APK with Buildozer** -> click **Run workflow**.
4. GitHub's Ubuntu runner installs Buildozer, Android NDK, compiles the app, and provides the compiled **`SmartDocExtractor-1.0.0-arm64-v8a-debug.apk`** under **Artifacts** for direct download!

---

### Option B: Free Google Colab Build (Build in Your Browser)
*Takes ~10 minutes, zero software installed on your machine.*

1. Open [Google Colab](https://colab.research.google.com).
2. Upload the notebook file: [`android/build_apk_colab.ipynb`](file:///c:/Users/uif57044/OneDrive%20-%20Aumovio%20SE/OFFICE%20DESKTOP/my%20documents/Personal/Django/myenv/Scripts/python_projects/AI_Python/document_extractor/android/build_apk_colab.ipynb).
3. Upload `main.py`, `extractor_core.py`, and `buildozer.spec`.
4. Run all cells in sequence.
5. The notebook will automatically compile the APK and trigger a browser download of the `.apk` file!

---

### Option C: Local Build on Windows via WSL2 (Ubuntu)
If you have WSL2 (Ubuntu) installed on Windows:

```bash
# 1. Open Ubuntu terminal in WSL2
sudo apt update
sudo apt install -y git zip unzip openjdk-17-jdk autoconf libtool pkg-config zlib1g-dev \
    libncurses5-dev libncursesw5-dev libtinfo5 cmake libffi-dev libssl-dev

# 2. Install Buildozer
pip3 install --user --upgrade buildozer cython virtualenv

# 3. Navigate to your android directory
cd "/mnt/c/Users/uif57044/OneDrive - Aumovio SE/OFFICE DESKTOP/my documents/Personal/Django/myenv/Scripts/python_projects/AI_Python/document_extractor/android"

# 4. Compile APK (outputs to ./bin/)
buildozer android debug
```

The resulting file will be in `android/bin/smartdocextractor-1.0.0-arm64-v8a-debug.apk`.

---

## 📲 3. Installing and Running the APK on Android

1. Transfer the `.apk` to your phone via USB cable, Google Drive, or WhatsApp/Telegram.
2. On your Android phone, tap on the `.apk` file.
3. If prompted, allow **"Install unknown apps"** from your file manager / browser.
4. Tap **Install** and open **Smart PDF Extractor**!
5. When prompted, grant **Storage / All Files access** so the app can read your input PDFs and save the generated Word (.docx) and PowerPoint (.pptx) files to your phone's `Download/SmartDocExtractor` folder.

---

## 💻 4. Bonus: Build Single Standalone Windows `.exe`

If you also need a single standalone executable for Windows desktop:

```powershell
pyinstaller --noconfirm --onedir --windowed `
  --name "Smart_Document_Extractor" `
  --collect-all docx `
  --collect-all pptx `
  --collect-all fitz `
  --collect-all PIL `
  document_extractor.py
```

