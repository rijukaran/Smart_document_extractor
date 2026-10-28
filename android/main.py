"""
Smart PDF Document Extractor - Android & Cross-Platform Mobile App
==================================================================
Built with Kivy for Android (packaged with Buildozer)
Extracts OCR text & circuit diagrams directly into Word (.docx) & PowerPoint (.pptx)
"""

import os
import sys
import threading
from pathlib import Path

# Kivy UI framework
import kivy
from kivy.app import App
from kivy.clock import Clock, mainthread
from kivy.core.window import Window
from kivy.metrics import dp, sp
from kivy.properties import BooleanProperty, NumericProperty, StringProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.label import Label
from kivy.uix.modalview import ModalView
from kivy.uix.progressbar import ProgressBar
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput

# Import extraction engine
from extractor_core import process_single_pdf

# Set mobile default background color (Clean Executive Slate)
Window.clearcolor = (0.95, 0.96, 0.98, 1.0)


# =====================================================
# ANDROID PLATFORM HOOKS & PERMISSIONS
# =====================================================

IS_ANDROID = False
try:
    from android.permissions import Permission, request_permissions
    IS_ANDROID = True
except ImportError:
    pass


def request_android_permissions():
    """Request runtime permissions required for reading/writing storage on Android."""
    if IS_ANDROID:
        try:
            permissions = [
                Permission.READ_EXTERNAL_STORAGE,
                Permission.WRITE_EXTERNAL_STORAGE,
            ]
            request_permissions(permissions)
        except Exception as e:
            print(f"Android permission request error: {e}")


def get_default_output_dir():
    """Returns safe user-visible storage path on Android or desktop."""
    if IS_ANDROID:
        try:
            from android.storage import primary_external_storage_path
            storage_path = primary_external_storage_path()
            out_dir = os.path.join(storage_path, "Download", "SmartDocExtractor")
            os.makedirs(out_dir, exist_ok=True)
            return out_dir
        except Exception:
            pass

    # Desktop / fallback
    user_home = Path.home()
    out_dir = user_home / "Documents" / "SmartDocExtractor"
    out_dir.mkdir(parents=True, exist_ok=True)
    return str(out_dir)


# =====================================================
# TOUCH-FRIENDLY POPUP FILE PICKER
# =====================================================

class PDFChooserPopup(ModalView):
    """Touch-friendly full-screen modal file picker for selecting PDF documents."""
    def __init__(self, on_select_callback, **kwargs):
        super().__init__(**kwargs)
        self.size_hint = (0.94, 0.92)
        self.auto_dismiss = True
        self.on_select_callback = on_select_callback

        layout = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(8))

        # Title
        title_label = Label(
            text="Select PDF Document",
            size_hint_y=None,
            height=dp(40),
            font_size=sp(18),
            bold=True,
            color=(0.12, 0.23, 0.54, 1.0)
        )
        layout.add_widget(title_label)

        # Start directory
        initial_path = os.getcwd()
        if IS_ANDROID:
            try:
                from android.storage import primary_external_storage_path
                initial_path = primary_external_storage_path()
            except Exception:
                pass

        self.file_chooser = FileChooserListView(
            path=initial_path,
            filters=["*.pdf", "*.PDF"],
            size_hint=(1, 1)
        )
        layout.add_widget(self.file_chooser)

        # Buttons bar
        btn_bar = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(10))

        cancel_btn = Button(
            text="Cancel",
            background_color=(0.85, 0.85, 0.85, 1.0),
            color=(0.2, 0.2, 0.2, 1.0)
        )
        cancel_btn.bind(on_release=lambda x: self.dismiss())
        btn_bar.add_widget(cancel_btn)

        select_btn = Button(
            text="Open Selected",
            background_color=(0.14, 0.40, 0.80, 1.0),
            color=(1.0, 1.0, 1.0, 1.0),
            bold=True
        )
        select_btn.bind(on_release=self._confirm_selection)
        btn_bar.add_widget(select_btn)

        layout.add_widget(btn_bar)
        self.add_widget(layout)

    def _confirm_selection(self, instance):
        selection = self.file_chooser.selection
        if selection and len(selection) > 0:
            chosen = selection[0]
            if chosen.lower().endswith(".pdf"):
                self.on_select_callback(chosen)
                self.dismiss()


# =====================================================
# MAIN APPLICATION ROOT WIDGET
# =====================================================

class MainExtractorWidget(BoxLayout):
    selected_pdf = StringProperty("")
    status_text = StringProperty("Ready. Please select a PDF file.")
    progress_val = NumericProperty(0)
    is_processing = BooleanProperty(False)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = dp(14)
        self.spacing = dp(10)

        self.output_folder = get_default_output_dir()
        self.log_messages = []

        self._build_ui()
        request_android_permissions()

    def _build_ui(self):
        # 1. Header Card (Navy Blue #1E3A8A)
        header = BoxLayout(orientation='vertical', size_hint_y=None, height=dp(64), padding=dp(6))
        title = Label(
            text="Smart PDF Extractor",
            font_size=sp(20),
            bold=True,
            color=(0.12, 0.23, 0.54, 1.0),
            halign='center'
        )
        subtitle = Label(
            text="Text, Circuits & Diagrams into Word (.docx) & PPTX",
            font_size=sp(12),
            color=(0.35, 0.45, 0.60, 1.0),
            halign='center'
        )
        header.add_widget(title)
        header.add_widget(subtitle)
        self.add_widget(header)

        # 2. File Selection Section
        file_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(48), spacing=dp(8))
        self.file_label = Label(
            text="No PDF Selected",
            font_size=sp(13),
            color=(0.2, 0.2, 0.2, 1.0),
            halign='left',
            valign='middle',
            shorten=True,
            shorten_from='center'
        )
        self.file_label.bind(size=self.file_label.setter('text_size'))
        file_box.add_widget(self.file_label)

        select_btn = Button(
            text="Choose PDF",
            size_hint_x=None,
            width=dp(125),
            background_color=(0.14, 0.40, 0.80, 1.0),
            color=(1, 1, 1, 1),
            bold=True
        )
        select_btn.bind(on_release=self.show_file_chooser)
        file_box.add_widget(select_btn)
        self.add_widget(file_box)

        # 3. Output Format Options (Checkboxes)
        opts_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(36), spacing=dp(4))
        
        # Docx Checkbox
        self.chk_docx = CheckBox(active=True, size_hint_x=None, width=dp(28))
        opts_box.add_widget(self.chk_docx)
        opts_box.add_widget(Label(text="Word", font_size=sp(12), color=(0.2, 0.2, 0.2, 1.0), size_hint_x=None, width=dp(42)))

        # PPTX Checkbox
        self.chk_ppt = CheckBox(active=True, size_hint_x=None, width=dp(28))
        opts_box.add_widget(self.chk_ppt)
        opts_box.add_widget(Label(text="PPTX", font_size=sp(12), color=(0.2, 0.2, 0.2, 1.0), size_hint_x=None, width=dp(42)))

        # TXT Checkbox
        self.chk_txt = CheckBox(active=True, size_hint_x=None, width=dp(28))
        opts_box.add_widget(self.chk_txt)
        opts_box.add_widget(Label(text="TXT", font_size=sp(12), color=(0.2, 0.2, 0.2, 1.0), size_hint_x=None, width=dp(36)))

        # Page Range input
        opts_box.add_widget(Label(text="Pages:", font_size=sp(12), color=(0.2, 0.2, 0.2, 1.0), size_hint_x=None, width=dp(48)))
        self.page_input = TextInput(
            text="all",
            multiline=False,
            font_size=sp(12),
            size_hint=(1, 1)
        )
        opts_box.add_widget(self.page_input)
        self.add_widget(opts_box)

        # 4. Action Button (Start Extraction)
        self.action_btn = Button(
            text="START EXTRACTION",
            size_hint_y=None,
            height=dp(52),
            background_color=(0.06, 0.72, 0.50, 1.0),  # Emerald Green
            color=(1, 1, 1, 1),
            font_size=sp(16),
            bold=True
        )
        self.action_btn.bind(on_release=self.start_extraction)
        self.add_widget(self.action_btn)

        # 5. Progress Bar & Status
        progress_box = BoxLayout(orientation='vertical', size_hint_y=None, height=dp(40), spacing=dp(4))
        self.progress_bar = ProgressBar(max=100, value=0, size_hint_y=None, height=dp(10))
        self.status_lbl = Label(
            text="Ready",
            font_size=sp(12),
            color=(0.4, 0.4, 0.4, 1.0),
            size_hint_y=None,
            height=dp(22)
        )
        progress_box.add_widget(self.progress_bar)
        progress_box.add_widget(self.status_lbl)
        self.add_widget(progress_box)

        # 6. Live Console Log
        log_header = Label(
            text="Live Extraction Log:",
            font_size=sp(13),
            bold=True,
            color=(0.2, 0.2, 0.2, 1.0),
            size_hint_y=None,
            height=dp(24),
            halign='left'
        )
        log_header.bind(size=log_header.setter('text_size'))
        self.add_widget(log_header)

        scroll = ScrollView(size_hint=(1, 1), bar_width=dp(6))
        self.log_view = Label(
            text="App initialized. Select a PDF file to begin extraction.",
            font_size=sp(11),
            color=(0.15, 0.15, 0.15, 1.0),
            size_hint_y=None,
            halign='left',
            valign='top'
        )
        self.log_view.bind(texture_size=lambda instance, size: setattr(instance, 'height', max(size[1], dp(200))))
        self.log_view.bind(size=lambda instance, size: setattr(instance, 'text_size', (size[0], None)))
        scroll.add_widget(self.log_view)
        self.add_widget(scroll)

    def show_file_chooser(self, instance):
        if self.is_processing:
            return
        popup = PDFChooserPopup(on_select_callback=self.on_pdf_selected)
        popup.open()

    def on_pdf_selected(self, file_path):
        self.selected_pdf = file_path
        file_name = os.path.basename(file_path)
        self.file_label.text = f"Selected: {file_name}"
        self.append_log(f"Selected file: {file_path}")

    @mainthread
    def append_log(self, text, level="INFO"):
        self.log_messages.append(f"[{level}] {text}")
        if len(self.log_messages) > 300:
            self.log_messages = self.log_messages[-300:]
        self.log_view.text = "\n".join(self.log_messages)

    @mainthread
    def update_progress(self, current, total):
        pct = int((current / max(1, total)) * 100)
        self.progress_bar.value = pct
        self.status_lbl.text = f"Processing page {current} of {total} ({pct}%)"

    def start_extraction(self, instance):
        if self.is_processing:
            return
        if not self.selected_pdf or not os.path.isfile(self.selected_pdf):
            self.append_log("Please choose a valid PDF file first.", "ERROR")
            self.status_lbl.text = "Error: No PDF selected."
            return

        self.is_processing = True
        self.action_btn.disabled = True
        self.progress_bar.value = 0
        self.status_lbl.text = "Starting extraction..."
        self.append_log(f"Beginning extraction pipeline...")
        self.append_log(f"Output directory: {self.output_folder}")

        # Run extraction in background thread to keep UI completely responsive
        threading.Thread(target=self._worker_thread, daemon=True).start()

    def _worker_thread(self):
        try:
            ok, report = process_single_pdf(
                pdf_path=self.selected_pdf,
                output_folder=self.output_folder,
                log_func=self.append_log,
                progress_func=self.update_progress,
                opt_docx=self.chk_docx.active,
                opt_ppt=self.chk_ppt.active,
                opt_txt=self.chk_txt.active,
                opt_verify=True,
                page_range_str=self.page_input.text
            )
            self._on_finished(ok)
        except Exception as exc:
            self.append_log(f"Unexpected fatal error: {exc}", "ERROR")
            self._on_finished(False)

    @mainthread
    def _on_finished(self, success):
        self.is_processing = False
        self.action_btn.disabled = False
        self.progress_bar.value = 100
        if success:
            self.status_lbl.text = "Extraction completed successfully!"
            self.append_log(f"Files saved in: {self.output_folder}", "SUCCESS")
        else:
            self.status_lbl.text = "Extraction failed. Review log."


class SmartDocExtractorApp(App):
    title = "Smart PDF Extractor"

    def build(self):
        return MainExtractorWidget()


if __name__ == '__main__':
    SmartDocExtractorApp().run()

