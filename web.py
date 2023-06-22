from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
from PyQt5.QtWidgets import QApplication, QWidget, QGridLayout, QLabel, QPushButton, QFileDialog, QComboBox

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Generate RD")
        self.layout = QGridLayout()
        self.sheet_combo_box = QComboBox()


        # Labels
        self.word_template_label = QLabel("Word Template:")
        self.word_template_label_toc = QLabel("Word Template TOC:")
        self.excel_label = QLabel("Excel File:")
        self.output_dir_label = QLabel("Output Directory:")
        self.output_dir_label_toc = QLabel("Output Directory TOC:")

        # Buttons
        self.word_template_button = QPushButton("Select Word Template")
        self.word_template_button_toc = QPushButton("Select Word Template TOC")
        self.excel_button = QPushButton("Select Excel File")
        self.output_dir_button = QPushButton("Select Output Directory")
        self.output_dir_button_toc = QPushButton("Select Output Directory TOC")
        self.generate_button = QPushButton("Generate RD AND TOC")

        # Button connections
        self.word_template_button.clicked.connect(self.select_word_template)
        self.word_template_button_toc.clicked.connect(self.select_word_template_toc)
        self.excel_button.clicked.connect(self.select_excel)
        self.output_dir_button.clicked.connect(self.select_output_dir)
        self.output_dir_button_toc.clicked.connect(self.select_output_dir_toc)
        self.generate_button.clicked.connect(self.generate_rd_toc)

        # Add widgets to layout
        self.layout.addWidget(self.word_template_label, 0, 0)
        self.layout.addWidget(self.word_template_label_toc, 1, 0)
        self.layout.addWidget(self.word_template_button, 0, 1)
        self.layout.addWidget(self.word_template_button_toc, 1, 1)
        self.layout.addWidget(self.excel_label, 2, 0)
        self.layout.addWidget(self.excel_button, 2, 1)
        self.layout.addWidget(self.output_dir_label, 3, 0)
        self.layout.addWidget(self.output_dir_label_toc, 4, 0)
        self.layout.addWidget(self.output_dir_button, 3, 1)
        self.layout.addWidget(self.output_dir_button_toc, 4, 1)
        self.layout.addWidget(self.generate_button, 5, 1)
        self.layout.addWidget(self.sheet_combo_box, 2, 2)  # Add the dropdown list

        self.setLayout(self.layout)

    def select_word_template(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Word Template", "", "Word Files (*.docx)", options=options)
        self.word_template_path = Path(file_path)

    def select_word_template_toc(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Word Template TOC", "", "Word Files (*.docx)", options=options)
        self.word_template_path_toc = Path(file_path)

    def select_excel(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)", options=options)
        self.excel_path = Path(file_path)

        if self.excel_path:
            excel_data = pd.ExcelFile(self.excel_path)
            sheet_names = excel_data.sheet_names
            self.sheet_combo_box.clear()
            self.sheet_combo_box.addItems(sheet_names)


    def select_output_dir(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        dir_path = QFileDialog.getExistingDirectory(self, "Select Output Directory", options=options)
        self.output_dir = Path(dir_path)

    def select_output_dir_toc(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        dir_path = QFileDialog.getExistingDirectory(self, "Select Output Directory TOC", options=options)
        self.output_dir_toc = Path(dir_path)

    def generate_rd_toc(self):
        selected_sheet = self.sheet_combo_box.currentText()
        df = pd.read_excel(self.excel_path, sheet_name=selected_sheet)


        for record in df.to_dict(orient="records"):
            doc = DocxTemplate(self.word_template_path)
            doc.render(record)
            output_path = self.output_dir / f"{record['TitleRD']}.docx"
            doc.save(output_path)

        for record in df.to_dict(orient="records"):
            doc = DocxTemplate(self.word_template_path_toc)
            doc.render(record)
            output_path_toc = self.output_dir_toc / f"{record['TitleTOC']}.docx"
            doc.save(output_path_toc)

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()
