import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel
from pdf_to_excel import PdfToExcelConverter  # Importa la clase PdfToExcelConverter desde el módulo pdf_to_excel

class PdfToExcelGUI(QMainWindow):
    def __init__(self):
        super().__init__()

        # Configuración de la ventana principal
        self.setWindowTitle("PDF to Excel Converter")  # Título de la ventana
        self.setGeometry(100, 100, 400, 200)  # Posición y tamaño de la ventana

        # Etiqueta para indicar la acción de seleccionar un archivo PDF
        self.file_label = QLabel("Select PDF File:", self)
        self.file_label.move(20, 20)  # Posición de la etiqueta

        # Botón para seleccionar un archivo PDF
        self.file_button = QPushButton("Browse", self)
        self.file_button.move(150, 15)  # Posición del botón
        self.file_button.clicked.connect(self.browse_pdf)  # Conexión del botón con el método browse_pdf

        # Botón para iniciar la conversión a Excel
        self.convert_button = QPushButton("Convert", self)
        self.convert_button.move(150, 60)  # Posición del botón
        self.convert_button.clicked.connect(self.convert_to_excel)  # Conexión del botón con el método convert_to_excel

        # Etiqueta para mostrar el estado de la conversión
        self.status_label = QLabel("", self)
        self.status_label.move(20, 100)  # Posición de la etiqueta
        self.status_label.setMinimumWidth(400)  # Ancho mínimo de la etiqueta

        self.pdf_path = ""  # Ruta del archivo PDF seleccionado

    def browse_pdf(self):
        # Abre un diálogo para seleccionar un archivo PDF
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Open PDF File")  # Obtiene la ruta del archivo seleccionado
        if file_path:
            self.pdf_path = file_path  # Almacena la ruta del archivo PDF seleccionado
            self.status_label.setText("PDF file selected.")  # Actualiza el texto de la etiqueta de estado
        else:
            self.status_label.setText("No file selected.")  # Actualiza el texto de la etiqueta de estado

    def convert_to_excel(self):
        if self.pdf_path:  # Verifica si se ha seleccionado un archivo PDF
            try:
                # Crea una instancia de la clase PdfToExcelConverter
                converter = PdfToExcelConverter()

                # Extrae las tablas del archivo PDF y realiza la conversión a Excel
                tables = converter.extract_tables_from_pdf(self.pdf_path)
                combined_table = converter.combine_tables(tables)
                combined_table = converter.split_columns_by_newline(combined_table)
                combined_table = converter.remove_rows_starting_with_puesto(combined_table)
                output_excel_path = "output.xlsx"
                combined_table.to_excel(output_excel_path, index=False, header=False)
                converter.apply_bold_font_to_first_row(output_excel_path)

                # Actualiza el texto de la etiqueta de estado
                self.status_label.setText("Conversion successful! Output saved as output.xlsx")
            except Exception as e:
                # En caso de error, muestra el mensaje de error en la etiqueta de estado
                self.status_label.setText(f"Error: {str(e)}")
        else:
            # Si no se ha seleccionado un archivo PDF, muestra un mensaje en la etiqueta de estado
            self.status_label.setText("Please select a PDF file.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfToExcelGUI()
    window.show()
    sys.exit(app.exec_())
