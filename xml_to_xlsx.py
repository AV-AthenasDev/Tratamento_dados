

# criar ambiente virtual
# python -m venv nome_ambiente_virtual
# converte arquivo
#pyside6-uic main_xmlo.ui -o main_xml.py  pyside6-uic main_hid.ui -o main_hid.py
# pyside6-rcc Icons-AF.qrc -o Icons_cr.py
# quando tem janela que interage precisa por o pyinstaller --onefile PyAF04.py -i icone_hid.ico  -w 

#pyinstaller --onefile XML_to_XLSX.py -i XmlToXslx.ico -w

# xmlconverter\Scripts\activate   ativar ambiente virtual isso no powershell 




    
import sys
import os
import csv
import pandas as pd
from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog, QPushButton, QVBoxLayout, QWidget, QTextEdit, QLabel, QMessageBox
from PySide6.QtCore import Qt, QEvent
from PySide6.QtGui import QPixmap, QIcon
from bs4 import BeautifulSoup

class CustomButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

        self.setStyleSheet("background-color: rgb(100, 100, 100); color: rgb(255, 255, 255); height: 30px;")
        self.setStyleSheet_hover = "background-color: rgb(125, 125, 125); color: rgb(255, 255, 255); height: 30px;"
        
        self.setCursor(Qt.PointingHandCursor)  # Muda o cursor para indicar que é clicável

        self.installEventFilter(self)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Enter:
            self.setStyleSheet(self.setStyleSheet_hover)
        elif event.type() == QEvent.Leave:
            self.setStyleSheet("background-color: rgb(100, 100, 100); color: rgb(255, 255, 255); height: 30px;")
        return super().eventFilter(obj, event)

class XMLToDataFrameConverter(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("XML to DataFrame Converter")
        self.setGeometry(100, 100, 800, 600)

        # Defina o ícone da janela
        self.setWindowIcon(QIcon("T:/MEUS CODIGOS/XML to XLSX converter/xmlconverter/XmlToXslx32.png"))  # Ajuste o nome do arquivo de ícone e o caminho conforme necessário

        # Personalizando o estilo da janela
        self.setStyleSheet("background-color: rgb(64, 64, 64); color: rgb(255, 255, 255);")

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        self.open_button = CustomButton("Abrir arquivos XML")
        self.open_button.clicked.connect(self.open_xml_files)
        self.layout.addWidget(self.open_button)

        self.convert_button = CustomButton("Converter para XLSX")
        self.convert_button.clicked.connect(self.convert_to_xlsx)
        self.convert_button.setEnabled(False)
        self.layout.addWidget(self.convert_button)

        self.xml_content_textedit = QTextEdit()
        self.xml_content_textedit.setPlaceholderText("Conteúdo do arquivo XML")
        self.layout.addWidget(self.xml_content_textedit)

        # Adicione o QLabel com o logotipo e texto "Athena Devs"
        logo_label = QLabel("Athena Devs")
        logo_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(logo_label)

        # Adicione um QLabel com um ícone
        label_logo = QLabel()
        label_logo_pixmap = QPixmap("T:/MEUS CODIGOS/XML to XLSX converter/xmlconverter/logo32.png")  # Ajuste o nome do arquivo de ícone e o caminho conforme necessário
        label_logo.setPixmap(label_logo_pixmap)
        self.layout.addWidget(label_logo, alignment=Qt.AlignCenter)

        # Adicione o QLabel com o texto "athenadevs9.gmail.com"
        label_contato = QLabel("athenadevs9@gmail.com")
        label_contato.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(label_contato)

        self.central_widget.setLayout(self.layout)

    def open_xml_files(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setNameFilter("Arquivos XML (*.xml);;Todos os arquivos (*)")

        if file_dialog.exec_():
            self.xml_files = file_dialog.selectedFiles()
            self.convert_button.setEnabled(True)
            self.xml_content_textedit.clear()
            for xml_file in self.xml_files:
                with open(xml_file, 'r') as file:
                    xml_content = file.read()
                    self.xml_content_textedit.append(xml_content)

    def convert_to_xlsx(self):
        if hasattr(self, "xml_files"):
            save_dir = QFileDialog.getExistingDirectory(self, "Escolha a pasta de destino para salvar os arquivos XLSX")

            for xml_file in self.xml_files:
                try:
                    xml_content = self.xml_content_textedit.toPlainText()

                    soup = BeautifulSoup(xml_content, 'html.parser')
                    data = []

                    for tag in soup.find_all(True):
                        if tag.text:
                            data.append([tag.name, tag.text.strip()])

                    df = pd.DataFrame(data, columns=['Tag', 'Conteúdo'])

                    xlsx_filename = os.path.splitext(os.path.basename(xml_file))[0] + ".xlsx"
                    xlsx_path = os.path.join(save_dir, xlsx_filename)

                    df.to_excel(xlsx_path, index=False)
                    print(f"DataFrame salvo como XLSX em: {xlsx_path}")

                    # Exiba uma caixa de mensagem de sucesso
                    QMessageBox.information(self, "Conversão Concluída", "Atividade de conversão para XLSX concluída com sucesso!")

                except Exception as e:
                    print(f"Erro ao converter {xml_file}: {e}")

            self.convert_button.setEnabled(False)
            self.xml_content_textedit.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    converter = XMLToDataFrameConverter()
    converter.show()
    sys.exit(app.exec())
