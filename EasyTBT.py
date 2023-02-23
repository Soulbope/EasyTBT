import datetime
import sys
import webbrowser
import json
from xml.etree.ElementTree import tostring
import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets
from designEasyTBT import *
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtWidgets import QMainWindow, QApplication, QInputDialog, QMessageBox, QCompleter, QLineEdit
from docxtpl import DocxTemplate
import ctypes
from pkg_resources import resource_filename
import os
import win32print
import win32api
import time
import shutil


if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)


class Novo(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.oldPos = self.pos() #poder mover a tela clicando e arrastando em qualquer lugar
        self.combobox() #preencher combobox
        self.setWindowTitle('EasyTBT')
        self.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'fflaranjaSemFundo.ico'))))
        self.btnEscape.clicked.connect(self.fecha)    
        self.btnSalvar.clicked.connect(self.salvarInputGroup)
        self.btnConfig.clicked.connect(self.minimiza)
        self.btnChoosePrinter.clicked.connect(self.escolherImpressora)
        self.btnImprimir.clicked.connect(self.saveAndPrint)
        self.btnAbrir.clicked.connect(self.escolherInputGroup)
        self.btnLogo.clicked.connect(self.logo)
        self.btnVersion.clicked.connect(self.versionButton)
        self.btnVersion.clicked.connect(self.versionButton)
        self.comboBoxRD.currentIndexChanged.connect(self.on_combobox_changed)
        
        Namecompleter = QCompleter(['ALEXANDRO LIMA DOS SANTOS', 'ANDRE AUGUSTO DOS SANTOS NASCIMENTO', 'BERNARDO BEHNKEN PIMENTA', 'BRUNO PORTO DA SILVA', 'CARLOS HENRIQUE NASCIMENTO DA COSTA', 'CARLOS ROBERTO REIS JUNIOR', 'CHRISTIAN HENRI JOSEPH HERTAY', 'DAMIANA FERREIRA DE LIMA', 'DIEGO SOUZA JULIO', 'DIOGO HENRIQUE ANDRADE CORREA', 'DIONI CANTELLI', 'FABIO HENRIQUE SIMAS ABREU', 'FABIO PAULINO ALVES SOARES', 'FREDERIC PHILIPPE BOUDOUX', 'GIOVANI FERREIRA DA CONCEICAO', 'GLEISON LUIZ DA SILVA', 'HENRIQUE FERNANDES DE ARAUJO', 'JONAS DE ALMEIDA SANTOS', 'LEANDRO DE LIMA DA SILVA', 'LUCAS EDUARDO DE CAIO CARVALHO', 'MARCEL NOGUEIRA PINHO', 'LUCAS LOPES DO NASCIMENTO', 'MARCELO MENDONÇA DOS SANTOS', 'MARCOS ADRIANI SANTOS DE OLIVEIRA', 'MARCOS AURELIO MENDES PIMENTA', 'MARCOS MOISES PEREIRA DOS SANTOS', 'MATHEUS GONCALVES LEITE', 'NORBERT HUGUES ZUNINO', 'RAFAEL GOMES MARTINS', 'ROBERTO ALEXANDRE FERREIRA', 'ROBERTO ALVES JUNIOR', 'RODRIGO JOSE DOS SANTOS RIBEIRO', 'ROGERIO VIEIRA GUIMARAES', 'ROMULO FERREIRA DE ALVARENGA', 'RUAN KAIQUE ANTUNES ANDRE', 'SHIRLEI DA SILVA DE CASTRO', 'THIAGO BRAGA DA PAIXAO', 'WAGNER LENI DE OLIVEIRA JUNIOR', 'WAGNER VIANA VIEIRA', 'WANDERSON DE ASSIS MARTINS', 'WESLEY LOROZA CORREA', 'WILLIAN PIMENTEL BATISTA'])
        Namecompleter.setCaseSensitivity(Qt.CaseInsensitive)
        
        for input in [self.inputExecutante1, self.inputExecutante2, self.inputExecutante3, self.inputExecutante4, self.inputExecutante5,
                      self.inputExecutante6, self.inputExecutante7, self.inputExecutante8, self.inputExecutante9, self.inputExecutante10,
                      self.inputAutoridade
                      ]:
            input.setCompleter(Namecompleter)
            
        Equipcompleter = QCompleter(["10MT GUINCHO #1", "10MT GUINCHO #2", "10MT GUINCHO #3 + POLIA", "10MT GUINCHO #4 + POLIA", "10MT GUINCHO #5", "15MT GUINCHO #1", "15MT GUINCHO #2", "2500 MT BASKET", "2500 MT BASKET - CCR", "2500 MT BASKET - SPOOLING ARM", "2500MT BASKET - ESCADA", "2500MT BASKET - TROLLEY", "30MT GUINCHO #1 + POLIA", "30MT GUINCHO #2", "50MT GUINCHO DE INICIAÇÃO + SWIVEL + POLIA", "A&R MAST 2ND LEVEL LADDER", "ALIGNER_PS SIDE LADDER", "ALIGNER_STB SIDE LADDER", "ALINHADOR + EHS", "ALMOXARIFADO PIPELAY", "ALMOXARIFADO PROJETO", "ARDS", "ARDS - WIRE CLAMP 1 (UPPER)", "ARDS - WIRE CLAMP 2 (LOWER)", "CASCATA (COMPRESSOR BAUER)", "CCR VLS", "CENTRALISADOR 1", "CENTRALISADOR 2", "CORRENTÔMETRO", "DEFLETOR DA TORRE", "DEFLETOR DE CABO A&R (CHUPACABRA)", "DEFLETOR VERTICAL - ROLO (POPA)", "EQUIPAMENTO DO CONVES", "ESCRITORIO DO PIPELAY", "FERRAMENTA HYDRATIGHT 18XLCT", "FERRAMENTA HYDRATIGHT 1MXT", "FERRAMENTA HYDRATIGHT 1MXT", "FERRAMENTA HYDRATIGHT 2XLCT", "FERRAMENTA HYDRATIGHT 2XLCT", "FERRAMENTA HYDRATIGHT 5MXT", "FERRAMENTA HYDRATIGHT 5MXT", "FERRAMENTA HYDRATIGHT 8XLCT", "FERRAMENTA HYDRATIGHT 8XLCT", "FERRAMENTAS", "FERRAMENTAS PROJETO", "FLOWLINE - ESTAÇÃO DE TESTE - BASKET", "FLOWLINE - ESTAÇÃO DE TESTE - MESA", "FLOWLINE - ESTAÇÃO DE TESTE - RDS", "FLOWLINE - SKID", "GUINCHO 5T - CESTA 2500T", "GUINCHO 5T - RDS (BB)", "GUINCHO 5T - RDS (BE)", "GUINCHO ARMAZ. + CABO + SWIVEL + ADAPTADOR", "GUINCHO DE TRAÇÃO", "GUINDASTE DE SERVIÇO", "HANG OFF COLLAR", "HANG OFF COLLAR - NOVO", "HANG-OFF : BODY PLATE LARGE", "HANG-OFF : BODY PLATE MEDIUM", "HANG-OFF : BODY PLATE SMALL", "HOISTING BEAMS 1 + 2", "HPU VLS / HIDRÁULICO", "HYTORC - BOMBA HIDRATIGHT 1", "HYTORC - BOMBA HIDRATIGHT 2", "INJEÇÃO QUEMICAL", "IOS - POLIA", "LOWER LEVEL 1 - STORAGE WINCH SECOND LADDER","LOWER LEVEL 2 - STORAGE WINCH SECOND LADDER", "MANILHA HIDROACÚSTICA (REIGNIER)", "MANILHA HIDROACÚSTICA (SKV)", "MARINE", "MESA", "MÓDULO TENSIONADOR 1", "MÓDULO TENSIONADOR 2", "NITROGÊNIO - RACK", "NITROGÊNIO - SKID", "OFICINA DE MANGUEIRAS", "OFICINA DE SOLDA", "OFICINA ELÉTRICA", "OFICINA MECÂNICA", "RDS 1 - BB", "RDS 2 - BE", "RDS TOWER 01 UPPER LEVEL LADDER", "RDS TOWER 02 UPPER LEVEL LADDER", "ROV", "SERVICE CRANE 1ST LEVEL LADDER", "SERVICE CRANE 2ND LEVEL LADDER", "TAMPA (HATCH PRINCIPAL)", "TENSIONADOR 1", "TENSIONADOR 2", "TENSIONADOR DE CARREGAMENTO", "TENSIONADOR DE CARREGAMENTO - CCR", "TENSIONER 1 BACKWARD MAIN LADDER", "TENSIONER 2 BACKWARD MAIN LADDER", "TORRE VLS", "TRAVA QUEDA", "UMBILICAL - BANCA DE TESTE", "UMBILICAL - SKID", "UNDER DECK"])
        Equipcompleter.setCaseSensitivity(Qt.CaseInsensitive)
        
        self.inputEquipto.setCompleter(Equipcompleter)
                               
        self.inputs = {}
        self.file_name = 'inputs.json'
        self.btnVersion.setText('Feito por WAGNER OLIVEIRA - Ver. 1.0.1')       
        
        #popup de IMPRESSÂO
        self.imsg = QMessageBox()
        self.imsg.setWindowTitle("Imprimindo")
        self.imsg.setText("AGUARDE IMPRESSÃO...")
        self.imsg.setIcon(QMessageBox.Information)
        self.imsg.setStandardButtons(QMessageBox.Ok)
        self.imsg.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'fflaranjaSemFundo.ico'))))
        
        #popup de Preparando arquivos
        self.pmsg = QMessageBox()
        self.pmsg.setWindowTitle("Preparando")
        self.pmsg.setText("AGUARDE PREPARAÇÃO...")
        self.pmsg.setIcon(QMessageBox.Information)
        self.pmsg.setStandardButtons(QMessageBox.Ok)
        self.pmsg.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'fflaranjaSemFundo.ico'))))
                
        #popup de ERRO
        self.emsg = QMessageBox()
        self.emsg.setWindowTitle("Erro")
        self.emsg.setText("Ocorreu um erro de execução")
        self.emsg.setIcon(QMessageBox.Critical)
        self.emsg.setStandardButtons(QMessageBox.Ok)
        self.emsg.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'fflaranjaSemFundo.ico'))))
        
        #popup de Update
        self.umsg = QMessageBox()
        self.umsg.setWindowTitle("Foi mal!")
        self.umsg.setText("Entrei de férias antes de conseguir adicionar esse botão, mas na volta eu faço!")
        self.umsg.setIcon(QMessageBox.Information)
        self.umsg.setStandardButtons(QMessageBox.Ok)
        self.umsg.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'fflaranjaSemFundo.ico'))))
              
        #popup de IMPRESSÂO
        self.msg2 = QMessageBox()
        self.msg2.setWindowTitle("Atenção")
        self.msg2.setText("Arquivo de mesmo nome, favor alterar!")
        self.msg2.setIcon(QMessageBox.Warning)
        self.msg2.setStandardButtons(QMessageBox.Ok)
        self.msg2.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'fflaranjaSemFundo.ico'))))
        
        # Algumas coisas já padrão
        self.dateEditHoje.setText(datetime.date.today().strftime("%d/%m/%Y"))
        self.inputEmbarcacao.setText('Skandi Olinda')
        self.inputNumRD = ''
   
    def logo(self):
        webbrowser.open('https://github.com/Soulbope')  
        
    def versionButton(self):
        webbrowser.open('https://github.com/Soulbope/EasyPT/tree/main#readme')  

    def fecha(self):        
        escolha = ('Salvar', 'Apenas Fechar')
        item, ok = QInputDialog.getItem(self, "Salvar", "Deseja salvar anter de fechar?", escolha, 0, False)
        
        if ok and (item=='Salvar'):
            self.salvarInputGroup()
        elif ok and (item=='Apenas Fechar'):
            self.close()
    
        self.close()                
        
    def minimiza(self):
        self.showMinimized()                  
        
    # Funções para mover a tela segurando em qualquer lugar    
    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        delta = QPoint (event.globalPos() - self.oldPos)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPos = event.globalPos()
    
    #Função para salvar com o Enter    
    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Return or e.key() == Qt.Key_Enter:
            self.salvarInputGroup()
            
                     
        
    
    #Função para add as opções do Combobox
    def combobox(self):
        rdType = ['Não é Rotine Duty','Pequena Manutenção/Reparo Mecânico','Pequena Manutenção/Reparo Hidráulico','Pequena Manutenção/Reparo Elétrico',
                  'Inspeção Torre VLS','Operação Torre VLS','Inspeção equipamento Pipelay','Limpeza e uso de máquina de limpeza de alta pressão',
                  'Colocação de torre VLS nas posições Seafasting e Bridge Passage','Operações com o RDS','Operação de Bancada e Fabricação de Mangueiras',
                  'Util. de equip. de bancada e ferram. gerais nas oficinas de Elétrica, Mecânica e Solda','Utilização de Torno','Troca de sapata'
        ]
        for i in rdType:
            self.comboBoxRD.addItem(i)
            
    def on_combobox_changed(self):
        rdType = self.comboBoxRD.currentText()
        
        match rdType:
                case 'Pequena Manutenção/Reparo Mecânico':#
                    self.inputNumRD = 'RD-SKO-PL-002 rev.1'
                    self.inputJRA.setText('JRA-SKO-PL-022')
                    self.textEditJobDescription.setText('Pequena Manutenção/Reparo Mecânico')
                case 'Pequena Manutenção/Reparo Hidráulico':#
                    self.inputNumRD = 'RD-SKO-PL-003 rev.1'
                    self.inputJRA.setText('JRA-SKO-PL-024')
                    self.textEditJobDescription.setText('Pequena Manutenção/Reparo Hidráulico')
                case 'Pequena Manutenção/Reparo Elétrico':#
                    self.inputNumRD = 'RD-SKO-PL-004 rev.1'
                    self.inputJRA.setText('JRA-SKO-PL-026')
                    self.textEditJobDescription.setText('Pequena Manutenção/Reparo Elétrico')
                case 'Inspeção Torre VLS':#
                    self.inputNumRD = 'RD-SKO-PL-005 rev.3'
                    self.inputJRA.setText('JRA-SKO-PL-027')
                    self.textEditJobDescription.setText('Inspeção Torre VLS')
                    self.inputEquipto.setText('Torre VLS')
                case 'Operação Torre VLS':#
                    self.inputNumRD = 'RD-SKO-PL-006 rev.3'
                    self.inputJRA.setText('JRA-SKO-PL-028')
                    self.textEditJobDescription.setText('Operação Torre VLS')
                    self.inputEquipto.setText('Torre VLS')
                case 'Inspeção equipamento Pipelay':#
                    self.inputNumRD = 'RD-SKO-PL-007 rev.3'
                    self.inputJRA.setText('JRA-SKO-PL-029')
                    self.textEditJobDescription.setText('Inspeção equipamento Pipelay')
                case 'Limpeza e uso de máquina de limpeza de alta pressão':#
                    self.inputNumRD = 'RD-SKO-PL-008 rev.2'
                    self.inputJRA.setText('JRA-SKO-PL-030')
                    self.textEditJobDescription.setText('Limpeza e uso de máquina de limpeza de alta pressão')
                    self.inputEquipto.setText('máquina de limpeza de alta pressão')
                case 'Colocação de torre VLS nas posições Seafasting e Bridge Passage':#
                    self.inputNumRD = 'RD-SKO-PL-009 rev.2'
                    self.inputJRA.setText('JRA-SKO-PL-032')
                    self.textEditJobDescription.setText('Colocação de torre VLS nas posições Seafasting e Bridge Passage')
                    self.inputEquipto.setText('Torre VLS')
                case 'Operações com o RDS':#
                    self.inputNumRD = 'RD-SKO-PL-010 rev.1'
                    self.inputJRA.setText('JRA-SKO-PL-020')
                    self.textEditJobDescription.setText('Operações com o RDS')
                    self.inputEquipto.setText('RDS')
                case 'Operação de Bancada e Fabricação de Mangueiras':#
                    self.inputNumRD = 'RD-SKO-PL-011 rev.1'
                    self.inputJRA.setText('JRA-SKO-PL-064')
                    self.textEditJobDescription.setText('Operação de Bancada e Fabricação de Mangueiras')
                    self.inputEquipto.setText('Bancada de Fabr. de mangueiras')
                case 'Util. de equip. de bancada e ferram. gerais nas oficinas de Elétrica, Mecânica e Solda':#
                    self.inputNumRD = 'RD-SKO-PL-012 rev.1'
                    self.inputJRA.setText('JRA-SKO-PL-065')
                    self.textEditJobDescription.setText('Util. de equip. de bancada e ferram. gerais nas oficinas de Elétrica, Mecânica e Solda')
                    self.inputEquipto.setText('Oficina')
                case 'Utilização de Torno':#
                    self.inputNumRD = 'RD-SKO-PL-013 rev.1'
                    self.inputJRA.setText('JRA-SKO-PL-034')
                    self.textEditJobDescription.setText('Utilização de Torno')
                    self.inputEquipto.setText('Torno (almoxarifado)')
                case 'Troca de sapata':#
                    self.inputNumRD = 'RD-SKO-PL-014 rev.0'
                    self.inputJRA.setText('JRA-SKO-PL-025')
                    self.textEditJobDescription.setText('Troca de sapata')
                    self.inputEquipto.setText('Torre VLS')
                case 'Não é Rotine Duty':#
                    self.inputNumRD = ''
                case _:#
                    self.inputNumRD = ''        
    
        
    def saveAndPrint(self):
        try:
            inputNumRD = self.inputNumRD
            inputJRA = self.inputJRA.text()#
            inputEquipto = self.inputEquipto.text()#
            inputEmbarcacao = self.inputEmbarcacao.text()#
            textEditJobDescription = self.textEditJobDescription.toPlainText()#
            inputAutoridade = self.inputAutoridade.text()#
            dateEditHoje = self.dateEditHoje.text()#
            inputExecutante1 = self.inputExecutante1.text()#
            inputExecutante2 = self.inputExecutante2.text()#
            inputExecutante3 = self.inputExecutante3.text()#
            inputExecutante4 = self.inputExecutante4.text()#
            inputExecutante5 = self.inputExecutante5.text()#
            inputExecutante6 = self.inputExecutante6.text()#
            inputExecutante7 = self.inputExecutante7.text()#
            inputExecutante8 = self.inputExecutante8.text()#
            inputExecutante9 = self.inputExecutante9.text()#
            inputExecutante10 = self.inputExecutante10.text()#
            checkBoxErgonomia = "Sim" if self.checkBoxErgonomia.isChecked() else "Não"
            checkBoxColuna = "Sim" if self.checkBoxColuna.isChecked() else "Não"
            
            
                    
            docTbt = DocxTemplate(os.path.join(resource_filename(__name__, 'tbttemplate.docx')))     
            context = {
                'inputNumRD' : inputNumRD,
                'inputJRA' : inputJRA,
                'textEditJobDescription' : textEditJobDescription,
                'inputAutoridade' : inputAutoridade,
                'dateEditHoje' : dateEditHoje,
                'inputEquipto' : inputEquipto,
                'inputEmbarcacao' : inputEmbarcacao,
                'inputExecutante1' : inputExecutante1,
                'inputExecutante2' : inputExecutante2,
                'inputExecutante3' : inputExecutante3,
                'inputExecutante4' : inputExecutante4,
                'inputExecutante5' : inputExecutante5,
                'inputExecutante6' : inputExecutante6,
                'inputExecutante7' : inputExecutante7,
                'inputExecutante8' : inputExecutante8,
                'inputExecutante9' : inputExecutante9,
                'inputExecutante10' : inputExecutante10,
                'checkBoxErgonomia' : checkBoxErgonomia,
                'checkBoxColuna' : checkBoxColuna,
            }
            docTbt.render(context)
            
            self.pmsg.show()
            QApplication.processEvents()
            
            #cria pasta temporária, salva os arquivos
            pastaTempAtual = os.path.join(resource_filename(__name__, 'temp'))
            os.mkdir(pastaTempAtual) #cria a pasta
            docTbt.save(f'{pastaTempAtual}/tbt.docx')
            
            time.sleep(5)
            
            self.pmsg.close()             
            
            printer_handle = win32print.OpenPrinter(win32print.GetDefaultPrinter())
            status = win32print.GetPrinter(printer_handle, 2)['Status']
            while status == win32print.PRINTER_STATUS_BUSY:
                time.sleep(1)
                status = win32print.GetPrinter(printer_handle, 2)['Status']
            win32print.ClosePrinter(printer_handle)     
                
            #imprime todos os arquivos da pasta
            listaPts = os.listdir(pastaTempAtual)
            try:
                self.imsg.show()
                QApplication.processEvents()
                for arquivo in listaPts:
                    time.sleep(1)
                    win32api.ShellExecute(0, "print", arquivo , None, pastaTempAtual , 0)
                 
            except Exception as e:
                x = self.emsg.setInformativeText(str(e))
                x = self.emsg.exec_()
                print(f'O erro ao inprimir foi: {e}')
                pass
                            
            time.sleep(15)
            
            self.imsg.close()     
            
            try:
                shutil.rmtree(pastaTempAtual) #apaga a pasta
            except Exception as e:
                print(f'O erro é: {e}')
        
        except Exception as e:
            x = self.emsg.setInformativeText(str(e))
            x = self.emsg.exec_()
            print(f'o erro é: {e}')
            pass 
        
        
    def escolherImpressora(self):
        try:
            lista_impressoras = win32print.EnumPrinters(2)
            
            impressoras = []       
                    
            for infoTotalImpressoras in lista_impressoras:
                impressoras.append(infoTotalImpressoras[2])   
                
            item, ok = QInputDialog.getItem(self, "IMPRESSORA", "Selecionar impressora", impressoras, 0, False)
                    
            if ok and item:
                for impressorass in lista_impressoras:
                    if impressorass[2] == item:
                        impressoraAtual = impressorass[2]
                        
            win32print.SetDefaultPrinter(impressoraAtual)
            
        except Exception as e:
            x = self.emsg.setInformativeText(str(e))
            x = self.emsg.exec_()
            print(f'o erro é: {e}')
            pass 
        
    #Preciso colocar um botão pasa salvar o template que o usuário acabou de preencher (deseja salvar essas informações?)    
    def salvarInputGroup(self):
        try:
            inputNumRD = self.inputNumRD
            inputJRA = self.inputJRA.text()#
            inputEquipto = self.inputEquipto.text()#
            inputEmbarcacao = self.inputEmbarcacao.text()#
            textEditJobDescription = self.textEditJobDescription.toPlainText()#
            inputAutoridade = self.inputAutoridade.text()#
            dateEditHoje = self.dateEditHoje.text()#
            comboBoxRD = self.comboBoxRD.currentText()#
            inputExecutante1 = self.inputExecutante1.text()#
            inputExecutante2 = self.inputExecutante2.text()#
            inputExecutante3 = self.inputExecutante3.text()#
            inputExecutante4 = self.inputExecutante4.text()#
            inputExecutante5 = self.inputExecutante5.text()#
            inputExecutante6 = self.inputExecutante6.text()#
            inputExecutante7 = self.inputExecutante7.text()#
            inputExecutante8 = self.inputExecutante8.text()#
            inputExecutante9 = self.inputExecutante9.text()#
            inputExecutante10 = self.inputExecutante10.text()#
            checkBoxErgonomia = self.checkBoxErgonomia.isChecked()#
            checkBoxColuna = self.checkBoxColuna.isChecked()#
            
            name, _ = QtWidgets.QInputDialog.getText(self, 'Salvar', 'Dê um nome para esse template:')

            if name:
                self.inputs[name] = {
                'inputNumRD' : inputNumRD,
                'comboBoxRD' : comboBoxRD,
                'inputJRA' : inputJRA,
                'textEditJobDescription' : textEditJobDescription,
                'inputAutoridade' : inputAutoridade,
                'dateEditHoje' : dateEditHoje,
                'inputEquipto' : inputEquipto,
                'inputEmbarcacao' : inputEmbarcacao,
                'inputExecutante1' : inputExecutante1,
                'inputExecutante2' : inputExecutante2,
                'inputExecutante3' : inputExecutante3,
                'inputExecutante4' : inputExecutante4,
                'inputExecutante5' : inputExecutante5,
                'inputExecutante6' : inputExecutante6,
                'inputExecutante7' : inputExecutante7,
                'inputExecutante8' : inputExecutante8,
                'inputExecutante9' : inputExecutante9,
                'inputExecutante10' : inputExecutante10,
                'checkBoxErgonomia' : checkBoxErgonomia,
                'checkBoxColuna' : checkBoxColuna,
            }         

                with open(os.path.join(resource_filename(__name__, 'inputs.json')), "w") as f:
                    json.dump(self.inputs, f)
                    
            
        except Exception as e:
                x = self.emsg.setInformativeText(str(e))
                x = self.emsg.exec_()
                print(f'o erro ao tentar salvar: {e}')
                pass
    
    #Seria o "Abrir", só que como seria mais complicado, abrir de fato um arquivo, preferi deixar o usuário salvar um template.         
    def escolherInputGroup(self):
        try:             
            if os.path.exists(self.file_name):
                with open(self.file_name, 'r') as f:
                    self.inputs = json.load(f)  
                
            name, _ = QInputDialog.getItem(self, 'Carregar', 'Favor selecionar seu template:', list(self.inputs.keys()), 0, False)

            if name:
                inputs = self.inputs[name]
                                
                self.inputNumRD = inputs['inputNumRD']
                self.inputJRA.setText(inputs['inputJRA'])
                self.comboBoxRD.setCurrentText(inputs['comboBoxRD'])
                self.textEditJobDescription.setText(inputs['textEditJobDescription'])
                self.inputAutoridade.setText(inputs['inputAutoridade']) 
                self.inputEquipto.setText(inputs['inputEquipto']) 
                self.inputEmbarcacao.setText(inputs['inputEmbarcacao']) 
                self.dateEditHoje.setText(inputs['dateEditHoje']) 
                self.inputExecutante1.setText(inputs['inputExecutante1']) 
                self.inputExecutante2.setText(inputs['inputExecutante2']) 
                self.inputExecutante3.setText(inputs['inputExecutante3']) 
                self.inputExecutante1.setText(inputs['inputExecutante1']) 
                self.inputExecutante4.setText(inputs['inputExecutante4']) 
                self.inputExecutante5.setText(inputs['inputExecutante5']) 
                self.inputExecutante6.setText(inputs['inputExecutante6']) 
                self.inputExecutante7.setText(inputs['inputExecutante7']) 
                self.inputExecutante8.setText(inputs['inputExecutante8']) 
                self.inputExecutante9.setText(inputs['inputExecutante9']) 
                self.inputExecutante10.setText(inputs['inputExecutante10']) 
                self.checkBoxErgonomia.setChecked(inputs['checkBoxErgonomia']) 
                self.checkBoxColuna.setChecked(inputs['checkBoxColuna']) 
            
        except Exception as e:
            x = self.emsg.setInformativeText(str(e))
            x = self.emsg.exec_()
            pass 
        
        
if __name__ == '__main__':
    qt = QApplication(sys.argv)
    novo = Novo()
    novo.show()    
    
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(os.path.join(resource_filename(__name__, 'EasyPT.py')))
    
    with open(os.path.join(resource_filename(__name__, 'style.qss')), "r") as f:
        _style = f.read()
        qt.setStyleSheet(_style)
        
    os.system("ie4uinit.exe -show") #reseta os icones do sistema pro nosso aparecer
        
    qt.exec_()
