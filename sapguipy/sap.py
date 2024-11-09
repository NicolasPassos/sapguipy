from datetime import datetime, timedelta
from pythoncom import CoInitialize
from pywinauto import Application
from subprocess import Popen
from threading import Thread
from pandas import DataFrame
from platform import system
from random import randint
from pathlib import Path
import pygetwindow as gw
from time import sleep
import win32com.client
import psutil

class SapGui:
    def __init__(self, sid: str, user: str, pwd: str, mandante: str, language:str='PT', root_sap_dir: str='C:\Program Files (x86)\SAP\FrontEnd\SAPGUI'):
        '''
        sid: identificador do sistema, cada ambiente tem seu próprio SID. Normalmente são: PRD (produção) / DEV (desenvolvimento) / QAS (qualidade).
        usuario: usuário que a automação utilizará para realizar login.
        senha: senha que a automação utilizará para realizar login.
        diretorio_instalacao: diretório onde o sapshcut.exe se encontra, onde foi realizado a instalação do SAP.
        '''
        self.sid = sid
        self.user = user
        self.__pwd = pwd
        self.language = language
        self.mandante = mandante
        self.root_sap_dir = Path(root_sap_dir)
        self.new_pwd = None
        self.logado = False
        self.id_scheduling = None

    def __enter__(self):
        self.start_sap()
        self.thread_conexao = Thread(target=self.verify_sap_connection,daemon=True)
        self.thread_conexao.start()
        return self
    
    def __exit__(self, exc_type, exc_value, traceback):
        if self.logado:
            self.quit()
        
    def start_sap(self):
        if system() == 'Windows':
            tentativas = 0
            while tentativas <=5:
                try:
                    self.program = Popen(args=f'{self.root_sap_dir}/sapshcut.exe -system={self.sid} -client={self.mandante} -user={self.user} -pw={self.__pwd} -language={self.language}')
                    sleep(1)
                    self.SapGuiAuto = win32com.client.GetObject('SAPGUI',CoInitialize())
                    if self.SapGuiAuto:
                        break
                except Exception as e:
                    if tentativas >= 5:
                        raise Exception(f'Failed to get control of SAP: {e}.')
                    tentativas += 1

                    sleep(3)

            # Get the SAP Application object
            self.application = self.SapGuiAuto.GetScriptingEngine
            if not self.application:
                print("Failed to get SAP Application object.")
                exit()

            sleep(3)
            self.connection = self.application.Children(0)
            self.session = self.connection.Children(0)

            # Se aparecer a janela de x tentativas com a senha incorreta
            if self.session.findById("wnd[1]/usr/txtMESSTXT1",False):
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

            if self.session.Info.User == '':
                if self.session.findById("wnd[0]/sbar/pane[0]").text == 'O nome ou a senha não está correto (repetir o logon)':
                    raise ValueError('Failed to login with the provided credentials.')
                
                if self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1",False):
                    self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            if self.session.findById("wnd[1]/usr/lblRSYST-NCODE_TEXT",False):
                # Se aparecer a tela de troca de senha, resetar a senha
                _data_atual = datetime.now()
                _random = randint(0,100)
                _data = _data_atual + timedelta(days=_random)
                self.new_pwd = _data.strftime("%B@%Y%H%M%f")

                self.session.findById("wnd[1]/usr/pwdRSYST-NCODE").text = self.new_pwd
                self.session.findById("wnd[1]/usr/pwdRSYST-NCOD2").text = self.new_pwd
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                if self.session.findById("wnd[1]/usr/txtMESSTXT1",False):
                    self.findById("wnd[1]/tbar[0]/btn[0]").press()

            self.logado = True
            self.session.findById("wnd[0]").maximize()
        else:
            raise Exception('This library only supports Windows OS')


    def login(self):
        self.session.findById("wnd[0]").maximize()
        self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.user
        self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.__pwd
        self.session.findById("wnd[0]").sendVKey(0)

        if self.session.Info.User == '':
            if self.session.findById("wnd[0]/sbar/pane[0]").text == 'O nome ou a senha não está correto (repetir o logon)':
                raise ValueError('Failed to login with the provided credentials.')
            
            self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

    def logoff(self):
        self.session.findById("wnd[0]").maximize()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nend"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

    def quit(self):
        self.program.terminate()

        for proc in psutil.process_iter(['pid', 'name']):
            if 'saplogon' in proc.info['name'].lower():
                proc.kill()

        self.logado = False

    def open_transaction(self,transacao: str):
        self.session.startTransaction(transacao)
    
    def get_window_size(self):
        try:
            height = self.session.findById("wnd[0]").Height
            width = self.session.findById("wnd[0]").Width
            return width, height
        except Exception:
            raise Exception('Não foi possível obter o tamanho da janela do SAP')
        
    def verify_sap_connection(self):
        while self.logado:
            sleep(10)
            #Se a janela com o titulo 'SAP GUI for Windows 800' está aparecendo, significa que o SAP caiu. Então a execução deve ser interrompida.
            windows = gw.getWindowsWithTitle('SAP GUI for Windows 800')
            if windows:
                # Fechar a janela
                for window in windows:
                    try:
                        app = Application().connect(handle=window._hWnd)
                        app_window = app.window(handle=window._hWnd)
                        app_window.close()
                        self.logado = False
                    except Exception as e:
                        raise Exception(f'Não foi possível fechar a janela do SAP: {e}')
        else:
            return
        
    def read_shell_table(self, id_tabela: str):
        try:
            shell = self.session.findById(id_tabela)

            columns = shell.ColumnOrder
            rows_count = shell.RowCount
        except Exception as e:
            raise Exception(
                f"Error finding table {id_tabela}: {e}"
            )
        data = [
                {column: shell.GetCellValue(i, column) for column in columns}
                for i in range(rows_count)
                ]
        
        return DataFrame(data)