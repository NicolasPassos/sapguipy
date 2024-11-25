from datetime import datetime, timedelta
from pythoncom import CoInitialize
from pywinauto import Application
from subprocess import Popen
from threading import Thread
from platform import system
from random import randint
from pathlib import Path
import pygetwindow as gw
from time import sleep
import win32com.client
import psutil
from sapguipy.models.exceptions import ElementNotFound
from sapguipy.models.sap_controls import *

class SapGui:
    def __init__(self, sid: str, user: str, pwd: str, mandante: str, root_sap_dir: str='C:\Program Files (x86)\SAP\FrontEnd\SAPGUI'):
        """
        sid: identificador do sistema, cada ambiente tem seu próprio SID. Normalmente são: PRD (produção) / DEV (desenvolvimento) / QAS (qualidade).
        usuario: usuário que a automação utilizará para realizar login.
        senha: senha que a automação utilizará para realizar login.
        diretorio_instalacao: diretório onde o sapshcut.exe se encontra, onde foi realizado a instalação do SAP.
        """
        self.sid = sid
        self.user = user
        self.__pwd = pwd
        self.language = 'PT'
        self.mandante = mandante
        self.root_sap_dir = Path(root_sap_dir)
        self.new_pwd = None
        self.logged = False
        self.session_info = None
        self.statusbar = None

    def __enter__(self):
        self.start_sap()
        self.thread_conexao = Thread(target=self.verify_sap_connection,daemon=True)
        self.thread_conexao.start()
        return self
    
    def __exit__(self, exc_type, exc_value, traceback):
        if self.logged:
            self.quit()
        
    def start_sap(self):
        """
        Starts the SAP application and establishes a session.

        This method attempts to launch the SAP application on a Windows OS and 
        establish a session by logging in with the provided credentials. It handles 
        multiple login attempts and manages password changes if required. The method 
        sets the `logged` attribute to `True` upon successful login.

        Raises:
            Exception: If the SAP application fails to start after multiple attempts 
            or if the SAP Application object cannot be obtained.
            ValueError: If the login attempt fails due to incorrect credentials.
        """
        if system() == 'Windows':
            tries = 0
            while tries <=5:
                try:
                    self.program = Popen(args=f'{self.root_sap_dir}/sapshcut.exe -system={self.sid} -client={self.mandante} -user={self.user} -pw={self.__pwd} -language={self.language}')
                    sleep(1)
                    self.SapGuiAuto = win32com.client.GetObject('SAPGUI',CoInitialize())
                    if self.SapGuiAuto:
                        break
                except Exception as e:
                    if tries >= 5:
                        raise Exception(f'Failed to get control of SAP: {e}.')
                    tries += 1

                    sleep(3)

            # Get the SAP Application object
            self.application = self.SapGuiAuto.GetScriptingEngine
            if not self.application:
                raise Exception("Failed to get SAP Application object.")

            sleep(2)
            self.connection = self.application.Children(0)
            self.session = self.connection.Children(0)
            self.session_info = GuiSessionInfo(self.session.info)
            self.statusbar = self.find_by_id("wnd[0]/sbar")

            # Se aparecer a janela de x tentativas com a senha incorreta
            if self.find_by_id("wnd[1]/usr/txtMESSTXT1", False):
                self.find_by_id("wnd[1]/tbar[0]/btn[0]").press()

            if self.session_info.get_user() is None:
                if self.statusbar.get_text() == 'O nome ou a senha não está correto (repetir o logon)':
                    raise ValueError('Failed to login with the provided credentials.')
                
                if self.find_by_id("wnd[1]/usr/radMULTI_LOGON_OPT1", False):
                    self.find_by_id("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
                    self.find_by_id("wnd[1]/tbar[0]/btn[0]").press()
            
            if self.find_by_id("wnd[1]/usr/lblRSYST-NCODE_TEXT", False):
                # Se aparecer a tela de troca de senha, resetar a senha
                self.change_password()
            self.logged = True
        else:
            raise Exception('This library only supports Windows OS')
        
    def change_password(self):
        """
        If the password is expired, SAP will open an modal to change the password.
        This method changes the password in this modal.
        """
        _current_date = datetime.now()
        _random = randint(0,100)
        _date = _current_date + timedelta(days=_random)
        self.new_pwd = _date.strftime("%B@%Y%H%M%f")

        self.find_by_id("wnd[1]/usr/pwdRSYST-NCODE").set_text(self.new_pwd)
        self.find_by_id("wnd[1]/usr/pwdRSYST-NCOD2").set_text(self.new_pwd)
        self.find_by_id("wnd[1]/tbar[0]/btn[0]").press()

        if self.find_by_id("wnd[1]/usr/txtMESSTXT1", False):
            self.find_by_id("wnd[1]/tbar[0]/btn[0]").press()
        
    def get_user_logged(self):
        
        """
        Returns the user currently logged in.
        """
        user = self.session_info.get_user()
        return None if user == '' else user
    
    def new_window(self):
        if self.connection.Sessions.Count < 3:
            self.session.CreateSession()
            sleep(2)
            new_session = self.connection.Children(self.connection.Sessions.Count-1)
            return SapGui(
                        sid=self.sid,
                        user=self.user,
                        pwd=self.__pwd,
                        mandante=self.mandante,
                        root_sap_dir=str(self.root_sap_dir)
                            )._initialize_new_session(new_session)
        else:
            raise Exception('Maximum number of windows reached.')
        
    def _initialize_new_session(self, session):
        """
        Initializes a new session within the SAP application and returns a new SapGui object.
        """
        self.session = session
        self.session_info = GuiSessionInfo(self.session.info)
        self.statusbar = self.find_by_id("wnd[0]/sbar")
        self.logged = True
        return self
    
    def login(self):
        """
        Logins into the SAP application using the provided credentials.
        """
        self.find_by_id("wnd[0]").maximize()
        self.find_by_id("wnd[0]/usr/txtRSYST-BNAME").set_text(self.user)
        self.find_by_id("wnd[0]/usr/pwdRSYST-BCODE").set_text(self.__pwd)
        self.find_by_id("wnd[0]").sendVKey(0)

        if self.get_user_logged() is None:
            if self.find_by_id("wnd[0]/sbar/pane[0]").text == 'O nome ou a senha não está correto (repetir o logon)':
                raise ValueError('Failed to login with the provided credentials.')
            
            self.find_by_id("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
            self.find_by_id("wnd[1]/tbar[0]/btn[0]").press()

    def logoff(self):
        """
        Logs out of the SAP application.
        """
        self.find_by_id("wnd[0]").maximize()
        self.find_by_id("wnd[0]/tbar[0]/okcd").set_text("/nend")
        self.find_by_id("wnd[0]").sendVKey(0)
        self.find_by_id("wnd[1]/usr/btnSPOP-OPTION1").press()

    def quit(self):
        """
        Forces the SAP application to quit.
        """
        self.program.terminate()

        for proc in psutil.process_iter(['pid', 'name']):
            if 'saplogon' in proc.info['name'].lower():
                proc.kill()

        self.logged = False

    def open_transaction(self,transacao: str):
        """Open an SAP transaction."""
        self.session.startTransaction(transacao)
    
    def get_window_size(self):
        return self.find_by_id("wnd[0]").width, self.find_by_id("wnd[0]").height
        
    def verify_sap_connection(self):
        while self.logged:
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
                        self.logged = False
                    except Exception as e:
                        raise Exception(f'Não foi possível fechar a janela do SAP: {e}')
        else:
            return
        
    def find_by_id(self, element_id: str, raise_error: bool = True):
        """
        Returns a instance of the GuiElement class supplied with the specified ID.
        """
        element = self.session.FindById(element_id, False)

        if element is None and raise_error:
            raise ElementNotFound(f"The element with ID '{element_id}' was not found.")
        elif element is None and not raise_error:
            return None
        
        element_type = element.Type
        
        match element_type:
            case "GuiButton":
                return GuiButton(element)
            case "GuiTextField":
                return GuiTextField(element)
            case "GuiComboBox":
                return GuiComboBox(element)
            case "GuiCheckBox":
                return GuiCheckBox(element)
            case "GuiCTextField":
                return GuiCTextField(element)
            case "GuiTab":
                return GuiTab(element)
            case "GuiGridView":
                return GuiGridView(element)
            case "GuiShell":
                return GuiShell(element)
            case "GuiTree":
                return GuiTree(element)
            case "GuiStatusbar":
                return GuiStatusbar(element)
            case "GuiFrameWindow":
                return GuiFrameWindow(element)
            case "GuiSessionInfo":
                return GuiSessionInfo(element)
            case "GuiLabel":
                return GuiLabel(element)
            case "GuiToolbar":
                return GuiToolbar(element)
            case "GuiTableControl":
                return GuiTableControl(element)
            case "GuiTitlebar":
                return GuiTitlebar(element)
            case "GuiContainer":
                return GuiContainer(element)
            case "GuiSplitter":
                return GuiSplitter(element)
            case "GuiUserArea":
                return GuiUserArea(element)
            case "GuiMainWindow":
                return GuiMainWindow(element)
            case _:
                raise TypeError(f"Element type '{element_type}' is not supported.")