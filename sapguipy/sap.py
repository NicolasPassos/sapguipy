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
from sapguipy.models.exceptions import ElementNotFound

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
        sets the `logado` attribute to `True` upon successful login.

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
                _current_date = datetime.now()
                _random = randint(0,100)
                _date = _current_date + timedelta(days=_random)
                self.new_pwd = _date.strftime("%B@%Y%H%M%f")

                self.find_by_id("wnd[1]/usr/pwdRSYST-NCODE").set_text(self.new_pwd)
                self.find_by_id("wnd[1]/usr/pwdRSYST-NCOD2").set_text(self.new_pwd)
                self.find_by_id("wnd[1]/tbar[0]/btn[0]").press()

                if self.find_by_id("wnd[1]/usr/txtMESSTXT1"):
                    self.find_by_id("wnd[1]/tbar[0]/btn[0]").press()

            self.logged = True
        else:
            raise Exception('This library only supports Windows OS')
        
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
        sleep(1)
        self.session.startTransaction(transacao)
    
    def get_window_size(self):
        try:
            height = self.find_by_id("wnd[0]").Height
            width = self.find_by_id("wnd[0]").Width
            return width, height
        except Exception:
            raise Exception('Não foi possível obter o tamanho da janela do SAP')
        
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
            case _:
                raise TypeError(f"Element type '{element_type}' is not supported.")
                

class GuiButton:
    def __init__(self, element):
        self.element = element
    
    def press(self):
        """Presses the button."""
        self.element.Press()

class GuiTextField:
    def __init__(self, element):
        self.element = element
    
    def set_text(self, text):
        """Sets the text of the text field."""
        self.element.Text = text
    
    def get_text(self):
        """Returns the text of the text field."""
        return self.element.Text

class GuiComboBox:
    def __init__(self, element):
        self.element = element
    
    def select_entry(self, entry):
        """Selects an entry in the combo box."""
        self.element.Key = entry

class GuiCheckBox:
    def __init__(self, element):
        self.element = element
    
    def select(self, value=True):
        """Sets the value of the CheckBox."""
        self.element.Selected = value

class GuiCTextField:
    def __init__(self, element):
        self.element = element
    
    def set_text(self, text):
        """Sets the text of the CTextField."""
        self.element.Text = text
    
    def get_text(self):
        """Returns the text of the CTextField."""
        return self.element.Text if self.element.Text != '' else None

class GuiTab:
    def __init__(self, element):
        self.element = element
    
    def select(self):
        """Selects the tab."""
        self.element.Select()

class GuiGridView:
    def __init__(self, element):
        self.element = element
    
    def select_row(self, row):
        """Selects a row in the grid view."""
        self.element.Rows.SelectedRow = row
    
    def get_cell_value(self, row, column):
        """Returns the value of a specific cell."""
        return self.element.GetCellValue(row, column)
    
    def set_cell_value(self, row, column, value):
        """Defines the value of a specific cell."""
        self.element.SetCellValue(row, column, value)

class GuiShell:
    def __init__(self, element):
        self.element = element

    @property
    def rows_count(self):
        return self.element.RowCount
    
    @property
    def columns_order(self):
        return [item for item in self.element.ColumnOrder]
    
    def send_command(self, command):
        """Sends a command to the shell."""
        self.element.SendCommand(command)
    
    def get_cell_value(self, row, column):
        """Returns the value of a specific cell."""
        return self.element.GetCellValue(row, column)

    def read_shell_table(self) -> DataFrame:
        """Return a shell table as a pandas DataFrame."""
        columns = self.columns_order
        rows_count = self.rows_count
        data = [
                {column: self.get_cell_value(i, column) for column in columns}
                for i in range(rows_count)
                ]
        
        return DataFrame(data)

class GuiTree:
    def __init__(self, element):
        self.element = element
    
    def expand_node(self):
        """Expands a node in the tree."""
        self.element.Expand()
    
    def collapse_node(self):
        """Collapses a node in the tree."""
        self.element.Collapse()
    
    def select_node(self, node_key):
        """Selects a node in the tree."""
        self.element.SelectNode(node_key)

class GuiStatusbar:
    def __init__(self, element):
        self.element = element
    
    def get_text(self):
        """Returns the text of the status bar."""
        return self.element.Text if self.element.Text != '' else None

class GuiFrameWindow:
    def __init__(self, element):
        self.element = element
    
    def maximize(self):
        """Maximizes the window."""
        self.element.Maximize()
    
    def minimize(self):
        """Minimizes the window."""
        self.element.Iconify()
    
    def restore(self):
        """Restore the original window size."""
        self.element.Restore()

    def close_session(self):
        """Ends the current session."""
        self.session.EndSession()

class GuiSession(SapGui):
    def __init__(self, element):
        self.element = element

    @property
    def info(self):
        return GuiSessionInfo(self.element)
    
    def send_vkey(self, vkey):
        """Sends a virtual key to the session."""
        self.element.SendVKey(vkey)
    
    def end_session(self):
        """Ends the current session."""
        self.element.EndSession()
    
    def start_transaction(self, transaction):
        """Starts a new transaction."""
        self.element.StartTransaction(transaction)
    
    def close_transaction(self):
        """Ends the current transaction."""
        self.element.EndTransaction()

class GuiSessionInfo:
    def __init__(self, element):
        self.element = element
    
    def get_user(self):
        """Returns the user of the session."""
        return self.element.User
    
    def get_client(self):
        """Returns the client of the session."""
        return self.element.Client
    
    def get_transaction(self):
        """Returns the transaction of the session."""
        return self.element.Transaction
    
    def get_program(self):
        """Returns the program of the session."""
        return self.element.Program
    
    def get_system(self):
        """Returns the system of the session."""
        return self.element.System
    
class GuiApplication:
    def __init__(self, element):
        self.element = element
    
    @property
    def name(self):
        """Returns the name of the application."""
        return self.element.Name
    
    @property
    def version(self):
        """Returns the version of the application."""
        return self.element.Version
    
    def open_connection(self, connection_string):
        """Returns the name of the application."""
        return self.element.OpenConnection(connection_string)
    
class GuiLabel:
    def __init__(self, element):
        self.element = element
    @property
    def text(self):
        """Gets the text of the label."""
        return self.element.Text
    
    def get_text(self):
        """Returns the text of the label."""
        return self.element.Text if self.element.Text != '' else None
    
    def set_text(self, text):
        """Sets the text of the label."""
        self.element.Text = text

class GuiToolbar:
    def __init__(self, element):
        self.element = element
    
    def press_button(self, button_id):
        """Pressiona um botão na barra de ferramentas."""
        self.element.PressButton(button_id)

class GuiTableControl:
    def __init__(self, element):
        self.element = element
    
    def set_cell_value(self, row, column, value):
        """Define o valor de uma célula específica no controle de tabela."""
        self.element.SetCellValue(row, column, value)
    
    def get_cell_value(self, row, column):
        """Obtém o valor de uma célula específica no controle de tabela."""
        return self.element.GetCellValue(row, column)

class GuiTitlebar:
    def __init__(self, element):
        self.element = element
    
    @property
    def text(self):
        """Obtém o texto da barra de título."""
        return self.element.Text

class GuiContainer:
    def __init__(self, element):
        self.element = element
    
    def find_by_name(self, name):
        """Encontra um elemento dentro do container pelo nome."""
        return self.element.FindByName(name)

class GuiSplitter:
    def __init__(self, element):
        self.element = element
    
    def set_position(self, position):
        """Define a posição do divisor (splitter)."""
        self.element.SetPosition(position)