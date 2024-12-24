from pandas import DataFrame
class GuiVComponent:
    def __init__(self, element):
        self.element = element

    @property
    def id(self):
        return self.element.Id
    
    @property
    def name(self):
        return self.element.Name
    @property
    def changeable(self):
        if not hasattr(self.element, 'changeable'):
            return False
        return self.element.Changeable
        
    
    @property
    def default_tooltip(self):
        if not hasattr(self.element, 'defaulttooltip'):
            return None
        return self.element.DefaultTooltip if self.element.DefaultTooltip != '' else None
    
    @property
    def height(self):
        return self.element.Height
    
    @property
    def icon_name(self):
        if not hasattr(self.element, 'iconname'):
            return None
        return self.element.IconName if self.element.IconName != '' else None
    
    @property
    def left(self):
        return self.element.Left
    
    @property
    def modified(self):
        if not hasattr(self.element, 'modified'):
            return None
        return self.element.Modified if self.element.ContainerType else None
    
    @property
    def screen_left(self):
        return self.element.ScreenLeft
    
    @property
    def screen_top(self):
        return self.element.ScreenTop
    
    @property
    def text(self):
        return self.element.Text if self.element.Text != '' else None
    
    @property
    def tooltip(self):
        if not hasattr(self.element, 'tooltip'):
            return None
        return self.element.Tooltip if self.element.Tooltip != '' else None
    
    @property
    def top(self):
        return self.element.Top
    
    @property
    def width(self):
        return self.element.Width
    
    @property
    def key(self):
        if not hasattr(self.element, 'key'):
            return None
        return self.element.Key
        
    @property
    def has_children(self):
        if hasattr(self.element, 'children'):
            return True
        else:
            return False
        
class GuiButton(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    def press(self):
        """Presses the button."""
        self.element.Press()

class GuiTextField(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    def set_text(self, text):
        """Sets the text of the text field."""
        self.element.Text = text
    
    def get_text(self):
        """Returns the text of the text field."""
        return self.element.Text

class GuiComboBox(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    def select_entry(self, entry):
        """Selects an entry in the combo box."""
        self.element.Key = entry

class GuiCheckBox(GuiVComponent):
    def __init__(self, element):
        self.element = element

    @property
    def selected(self):
        return self.element.Selected
    
    def select(self, value=True):
        """Sets the value of the CheckBox."""
        self.element.Selected = value

class GuiCTextField(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    def set_text(self, text):
        """Sets the text of the CTextField."""
        self.element.Text = text
    
    def get_text(self):
        """Returns the text of the CTextField."""
        return self.element.Text if self.element.Text != '' else None

class GuiTab(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    def select(self):
        """Selects the tab."""
        self.element.Select()

class GuiGridView(GuiVComponent):
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

class GuiShell(GuiVComponent):
    def __init__(self, element):
        self.element = element

    @property
    def rows_count(self):
        if not hasattr(self.element, 'rowcount'):
            return None
        return self.element.RowCount
    
    @property
    def columns_order(self):
        return [item for item in self.element.ColumnOrder]
    
    @property
    def drag_drop_supported(self):
        return self.element.DragDropSupported
    
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

class GuiTree(GuiVComponent):
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

class GuiStatusbar(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    @property
    def message_type(self):
        match self.element.MessageType:
            case "S":
                return "Success"
            case "E":
                return "Error"
            case "W":
                return "Warning"
            case "I":
                return "Information"
            case "A":
                return "Abort"
            case _:
                return None
    
    @property
    def has_popup(self):
        return self.element.MessageAsPopup
    
    def get_text(self):
        """Returns the text of the status bar."""
        return self.element.Text if self.element.Text != '' else None

class GuiFrameWindow(GuiVComponent):
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

class GuiSessionInfo(GuiVComponent):
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
    
class GuiApplication(GuiVComponent):
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
    
class GuiLabel(GuiVComponent):
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

class GuiToolbar(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    def press_button(self, button_id):
        """Presses a button in the toolbar."""
        self.element.PressButton(button_id)

class GuiTableControl(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    def set_cell_value(self, row, column, value):
        """Sets the value of a specific cell."""
        self.element.SetCellValue(row, column, value)
    
    def get_cell_value(self, row, column):
        """Returns the value of a specific cell."""
        return self.element.GetCellValue(row, column)

class GuiTitlebar(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    @property
    def text(self):
        """Gets the text of the title bar."""
        return self.element.Text

class GuiContainer(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    def find_by_name(self, name):
        """Finds a control by name."""
        return self.element.FindByName(name)

class GuiSplitter(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    @property
    def element(self):
        return self.element

class GuiUserArea(GuiVComponent):
    def __init__(self, class_instance, element):
        self.__class = class_instance
        self.element = element
    
    def list_children(self):
        """List all children of the user area."""
        return [self.__class.find_by_id(item) for item in self.element.Children]
    
class GuiMainWindow(GuiVComponent):
    def __init__(self, class_instance, element):
        self.__class = class_instance
        self.element = element
    
    @property
    def v_key_allowed(self):
        return self.element.IsVKeyAllowed

    def maximize(self):
        """Maximizes the window."""
        self.element.Maximize()
    
    def minimize(self):
        """Minimizes the window."""
        self.element.Iconify()
    
    def restore(self):
        """Restore the original window size."""
        self.element.Restore()

    def send_v_key(self, key):
        """Sends a VKey to the window."""
        self.element.SendVKey(key)

    def show_message_box(self):
        self.element.ShowMessageBox()

    def tab_backward(self):
        self.element.TabBackward()

    def tab_forward(self):
        self.element.TabForward()
    
    def close(self):
        self.element.Close()

    def list_children(self):
        """List all children."""
        return [self.__class.find_by_id(item) for item in self.element.Children]

class GuiComponentCollection(GuiVComponent):
    def __init__(self, element):
        self.element = element
    
    @property
    def count(self):
        return self.element.Count
    
    @property
    def Length(self):
        return self.element.Length
    
class GuiMenubar(GuiVComponent):
    def __init__(self, class_instance, element):
        self.__class = class_instance
        self.element = element
    
    def list_children(self):
        """List all children."""
        return [self.__class.find_by_id(item) for item in self.element.Children]
    
class GuiCustomControl(GuiVComponent):
    def __init__(self, class_instance, element):
        self.__class = class_instance
        self.element = element

    def list_children(self):
        """List all children."""
        return [self.__class.find_by_id(item) for item in self.element.Children]
    
class GuiContainerShell(GuiVComponent):
    def __init__(self, class_instance, element):
        self.__class = class_instance
        self.element = element

    
    
    def list_children(self):
        """List all children."""
        return [self.__class.find_by_id(item) for item in self.element.Children]