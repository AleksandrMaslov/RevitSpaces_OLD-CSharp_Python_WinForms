import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Application, Button, StatusBar, Form, ListView, ImageList, AnchorStyles, ColumnHeaderAutoResizeStyle,
                                  StatusBar, SelectionMode, ListViewItem, View, SortOrder, HorizontalAlignment, TabControl, TabPage, Padding,
                                  FormBorderStyle, Label, GroupBox, ComboBox)
from System.Drawing import Point, Size, Rectangle, Color, Font, FontStyle


class SpacesWindow(Form):
    def __init__(self):
        self._doc = "" # doc
        # self._data = get_spaces_info(doc, BuiltInCategory.OST_MEPSpaces)
        # self._lvls = get_levels_info(doc)
        # self._link_doc = get_link_doc()
        self._initialize_components()

    def _initialize_components(self):
        # window
        self.Text = 'Spaces from Linked Rooms'
        self.form_length = 330
        self.form_width = 220
        self.MinimumSize = Size(self.form_width, self.form_length)
        self.Size = Size(self.form_width, self.form_length)
        self.CenterToScreen()
        self.ShowIcon = False
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow


        # listview
        self.list_view_width = self.form_width - 10
        self.list_column_width_verification = (self.list_view_width - 3) * 0.6
        self.list_column_width_status = (self.list_view_width - 3) * 0.4
        self.list_view_length = 70
        self.list_offset_left = 3
        self.list_offset_top = 3

        self.list_view = ListView()
        self.list_view.Bounds = Rectangle(Point(self.list_offset_left, self.list_offset_top), Size(self.list_view_width, self.list_view_length))
        self.list_view.View = View.Details
        self.list_view.AllowColumnReorder = True
        self.list_view.FullRowSelect = True
        self.list_view.GridLines = True
        self.list_view.Sorting = SortOrder.Ascending
        self.list_view.GridLines = True
        self.list_view.TabIndex = 0
        self.list_view.Columns.Add('Verification', self.list_column_width_verification, HorizontalAlignment.Center)
        self.list_view.Columns.Add('Status', self.list_column_width_status, HorizontalAlignment.Center)
        self.Controls.Add(self.list_view)


        # groupboxes
        self.groupbox_offset_left = 5
        self.groupbox_offset_top = 18
        self.groupbox_width = self.list_view_width - 1
        self.groupbox_current_location_Y = self.list_view_length + 35
        self.groupbox_current_length = 70
        self.groupbox_linked_length = 95
        self.groupbox_linked_location_Y = self.groupbox_current_location_Y + self.groupbox_current_length + 10

        self.groupbox_current = GroupBox()
        self.groupbox_current.Location = Point(self.list_offset_left, self.groupbox_current_location_Y)
        self.groupbox_current.Text = "Current Model Spaces by Phase"
        self.groupbox_current.Size = Size(self.groupbox_width, self.groupbox_current_length)
        self.Controls.Add(self.groupbox_current)

        self.groupbox_linked = GroupBox()
        self.groupbox_linked.Location = Point(self.list_offset_left, self.groupbox_linked_location_Y)
        self.groupbox_linked.Text = "Linked Model"
        self.groupbox_linked.Size = Size(self.groupbox_width, self.groupbox_linked_length)
        self.Controls.Add(self.groupbox_linked)


        # comboboxes
        self.combobox_width = self.groupbox_width - self.groupbox_offset_left * 2

        self.combobox_phase = ComboBox()
        self.combobox_phase.Parent = self.groupbox_current
        self.combobox_phase.Location = Point(self.groupbox_offset_left, self.groupbox_offset_top)
        self.combobox_phase.Text = " - Select Phase - "
        self.combobox_phase.Size = Size(self.combobox_width, 0)

        self.combobox_link = ComboBox()
        self.combobox_link.Parent = self.groupbox_linked
        self.combobox_link.Location = Point(self.groupbox_offset_left, self.groupbox_offset_top)
        self.combobox_link.Text = " - Select Link - "
        self.combobox_link.Size = Size(self.combobox_width, 0)

        self.combobox_link_phase = ComboBox()
        self.combobox_link_phase.Parent = self.groupbox_linked
        self.combobox_link_phase.Location = Point(self.groupbox_offset_left, self.groupbox_offset_top + 25)
        self.combobox_link_phase.Text = " - Select Phase - "
        self.combobox_link_phase.Size = Size(self.combobox_width, 0)


        # buttons
        self.button_width = 50 
        self.button_length = 20
        self.button_width_large = 95 
        self.button_offset = 22
        self.button_location_current_Y = self.groupbox_current_length - self.button_length - 6
        self.button_location_linked_Y = self.groupbox_linked_length - self.button_length - 6

        btn_report = Button()
        btn_report.Text = 'Report'
        btn_report.Parent = self
        btn_report.Size = Size(self.button_width, self.button_length)
        btn_report.Location = Point(self.list_offset_left, self.list_view_length + 5)
        # btn_report.Click += self._is_click_btn1
    
        btn_help = Button()
        btn_help.Text = 'Help'
        btn_help.Parent = self
        btn_help.Size = Size(self.button_width, self.button_length)
        btn_help.Location = Point(self.list_offset_left + self.button_width + 2, self.list_view_length + 5)
        # btn_report.Click += self._is_click_btn1

        btn_delete_all = Button()
        btn_delete_all.Text = 'Delete All'
        btn_delete_all.Size = Size(self.button_width_large, self.button_length)
        btn_delete_all.Location = Point(self.groupbox_offset_left, self.button_location_current_Y)
        btn_delete_all.Parent = self.groupbox_current
        # btn_report.Click += self._is_click_btn1

        btn_delete_selected = Button()
        btn_delete_selected.Text = 'Delete Selected'
        btn_delete_selected.Size = Size(self.button_width_large, self.button_length)
        btn_delete_selected.Location = Point(self.groupbox_width - self.button_width_large - self.groupbox_offset_left - 1, self.button_location_current_Y)
        btn_delete_selected.Parent = self.groupbox_current
        # btn_report.Click += self._is_click_btn1

        btn_create_all = Button()
        btn_create_all.Text = 'Create All'
        btn_create_all.Size = Size(self.button_width_large, self.button_length)
        btn_create_all.Location = Point(self.groupbox_offset_left, self.button_location_linked_Y)
        btn_create_all.Parent = self.groupbox_linked
        # btn_report.Click += self._is_click_btn1

        btn_create_selected = Button()
        btn_create_selected.Text = 'Create Selected'
        btn_create_selected.Size = Size(self.button_width_large, self.button_length)
        btn_create_selected.Location = Point(self.groupbox_width - self.button_width_large - self.groupbox_offset_left - 1, self.button_location_linked_Y)
        btn_create_selected.Parent = self.groupbox_linked
        # btn_report.Click += self._is_click_btn1


        #satusbar
        self.statusbar = StatusBar()
        self.statusbar.Parent = self

        self.Load += self._load_window

    def _load_window(self, sender, e):
        self._fill_list_view()

    def _fill_list_view(self):
        row_names = ['Phase Matching', 'Levels Matching', 'Workset Model Spaces']
        statuses = ['*VERIFIED*', '*VERIFIED*', '*VERIFIED*']
        for row_name in row_names:
            item = ListViewItem(row_name)
            item.UseItemStyleForSubItems = True
            item.Font = Font("Arial", 8.6, FontStyle.Regular)
            self.list_view.Items.Add(item)
            for status in statuses:
                item.SubItems.Add(status)

        # self._list_view.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        # self.list_view.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)


main_window = SpacesWindow()
main_window.ShowDialog()
