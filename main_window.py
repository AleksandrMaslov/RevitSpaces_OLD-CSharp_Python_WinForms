import clr
import os
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPIUI')
from System.Windows.Forms import (Button, StatusBar, Form, ListView, StatusBar, ListViewItem, View, SortOrder, HorizontalAlignment,
                                  FormBorderStyle, GroupBox, ComboBox)
from System.Drawing import Point, Size, Rectangle
from Autodesk.Revit.DB import Transaction, TransactionStatus
from Autodesk.Revit.UI import TaskDialog
from lite_logging import Logger
# from main import Main

logger = Logger(parent_folders_path=os.path.join('Synergy Systems', 'Create Spaces From Linked Rooms'),
                file_name='test_log',
                default_status=Logger.WARNING)
doc = __revit__.ActiveUIDocument.Document
transaction = Transaction(doc)


class MainWindow(Form):
    def __init__(self, doc, current_spaces_by_phase, rooms_by_link_and_phase):
        self.doc = doc
        self.spaces_by_phase_dct = current_spaces_by_phase
        self.rooms_by_link_and_phase_dct = rooms_by_link_and_phase
        # self._data = get_spaces_info(doc, BuiltInCategory.OST_MEPSpaces)
        # self._lvls = get_levels_info(doc)
        # self._link_doc = get_link_doc()
        self._initialize_components()

    def _initialize_components(self):
        # window
        self.Text = 'Spaces Manager'
        self.form_length = 430
        self.form_width = 320
        self.MinimumSize = Size(self.form_width, self.form_length)
        self.Size = Size(self.form_width, self.form_length)
        self.CenterToScreen()
        self.ShowIcon = False
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow


        # listview
        self.list_view_width = self.form_width - 35
        self.list_column_width_verification = (self.list_view_width - 3) * 0.6
        self.list_column_width_status = (self.list_view_width - 3) * 0.4
        self.list_view_length = 110
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

        btn_errors = Button()
        btn_errors.Text = 'Errors'
        btn_errors.Parent = self
        btn_errors.Size = Size(self.button_width, self.button_length)
        btn_errors.Location = Point(self.list_offset_left, self.list_view_length + 5)
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
        btn_delete_selected.Click += self._click_btn_delete_selected

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
        self._fill_combobox_phase()
        self._fill_combobox_link()

    def _click_btn_delete_selected(self, sender, e):
        selected_item = self.combobox_phase.SelectedItem
        if selected_item:
            phase_name = selected_item.split(' - ')[1]
            with Transaction(doc) as t:
                deleted_counter = self._delete_selected_spaces(t, phase_name)
                if t.GetStatus() == TransactionStatus.Committed:
                    self.spaces_by_phase_dct.pop(phase_name)
                    self.combobox_phase.Items.Remove(selected_item)

                    msg = '\nTotal "{}" Spaces have been deleted\nin "{}" model Phase'.format(deleted_counter, phase_name)
                    logger.write_log('Spaces Deleted: {}'.format(deleted_counter), Logger.INFO)
                    TaskDialog.Show('Report', msg)
        else:
            TaskDialog.Show('Error', '\nPhase is not selected in the current model.')

    def _fill_list_view(self):
        row_names = ['Phase Matching', 'Levels Matching', 'Workset Model Spaces']
        statuses = ['Not Verified', 'VERIFIED', 'Verified']
        for index, row_name in enumerate(row_names):
            item = ListViewItem(row_name)
            item.UseItemStyleForSubItems = True
            self.list_view.Items.Add(item)
            item.SubItems.Add(statuses[index])

    def _fill_combobox_phase(self):
        for phase_name, spaces in self.spaces_by_phase_dct.items():
            spaces_number = len(spaces)
            item = '{} Spaces - {}'.format(spaces_number, phase_name)
            self.combobox_phase.Items.Add(item)

    def _fill_combobox_link(self):
        for link_name, phases in self.rooms_by_link_and_phase_dct.items():
            rooms_number_total = 0
            for rooms in phases:
                rooms_number_phase = len(rooms)
                rooms_number_total += rooms_number_phase
            item = '{} Rooms - {}'.format(rooms_number_total, link_name)
            self.combobox_link.Items.Add(item)

    def _delete_selected_spaces(self, t, phase_name):
        deleted_counter = 0
        spaces = self.spaces_by_phase_dct[phase_name]
        t.Start('Delete Spaces in "{}" Phase'.format(phase_name))
        for space in spaces.values():
            self.doc.Delete(space.Id)
            deleted_counter += 1 
        t.Commit()
        return deleted_counter 


if __name__ == '__main__':
    main_window = MainWindow()
main_window.ShowDialog()
