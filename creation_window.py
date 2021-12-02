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
# from lite_logging import Logger

# logger = Logger(parent_folders_path=os.path.join('Synergy Systems', 'Create Spaces From Linked Rooms'),
#                 file_name='test_log',
#                 default_status=Logger.WARNING)


class CreationWindow(Form):
    def __init__(self, doc, rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms):
        self.doc = doc
        self._initialize_components()

    def _initialize_components(self):
        # window
        self.Text = 'Spaces Creation'
        self.form_length = 300
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
        # self.list_view.Columns.Add('Verification', self.list_column_width_verification, HorizontalAlignment.Center)
        # self.list_view.Columns.Add('Status', self.list_column_width_status, HorizontalAlignment.Center)
        self.Controls.Add(self.list_view)

        # # buttons
        # self.button_width = 50 
        # self.button_length = 20
        # self.button_width_large = 95 
        # self.button_offset = 22
        # self.button_location_current_Y = self.groupbox_current_length - self.button_length - 6
        # self.button_location_linked_Y = self.groupbox_linked_length - self.button_length - 6

        # btn_errors = Button()
        # btn_errors.Text = 'Errors'
        # btn_errors.Parent = self
        # btn_errors.Size = Size(self.button_width, self.button_length)
        # btn_errors.Location = Point(self.list_offset_left, self.list_view_length + 5)
        # # btn_report.Click += self._is_click_btn1
    
        #satusbar
        self.statusbar = StatusBar()
        self.statusbar.Parent = self

        # self.Load += self._load_window

    def _load_window(self, sender, e):
        self._fill_list_view()
        self._fill_combobox_phase()
        self._fill_combobox_link()

    def _click_btn_delete_all(self, sender, e):
        if len(self.combobox_phase.Items) == 0:
            TaskDialog.Show('Warning', '\nNo Spaces in the current model')
            return
        with Transaction(doc) as t:
            t.Start('Delete All Spaces')
            deleleted_counter, phases_counter, phases_list = self._delete_all_spaces()
            t.Commit()
            if t.GetStatus() == TransactionStatus.Committed:
                self.spaces_by_phase_dct = {}
                self.combobox_phase.Items.Clear()

                msg = '\nTotal "{}" Spaces have been deleted\nin "{}" model Phases:\n\n{}'.format(deleleted_counter, phases_counter, phases_list)
                # logger.write_log('Spaces Deleted: {}'.format(deleleted_counter), Logger.INFO)
                TaskDialog.Show('Report', msg)


if __name__ == '__main__':
    information_window = CreationWindow()
    information_window.ShowDialog()
