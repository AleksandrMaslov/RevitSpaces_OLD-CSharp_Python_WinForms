import clr
import os
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Button, StatusBar, Form, StatusBar, FormBorderStyle, GroupBox, ComboBox, Label)
from System.Drawing import Point, Size
from Autodesk.Revit.DB import Transaction, TransactionStatus
from lite_logging import Logger
from creation_window import CreationWindow
from information_window import InformationWindow

logger = Logger(parent_folders_path=os.path.join('Synergy Systems', 'Create Spaces From Linked Rooms'),
                file_name='test_log',
                default_status=Logger.WARNING)


class MainWindow(Form):
    def __init__(self, doc, workset_spaces_id, current_spaces_by_phase, rooms_by_link_and_phase, current_levels, active_view_phase):
        self.doc = doc
        self.spaces_by_phase_dct = current_spaces_by_phase
        self.rooms_by_link_and_phase_dct = rooms_by_link_and_phase
        self.workset_spaces_id = workset_spaces_id
        self.current_levels = current_levels
        self.active_view_phase = active_view_phase
        self._initialize_components()

    def _initialize_components(self):
        # window
        self.Text = 'Spaces Manager'
        self.form_width = 320
        self.form_length = 400
        self.MinimumSize = Size(self.form_width, self.form_length)
        self.Size = Size(self.form_width, self.form_length)
        self.CenterToScreen()
        self.ShowIcon = False
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow

        # groupboxes
        self.form_offset_left = 9
        self.groupbox_offset_left = 5
        self.groupbox_offset_top = 18
        self.groupbox_width = self.form_width - 36
        self.groupbox_current_location_Y = 15
        self.groupbox_current_length = 80
        self.groupbox_linked_length = 190
        self.groupbox_linked_location_Y = self.groupbox_current_location_Y + self.groupbox_current_length + 10

        self.groupbox_current = GroupBox()
        self.groupbox_current.Location = Point(self.form_offset_left, self.groupbox_current_location_Y)
        self.groupbox_current.Text = "Current Model Spaces by Phase"
        self.groupbox_current.Size = Size(self.groupbox_width, self.groupbox_current_length)
        self.Controls.Add(self.groupbox_current)

        self.groupbox_linked = GroupBox()
        self.groupbox_linked.Location = Point(self.form_offset_left, self.groupbox_linked_location_Y)
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
        self.combobox_link.SelectedIndexChanged += self._changed_combobox_link_selection

        self.combobox_link_phase = ComboBox()
        self.combobox_link_phase.Parent = self.groupbox_linked
        self.combobox_link_phase.Location = Point(self.groupbox_offset_left, self.groupbox_offset_top + 25)
        self.combobox_link_phase.Text = " - Select Phase - "
        self.combobox_link_phase.Size = Size(self.combobox_width, 0)

        # buttons
        self.button_width = 60 
        self.button_length = 25
        self.button_width_large = 120 
        self.button_offset = 22
        self.button_location_current_Y = self.groupbox_current_length - self.button_length - 8
        self.button_location_linked_Y = self.groupbox_offset_top + 33 + self.button_length
    
        btn_help = Button()
        btn_help.Text = 'Help'
        btn_help.Parent = self
        btn_help.Size = Size(self.button_width, self.button_length)
        btn_help.Location = Point(self.form_offset_left, self.groupbox_linked_location_Y + self.groupbox_linked_length + 5)
        # btn_report.Click += self._is_click_btn1

        btn_delete_all = Button()
        btn_delete_all.Text = 'Delete All'
        btn_delete_all.Size = Size(self.button_width_large, self.button_length)
        btn_delete_all.Location = Point(self.groupbox_offset_left, self.button_location_current_Y)
        btn_delete_all.Parent = self.groupbox_current
        btn_delete_all.Click += self._click_btn_delete_all

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
        btn_create_selected.Click += self._click_btn_create_selected

        # label
        self.label_location_linked_Y = self.button_location_linked_Y + self.button_length + 5
        self.label_width = self.groupbox_width - self.groupbox_offset_left * 2

        label_current_phase = Label()
        label_current_phase.Parent = self.groupbox_linked
        label_current_phase.Location = Point(self.groupbox_offset_left, self.label_location_linked_Y)
        label_current_phase.Size = Size(self.label_width, 75)
        label_current_phase.Text = 'New spaces Phase:\n{}\n\n* If you need to change the Phase close the addin and open definite view in your model.'.format(self.active_view_phase)
        # self._label_current_phase.Font = Font("Arial", 10, FontStyle.Regular)

        #satusbar
        self.statusbar = StatusBar()
        self.statusbar.Parent = self

        self.Load += self._load_window

    def _load_window(self, sender, e):
        self._fill_combobox_phase()
        self._fill_combobox_link()

    def _click_btn_delete_all(self, sender, e):
        if len(self.combobox_phase.Items) == 0:
            message = 'There are no Spaces in the Current model.'
            information_window = InformationWindow('Information', message)
            information_window.ShowDialog() 
            return
        with Transaction(self.doc) as t:
            t.Start('Delete All Spaces')
            deleleted_counter, phases_counter, phases_list = self._delete_all_spaces()
            t.Commit()
            if t.GetStatus() == TransactionStatus.Committed:
                self.spaces_by_phase_dct = {}
                self.combobox_phase.Items.Clear()

                logger.write_log('{} Spaces have been deleted\nin {} model Phases:\n{}'.format(deleleted_counter, phases_counter, phases_list), Logger.INFO)
                message = 'Total {} Spaces have been deleted\nin {} model Phases:\n\n{}'.format(deleleted_counter, phases_counter, phases_list)
                information_window = InformationWindow('Report', message)
                information_window.ShowDialog() 

    def _click_btn_delete_selected(self, sender, e):
        selected_item = self.combobox_phase.SelectedItem
        if selected_item:
            phase_name = selected_item.split(' - ', 1)[1]
            with Transaction(self.doc) as t:
                t.Start('Delete Spaces in "{}" Phase'.format(phase_name))
                deleted_counter = self._delete_spaces_by_phase_name(phase_name)
                t.Commit()
                if t.GetStatus() == TransactionStatus.Committed:
                    self.spaces_by_phase_dct.pop(phase_name)
                    self.combobox_phase.Items.Remove(selected_item)

                    logger.write_log('{} Spaces have been deleted\nin "{}" model Phase'.format(deleted_counter, phase_name), Logger.INFO)
                    message = 'Total {} Spaces have been deleted\nin "{}" model Phase'.format(deleted_counter, phase_name)
                    information_window = InformationWindow('Report', message)
                    information_window.ShowDialog() 
        else:
            message = 'Phase is not selected in the Current model.'
            information_window = InformationWindow('Error', message)
            information_window.ShowDialog()

    def _changed_combobox_link_selection(self, sender, e):
        self.combobox_link_phase.SelectedIndex = -1
        selected_link_item = self.combobox_link.SelectedItem
        if selected_link_item:
            link_name = selected_link_item.split(' - ')[1]
            link_rooms_by_phase_dct = self.rooms_by_link_and_phase_dct[link_name]
            self.combobox_link_phase.Items.Clear()
            for phase_name, rooms in link_rooms_by_phase_dct.items():
                room_number = len(rooms)
                item = '{} Rooms - {}'.format(room_number, phase_name)
                self.combobox_link_phase.Items.Add(item)                
        else:
            message = 'Link is not selected for analize.'
            information_window = InformationWindow('Error', message)
            information_window.ShowDialog()

    def _click_btn_create_selected(self, sender, e):
        selected_link_item = self.combobox_link.SelectedItem
        selected_link_phase_item = self.combobox_link_phase.SelectedItem
        if selected_link_item and selected_link_phase_item:
            link_name = selected_link_item.split(' - ', 1)[1]
            phase_name = selected_link_phase_item.split(' - ', 1)[1]
            link_rooms_from_phase = self.rooms_by_link_and_phase_dct[link_name][phase_name]
            rooms_by_phase_dct = {phase_name: link_rooms_from_phase}

            rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms = self._analize_rooms_by_area_and_level(rooms_by_phase_dct)
            creation_window = CreationWindow(self.doc, self.workset_spaces_id, rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms, self.active_view_phase, self.current_levels)
            creation_window.ShowDialog()
        else:
            message = 'Phase is not selected in the Linked model.'
            information_window = InformationWindow('Error', message)
            information_window.ShowDialog()

    def _fill_combobox_phase(self):
        for phase_name, spaces in self.spaces_by_phase_dct.items():
            spaces_number = len(spaces)
            item = '{} Spaces - {}'.format(spaces_number, phase_name)
            self.combobox_phase.Items.Add(item)

    def _fill_combobox_link(self):
        for link_name, phases in self.rooms_by_link_and_phase_dct.items():
            rooms_number_total = 0
            for rooms in phases.values():
                rooms_number_phase = len(rooms)
                rooms_number_total += rooms_number_phase
            item = '{} Rooms - {}'.format(rooms_number_total, link_name)
            self.combobox_link.Items.Add(item)

    def _delete_all_spaces(self):
        deleleted_counter = 0
        phases_counter = 0
        phases_list = ''
        for phase_name in self.spaces_by_phase_dct.keys():
            phases_counter += 1
            phases_list += '{}\n'.format(phase_name)
            deleleted_counter += self._delete_spaces_by_phase_name(phase_name)
        return deleleted_counter, phases_counter, phases_list

    def _delete_spaces_by_phase_name(self, phase_name):
        deleted_counter = 0
        spaces = self.spaces_by_phase_dct[phase_name]
        for space in spaces.values():
            self.doc.Delete(space.Id)
            deleted_counter += 1 
        return deleted_counter

    def _analize_rooms_by_area_and_level(self, rooms_by_phase_dct):
        sorted_rooms = {'total': 0}
        rooms_area_incorrect = {'total': 0}
        rooms_level_is_missing = {'total': 0, 'names': []}
        rooms_level_incorrect = {'total': 0, 'names': []}

        for phase_name, rooms in rooms_by_phase_dct.items():
            for room_id, room in rooms.items():
                if phase_name not in sorted_rooms:
                    sorted_rooms[phase_name] = {}
                    rooms_area_incorrect[phase_name] = {}

                room_area = room.Area
                room_level = room.Level
                room_level_name = room_level.Name
                room_level_elevation = room_level.ProjectElevation
                room_upper_limit = room.UpperLimit
                if room_upper_limit:
                    room_upper_limit_name = room_upper_limit.Name
                    room_upper_limit_elevation = room_upper_limit.ProjectElevation

                if room_area == 0:
                    rooms_area_incorrect['total'] += 1
                    rooms_area_incorrect[phase_name].update({room_id: room})
                elif room_level_name not in self.current_levels:
                    rooms_level_is_missing['total'] += 1
                    if room_level_name not in rooms_level_is_missing['names']:
                        rooms_level_is_missing['names'].append(room_level_name)
                elif room_level_elevation != self.current_levels[room_level_name]['elevation']:
                    rooms_level_incorrect['total'] += 1
                    if room_level_name not in rooms_level_incorrect['names']:
                        rooms_level_incorrect['names'].append(room_level_name)
                elif (room_upper_limit) and (room_upper_limit_name not in self.current_levels):
                    rooms_level_is_missing['total'] += 1
                    if room_upper_limit_name not in rooms_level_is_missing['names']:
                        rooms_level_is_missing['names'].append(room_upper_limit_name)                    
                elif (room_upper_limit) and (room_upper_limit_elevation != self.current_levels[room_upper_limit_name]['elevation']):
                    rooms_level_incorrect['total'] += 1
                    if room_upper_limit_name not in rooms_level_incorrect['names']:
                        rooms_level_incorrect['names'].append(room_upper_limit_name)
                else:
                    sorted_rooms['total'] += 1
                    sorted_rooms[phase_name].update({room_id: room})
        # DEBUG
        # print '{} {} {} {}'.format(rooms_area_incorrect['total'], rooms_level_is_missing['total'], rooms_level_incorrect['total'], sorted_rooms['total'])
        return rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms


if __name__ == '__main__':
    main_window = MainWindow()
    main_window.ShowDialog()
