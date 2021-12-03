import clr
import os
import math
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPIUI')
from System.Windows.Forms import (StatusBar, Form, StatusBar, FormBorderStyle, Label, Button)
from System.Drawing import Size, Point, Font, FontStyle
from Autodesk.Revit.DB import Transaction, TransactionStatus, UV, BuiltInParameter
from Autodesk.Revit.UI import TaskDialog
from math import ceil
from lite_logging import Logger
from information_window import InformationWindow

logger = Logger(parent_folders_path=os.path.join('Synergy Systems', 'Create Spaces From Linked Rooms'),
                file_name='test_log',
                default_status=Logger.WARNING)


class CreationWindow(Form):
    def __init__(self, doc, workset_spaces_id, rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms, active_view_phase, current_levels):
        self.doc = doc
        self.workset_spaces_id = workset_spaces_id
        self.current_levels = current_levels
        self.rooms_area_incorrect = rooms_area_incorrect
        self.rooms_level_is_missing = rooms_level_is_missing
        self.rooms_level_incorrect = rooms_level_incorrect
        self.sorted_rooms = sorted_rooms
        self.active_view_phase = active_view_phase
        self._initialize_components()

    def _initialize_components(self):
        # window
        self.Text = 'Spaces Creation'
        self.form_length = 150
        self.form_width = 600
        self.MinimumSize = Size(self.form_width, self.form_length)
        self.Size = Size(self.form_width, self.form_length)
        self.CenterToScreen()
        self.ShowIcon = False
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow

        # labels
        self.label_offset_top = 15
        self.label_offset_left = 20
        self.label_size_width = self.form_width - self.label_offset_left * 2 - 5
        self.label_row_length = 22
        message = self._define_label_message()
        self.label_size_length = self._define_rows_number(message) * self.label_row_length
        if self.label_size_length > (self.form_length - 100):
            self.form_length = self.label_size_length + 80 + self.label_row_length * 2
            self.MinimumSize = Size(self.form_width, self.form_length)
            self.Size = Size(self.form_width, self.form_length)
            self.CenterToScreen()
        self._label_message = Label()
        self._label_message.Text = message
        self._label_message.Location = Point(self.label_offset_left, self.label_offset_top)
        self._label_message.Size = Size(self.label_size_width, self.label_size_length)
        self._label_message.Font = Font("Arial", 10, FontStyle.Regular)
        self.Controls.Add(self._label_message)

        # buttons
        self.button_width = 120 
        self.button_length = 25
        self.button_offset = 22
        self.button_location_Y = self.form_length - self.button_length - 80

        btn_continue = Button()
        btn_continue.Text = 'Continue'
        btn_continue.Size = Size(self.button_width, self.button_length)
        btn_continue.Location = Point(self.label_offset_left, self.button_location_Y)
        if self.sorted_rooms['total'] > 0:
            self.Controls.Add(btn_continue)
        btn_continue.Click += self._click_btn_continue

        btn_back = Button()
        btn_back.Text = 'Back'
        btn_back.Size = Size(self.button_width, self.button_length)
        btn_back.Location = Point(self.form_width - self.button_width - self.label_offset_left * 2, self.button_location_Y)
        self.Controls.Add(btn_back)
        btn_back.Click += self._click_btn_back
   
        #satusbar
        self.statusbar = StatusBar()
        self.statusbar.Parent = self

    def _define_rows_number(self, message):
        phrases = message.split('\n')
        rows_total = len(phrases)
        for phrase in phrases:
            phrase_length = len(phrase)
            if phrase_length > 0:
                additional_rows_number = ceil(phrase_length/ 95.0) - 1
            rows_total += additional_rows_number
        return rows_total

    def _define_label_message(self):
        message = ''
        warnings = 'WARNINGS (!)'

        sorted_rooms_total = self.sorted_rooms['total']
        if sorted_rooms_total > 0:
            if sorted_rooms_total == 1:
                phrase_begins = 'Space is'
            else:
                phrase_begins = 'Spaces are'
            message_sorted_rooms = 'Total {} {} ready for creation in the "{}" Phase.'.format(sorted_rooms_total, phrase_begins, self.active_view_phase)
            message += '{}\n\n'.format(message_sorted_rooms) 

        rooms_area_incorrect_total = self.rooms_area_incorrect['total']
        self.rooms_area_incorrect.pop('total')
        if rooms_area_incorrect_total > 0:
            if not warnings in message:
                message += '{}\n'.format(warnings)
            if rooms_area_incorrect_total == 1:
                phrase_begins = 'Space'
            else:
                phrase_begins = 'Spaces'
            phase_names = ''
            for phase_name in self.rooms_area_incorrect.keys():
                phase_names += '{}\n'.format(phase_name)
            message_rooms_area_incorrect = '- {} {} can not be created because of incorrect Rooms placement in the selected Link model.\nThey can be "Not placed", "Not in a properly enclosed region" or they can be a "Redundant Room". Please send the request to Architectural team to check Rooms in the following Phases of the Link model to solve this problem:\n{}'.format(rooms_area_incorrect_total, phrase_begins, phase_names)
            message += '{}\n'.format(message_rooms_area_incorrect)

        missing_levels_total = self.rooms_level_is_missing['total']
        if missing_levels_total > 0:
            if not warnings in message:
                message += '{}\n'.format(warnings)
            if missing_levels_total == 1:
                phrase_begins = 'Space'
            else:
                phrase_begins = 'Spaces'
            level_names = ''
            for level_name in self.rooms_level_is_missing['names']:
                level_names += '{}\n'.format(level_name)
            message_missing_levels = '- {} {} can not be created because of missing Levels in the Current model.\nPlease add the following levels with the same elvation to the Current Model to create new Spaces correctly:\n{}'.format(missing_levels_total, phrase_begins, level_names)
            message += '{}\n'.format(message_missing_levels)

        incorrect_levels_total = self.rooms_level_incorrect['total']
        if incorrect_levels_total > 0:
            if not warnings in message:
                message += '{}\n'.format(warnings)
            if incorrect_levels_total == 1:
                phrase_begins = 'Space'
            else:
                phrase_begins = 'Spaces'
            level_names = ''
            for level_name in self.rooms_level_incorrect['names']:
                level_names += '{}\n'.format(level_name)
            message_incorrect_levels = '- {} {} can not be created because of different Levels elevation in the Link and Current Model.\nPlease compare the elevation of the following Levels to solve the difference and create new Spaces correctly:\n{}'.format(incorrect_levels_total, phrase_begins, level_names)
            message += '{}\n'.format(message_incorrect_levels)  
        return message 

    def _click_btn_back(self, sender, e):
        self.Close()

    def _click_btn_continue(self, sender, e):
        self.Close()
        report_message = ''
        self.sorted_rooms.pop('total')
        report_counter = []
        with Transaction(self.doc) as t:
            t.Start('Create Spaces from selected Phase')
            for rooms in self.sorted_rooms.values():
                for room in rooms.values():
                    result = self._create_space_by_room_instance(room)
                    report_counter.append(result)
            self.doc.Regenerate()
            t.Commit()

        number_successful = report_counter.count('successful')
        number_warnings = report_counter.count('warnings')
        if number_successful > 0:
            if number_successful == 1:
                phrase = 'Space has'
            else:
                phrase = 'Spaces have'
            report_message += 'Total {} {} been created Successfully.\n'.format(number_successful, phrase)
        if number_warnings > 0:
            if number_warnings == 1:
                phrase = 'Space has'
            else:
                phrase = 'Spaces have'
            report_message += '{} {} been created with Warnings. Please check the Log.'.format(number_warnings, phrase)

        log_link = os.path.join(os.getenv('appdata'), 'Synergy Systems', 'Create Spaces From Linked Rooms')
        information_window = InformationWindow('Report', report_message, log_link, 'View Logs')
        information_window.ShowDialog()
  
    def _create_space_by_room_instance(self, room):
        room_location_point = room.Location.Point
        room_level_name = room.Level.Name
        space_level = self.current_levels[room_level_name]['instance']
        room_base_offset = room.BaseOffset
        room_limit_offset = room.LimitOffset
        room_number = room.Number
        room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
        room_upper_limit = room.UpperLimit
        if room_upper_limit:
            room_upper_limit_name = room_upper_limit.Name
            room_upper_limit = self.current_levels[room_upper_limit_name]['instance']
            room_upper_limit_level_id = room_upper_limit.Id #Level ID 
        else:
            room_upper_limit_level_id = space_level.Id

        try:
            space = self.doc.Create.NewSpace(space_level, UV(room_location_point.X, room_location_point.Y))
            space.get_Parameter(BuiltInParameter.ROOM_NUMBER).Set(room_number)
            space.get_Parameter(BuiltInParameter.ROOM_NAME).Set(room_name)
            space.get_Parameter(BuiltInParameter.ROOM_LOWER_OFFSET).Set(room_base_offset)
            space.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(room_limit_offset)
            space.get_Parameter(BuiltInParameter.ROOM_UPPER_LEVEL).Set(room_upper_limit_level_id)
            space.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(self.workset_spaces_id)
            space_phase_name = space.get_Parameter(BuiltInParameter.ROOM_PHASE).AsValueString()
            # DISCUSS. WHY EXCEPT BUT STILL CREATED? 
            logger.write_log('Space "{} {}" have been placed in the phase - {}'.format(room_number, room_name, space_phase_name), Logger.INFO)
            return 'successful'
        except Exception as e:
            logger.write_log('Space "{} {}" error: {}'.format(room_number, room_name, e), Logger.ERROR)
            return 'warnings'
        

if __name__ == '__main__':
    creation_window = CreationWindow('title', 'message')
    creation_window.ShowDialog()
