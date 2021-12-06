import clr
import math
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPIUI')
from System.Windows.Forms import (StatusBar, Form, StatusBar, FormBorderStyle, Label, Button, DialogResult)
from System.Drawing import Size, Point, Font, FontStyle
from math import ceil


class CreationWindow(Form):
    def __init__(self, rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms, active_view_phase, current_levels):
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
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def _click_btn_continue(self, sender, e):
        self.DialogResult = DialogResult.OK
        self.Close()
