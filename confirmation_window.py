import clr
import math
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPIUI')
from System.Windows.Forms import (StatusBar, Form, StatusBar, FormBorderStyle, Label, Button, DialogResult)
from System.Drawing import Size, Point, Font, FontStyle
from math import ceil


class ConfirmationWindow(Form):
    def __init__(self, title, width, message, continue_flag = True):
        self.title = title
        self.width = width
        self.message = message
        self.continue_flag = continue_flag
        self._initialize_components()

    def _initialize_components(self):
        # window
        self.Text = self.title
        self.form_length = 150
        self.form_width = self.width
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
        self.label_size_length = self._define_rows_number(self.message) * self.label_row_length
        if self.label_size_length > (self.form_length - 100):
            self.form_length = self.label_size_length + 80 + self.label_row_length * 2
            self.MinimumSize = Size(self.form_width, self.form_length)
            self.Size = Size(self.form_width, self.form_length)
            self.CenterToScreen()
        self._label_message = Label()
        self._label_message.Text = self.message
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
        if self.continue_flag:
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

    def _click_btn_back(self, sender, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def _click_btn_continue(self, sender, e):
        self.DialogResult = DialogResult.OK
        self.Close()
