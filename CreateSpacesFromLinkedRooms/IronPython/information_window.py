import clr
import math
import os
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (StatusBar, Form, StatusBar, FormBorderStyle, Label, LinkLabel)
from System.Drawing import Size, Point, Font, FontStyle
from math import ceil


class InformationWindow(Form):
    def __init__(self, title, message, link = '', link_text = ''):
        self.title = title
        self.message = message
        self.link = link
        self.link_text = link_text
        self._initialize_components()

    def _initialize_components(self):
        # window
        self.Text = self.title
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
        self.label_row_length = 18
        self.label_size_length = self._define_rows_number() * self.label_row_length
        if self.label_size_length > (self.form_length - 100):
            self.form_length = self.label_size_length + 60 + self.label_row_length * 2
            self.MinimumSize = Size(self.form_width, self.form_length)
            self.Size = Size(self.form_width, self.form_length)
            self.CenterToScreen()
        self._label_link_location_Y = self.form_length - 90

        self._label_message = Label()
        self._label_message.Location = Point(self.label_offset_left, self.label_offset_top)
        self._label_message.Size = Size(self.label_size_width, self.label_size_length)
        self._label_message.Text = self.message
        self._label_message.Font = Font("Arial", 10, FontStyle.Regular)
        self.Controls.Add(self._label_message)

        self._label_link = LinkLabel()
        self._label_link.Location = Point(self.label_offset_left, self._label_link_location_Y)
        self._label_link.Size = Size(100, 18)
        self._label_link.Text = self.link_text
        # self._label_link.Font = Font("Arial", 10, FontStyle.Regular)
        self.Controls.Add(self._label_link) 
        self._label_link.LinkClicked += self._click_label_link      
   
        #satusbar
        self.statusbar = StatusBar()
        self.statusbar.Parent = self

    def _define_rows_number(self):
        phrases = self.message.split('\n')
        rows_total = len(phrases)
        for phrase in phrases:
            phrase_length = len(phrase)
            if phrase_length > 0:
                additional_rows_number = ceil(phrase_length/ 95.0) - 1
            rows_total += additional_rows_number
        return rows_total

    def _click_label_link(self, sender, e):
        os.startfile('{}'.format(self.link))


if __name__ == '__main__':
    information_window = InformationWindow('title', 'message')
    information_window.ShowDialog()
