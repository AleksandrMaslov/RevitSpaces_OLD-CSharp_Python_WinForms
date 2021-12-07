# coding: utf-8
import clr
import os
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Button, StatusBar, Form, StatusBar, FormBorderStyle, GroupBox, ComboBox, Label, DialogResult, RadioButton)
from System.Drawing import Point, Size
from Autodesk.Revit.DB import Transaction, TransactionStatus, BuiltInParameter, UV
from lite_logging import Logger
from confirmation_window import ConfirmationWindow
from information_window import InformationWindow

logger = Logger(parent_folders_path=os.path.join('Synergy Systems', 'Create Spaces From Linked Rooms'),
                file_name='test_log',
                default_status=Logger.WARNING)


class MainWindow(Form):
    def __init__(self, doc, workset_spaces_id, workset_rooms_id, current_spaces_by_phase, current_rooms_by_phase, rooms_by_link_and_phase, current_levels, active_view_phase):
        self.doc = doc
        self.spaces_by_phase_dct = current_spaces_by_phase
        self.rooms_by_phase_dct = current_rooms_by_phase
        self.rooms_by_link_and_phase_dct = rooms_by_link_and_phase
        self.workset_spaces_id = workset_spaces_id
        self.current_levels = current_levels
        self.active_view_phase = active_view_phase
        self._initialize_components()

    def _initialize_components(self):
        # window
        self.Text = 'Spaces Manager'
        self.form_width = 320
        self.form_length = 470
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
        self.groupbox_current_length = 108
        self.groupbox_linked_length = 220
        self.groupbox_linked_location_Y = self.groupbox_current_location_Y + self.groupbox_current_length + 15

        self.groupbox_current = GroupBox()
        self.groupbox_current.Location = Point(self.form_offset_left, self.groupbox_current_location_Y)
        self.groupbox_current.Text = "Current Model Spaces\Rooms by Phase"
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
        self.button_width_xlarge = 145
        self.button_offset = 22
        self.button_location_current_Y = self.groupbox_current_length - self.button_length - 8
        self.button_location_linked_Y = self.groupbox_offset_top + 60 + self.button_length
    
        btn_help = Button()
        btn_help.Text = 'Help'
        btn_help.Parent = self
        btn_help.Size = Size(self.button_width, self.button_length)
        btn_help.Location = Point(self.form_offset_left, self.groupbox_linked_location_Y + self.groupbox_linked_length + 5)
        btn_help.Click += self._click_btn_help

        btn_delete_all = Button()
        btn_delete_all.Text = 'Delete All'
        btn_delete_all.Size = Size(self.button_width_large, self.button_length)
        btn_delete_all.Location = Point(self.groupbox_offset_left, self.button_location_current_Y)
        btn_delete_all.Parent = self.groupbox_current
        btn_delete_all.Click += self._click_btn_delete_all

        btn_delete_selected = Button()
        btn_delete_selected.Text = 'Delete Selected'
        btn_delete_selected.Size = Size(self.button_width_xlarge, self.button_length)
        btn_delete_selected.Location = Point(self.groupbox_width - self.button_width_xlarge - self.groupbox_offset_left - 1, self.button_location_current_Y)
        btn_delete_selected.Parent = self.groupbox_current
        btn_delete_selected.Click += self._click_btn_delete_selected

        btn_create_all = Button()
        btn_create_all.Text = 'Create All'
        btn_create_all.Size = Size(self.button_width_large, self.button_length)
        btn_create_all.Location = Point(self.groupbox_offset_left, self.button_location_linked_Y)
        btn_create_all.Parent = self.groupbox_linked
        btn_create_all.Click += self._click_btn_create_all

        btn_create_selected = Button()
        btn_create_selected.Text = 'Create Selected'
        btn_create_selected.Size = Size(self.button_width_xlarge, self.button_length)
        btn_create_selected.Location = Point(self.groupbox_width - self.button_width_xlarge - self.groupbox_offset_left - 1, self.button_location_linked_Y)
        btn_create_selected.Parent = self.groupbox_linked
        btn_create_selected.Click += self._click_btn_create_selected

        #radiobuttons
        self.radio_button_current_location_Y = self.button_location_current_Y - 28
        self.radio_button_link_location_Y = self.button_location_linked_Y - 28
        self.radio_button_width = 80

        self.radio_buttons_current_spaces = RadioButton()
        self.radio_buttons_current_spaces.Parent = self.groupbox_current
        self.radio_buttons_current_spaces.Text = "Spaces"
        self.radio_buttons_current_spaces.Checked = True
        self.radio_buttons_current_spaces.Size = Size(self.radio_button_width, 20)
        self.radio_buttons_current_spaces.Location = Point(self.groupbox_offset_left, self.radio_button_current_location_Y)

        self.radio_buttons_current_rooms = RadioButton()
        self.radio_buttons_current_rooms.Parent = self.groupbox_current
        self.radio_buttons_current_rooms.Text = "Rooms"
        self.radio_buttons_current_rooms.Size = Size(self.radio_button_width, 20)
        self.radio_buttons_current_rooms.Location = Point(self.groupbox_offset_left + self.radio_button_width + 3, self.radio_button_current_location_Y)

        self.radio_buttons_link_spaces = RadioButton()
        self.radio_buttons_link_spaces.Parent = self.groupbox_linked
        self.radio_buttons_link_spaces.Text = "Spaces"
        self.radio_buttons_link_spaces.Checked = True
        self.radio_buttons_link_spaces.Size = Size(self.radio_button_width, 20)
        self.radio_buttons_link_spaces.Location = Point(self.groupbox_offset_left, self.radio_button_link_location_Y)

        self.radio_buttons_link_rooms = RadioButton()
        self.radio_buttons_link_rooms.Parent = self.groupbox_linked
        self.radio_buttons_link_rooms.Text = "Rooms"
        self.radio_buttons_link_rooms.Size = Size(self.radio_button_width, 20)
        self.radio_buttons_link_rooms.Location = Point(self.groupbox_offset_left + self.radio_button_width + 3, self.radio_button_link_location_Y)


        # label
        self.label_location_linked_Y = self.button_location_linked_Y + self.button_length + 5
        self.label_width = self.groupbox_width - self.groupbox_offset_left * 2

        label_current_phase = Label()
        label_current_phase.Parent = self.groupbox_linked
        label_current_phase.Location = Point(self.groupbox_offset_left, self.label_location_linked_Y)
        label_current_phase.Size = Size(self.label_width, 75)
        label_current_phase.Text = 'New elements Phase:\n{}\n\n* If you need to change the Phase close the addin and open definite view in your model.'.format(self.active_view_phase)

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

        window_title = 'Delete All Spaces'
        message = 'You are going to delete All Spaces in the Current Model.\n\n'
        confirmation_window = ConfirmationWindow(window_title, message)
        if confirmation_window.ShowDialog() == DialogResult.Cancel:
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
            window_title = 'Delete Selected Spaces'
            phase_name = selected_item.split(' - ', 1)[1]
            message = 'You are going to delete Spaces in "{}" Phase in the Current model.\n\n'.format(phase_name)
            confirmation_window = ConfirmationWindow(window_title, message)
            if confirmation_window.ShowDialog() == DialogResult.Cancel:
                return

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

    def _click_btn_create_all(self, sender, e):
        # if workset_spaces_id:
        # else:
        #     logger.write_log('No "Model Spaces" workset. Create it.', Logger.ERROR)
        #     message = 'There is no "Model Spaces" workset in the Current model. Please create it and relaunch the Addin.'
        #     information_window = InformationWindow('Error', message)
        #     information_window.ShowDialog()  
        selected_link_item = self.combobox_link.SelectedItem
        if selected_link_item:
            link_name = selected_link_item.split(' - ', 1)[1]
            rooms_by_phase_dct = self.rooms_by_link_and_phase_dct[link_name]
            number_of_phases_with_spaces = len(rooms_by_phase_dct)
            if number_of_phases_with_spaces > 0:
                window_title = 'Spaces Creation'
                transaction_name = 'Create All Spaces'
                rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms = self._analize_rooms_by_area_and_level(rooms_by_phase_dct)
                message = self._define_creation_message(rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms)

                confirmation_window = ConfirmationWindow(window_title, message, sorted_rooms['total'] > 0)
                if confirmation_window.ShowDialog() == DialogResult.Cancel:
                    return

                self.Close()
                self._spaces_creation_by_sorted_rooms(sorted_rooms, transaction_name)
            else:
                message = 'There are no Rooms in the selected Linked model.'
                information_window = InformationWindow('Information', message)
                information_window.ShowDialog()
        else:
            message = 'Link model is not selected in the list of Links.'
            information_window = InformationWindow('Error', message)
            information_window.ShowDialog()             

    def _click_btn_create_selected(self, sender, e):
        selected_link_item = self.combobox_link.SelectedItem
        selected_link_phase_item = self.combobox_link_phase.SelectedItem
        if selected_link_item and selected_link_phase_item:
            window_title = 'Spaces Creation'
            link_name = selected_link_item.split(' - ', 1)[1]
            phase_name = selected_link_phase_item.split(' - ', 1)[1]
            transaction_name = 'Create Spaces from "{}" Phase of "{}" Link'.format(phase_name, link_name)
            link_rooms_from_phase = self.rooms_by_link_and_phase_dct[link_name][phase_name]
            rooms_by_phase_dct = {phase_name: link_rooms_from_phase}

            rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms = self._analize_rooms_by_area_and_level(rooms_by_phase_dct)
            message = self._define_creation_message(rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms)
            
            confirmation_window = ConfirmationWindow(window_title, message, sorted_rooms['total'] > 0)
            if confirmation_window.ShowDialog() == DialogResult.Cancel:
                return

            self.Close()
            self._spaces_creation_by_sorted_rooms(sorted_rooms, transaction_name)
        else:
            message = 'Phase is not selected in the Linked model.'
            information_window = InformationWindow('Error', message)
            information_window.ShowDialog()

    def _click_btn_help(self, sender, e):
        message = 'Алгоритм работы плагина:\n' \
        '- При запуске считываются пространства и помещения из текущей открытой модели для дальнейших действий с ними (полного и частичного удаления). Удаление осуществляется нажатием кнопок Delete All или Delete Selected.\n\n' \
        '- Считываются подгруженные линки и количество помещений в них для дальнейшего создания аналогичных пространств или помещений в текущей модели (полного и частичного создания). Создание осуществляется нажатием кнопок Create All или Create Selected для конкретного линка или конкретной фазы выбранного линка.\n\n' \
        '- Перед запуском плагина производится проверка на наличие в модели рабочего набора "Model Spaces".\n\n' \
        '- Перед созданием пространств и помещений производится проверка на корректность размещения помещений в выбранном линке, на наличие совпадающих по имени и отметке уровней, содержащих помещения, в линке и текущей модели. Помещения, не прошедшие проверку, не создаются, выводятся в информационном окне подтверждения создания новых пространств или помещений с рекомендациями по устранению ошибок.\n\n' \
        '- При создании новых пространств и помещений производится перенос данных об уровне, координатах расположения, верхней и нижней границе из модели линка. Созданные пространства автоматически попадают в рабочий набор "Model Spaces".\n\n' \
        '                                                     Молодец, читаешь инструкцию <3'
        information_window = InformationWindow('Help', message)
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

            logger.write_log('Space "{} {}" have been placed in the phase - {}'.format(room_number, room_name, space_phase_name), Logger.INFO)
            return 'successful'
        except Exception as e:
            logger.write_log('Space "{} {}" error: {}'.format(room_number, room_name, e), Logger.ERROR)
            return 'warnings'

    def _define_creation_message(self, rooms_area_incorrect, rooms_level_is_missing, rooms_level_incorrect, sorted_rooms):
        message = ''
        warnings = 'WARNINGS (!)'

        sorted_rooms_total = sorted_rooms['total']
        if sorted_rooms_total > 0:
            if sorted_rooms_total == 1:
                phrase_begins = 'Space is'
            else:
                phrase_begins = 'Spaces are'
            message_sorted_rooms = 'Total {} {} ready for creation in the "{}" Phase.'.format(sorted_rooms_total, phrase_begins, self.active_view_phase)
            message += '{}\n\n'.format(message_sorted_rooms) 

        rooms_area_incorrect_total = rooms_area_incorrect['total']
        rooms_area_incorrect.pop('total')
        if rooms_area_incorrect_total > 0:
            if not warnings in message:
                message += '{}\n'.format(warnings)
            if rooms_area_incorrect_total == 1:
                phrase_begins = 'Space'
            else:
                phrase_begins = 'Spaces'
            phase_names = ''
            for phase_name in rooms_area_incorrect.keys():
                phase_names += '{}\n'.format(phase_name)
            message_rooms_area_incorrect = '- {} {} can not be created because of incorrect Rooms placement in the selected Link model.\nThey can be "Not placed", "Not in a properly enclosed region" or they can be a "Redundant Room". Please send the request to Architectural team to check Rooms in the following Phases of the Link model to solve this problem:\n{}'.format(rooms_area_incorrect_total, phrase_begins, phase_names)
            message += '{}\n'.format(message_rooms_area_incorrect)

        missing_levels_total = rooms_level_is_missing['total']
        if missing_levels_total > 0:
            if not warnings in message:
                message += '{}\n'.format(warnings)
            if missing_levels_total == 1:
                phrase_begins = 'Space'
            else:
                phrase_begins = 'Spaces'
            level_names = ''
            for level_name in rooms_level_is_missing['names']:
                level_names += '{}\n'.format(level_name)
            message_missing_levels = '- {} {} can not be created because of missing Levels in the Current model.\nPlease add the following levels with the same elvation to the Current Model to create new Spaces correctly:\n{}'.format(missing_levels_total, phrase_begins, level_names)
            message += '{}\n'.format(message_missing_levels)

        incorrect_levels_total = rooms_level_incorrect['total']
        if incorrect_levels_total > 0:
            if not warnings in message:
                message += '{}\n'.format(warnings)
            if incorrect_levels_total == 1:
                phrase_begins = 'Space'
            else:
                phrase_begins = 'Spaces'
            level_names = ''
            for level_name in rooms_level_incorrect['names']:
                level_names += '{}\n'.format(level_name)
            message_incorrect_levels = '- {} {} can not be created because of different Levels elevation in the Link and Current Model.\nPlease compare the elevation of the following Levels to solve the difference and create new Spaces correctly:\n{}'.format(incorrect_levels_total, phrase_begins, level_names)
            message += '{}\n'.format(message_incorrect_levels)  
        return message 

    def _spaces_creation_by_sorted_rooms(self, sorted_rooms, transaction_name):
        report_message = ''
        sorted_rooms.pop('total')
        report_counter = []
        with Transaction(self.doc) as t:
            t.Start(transaction_name)
            for rooms in sorted_rooms.values():
                for room in rooms.values():
                    result = self._create_space_by_room_instance(room)
                    report_counter.append(result)
            self.doc.Regenerate()
            t.Commit()

            if t.GetStatus() == TransactionStatus.Committed:
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


if __name__ == '__main__':
    main_window = MainWindow()
    main_window.ShowDialog()
