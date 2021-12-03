# -*- coding: utf-8 -*-
import clr
import sys
sys.path.append(r'C:\Program Files (x86)\IronPython 2.7\Lib')
import os
# sys.path.append(os.path.dirname(__file__))
sys.path.append(r'C:\Users\Admin\Desktop\Addins\CreateSpacesFromLinkedRooms')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, Level, RevitLinkInstance, Transaction,
                               FilteredWorksetCollector, WorksetKind, BuiltInParameter)
from lite_logging import Logger
from main_window import MainWindow
from information_window import InformationWindow

logger = Logger(parent_folders_path=os.path.join('Synergy Systems', 'Create Spaces From Linked Rooms'),
                file_name='test_log',
                default_status=Logger.WARNING)
doc = __revit__.ActiveUIDocument.Document
active_view = doc.ActiveView
transaction = Transaction(doc)
# __window__.Close()


def _find_workset_modelspaces_id(doc):
    fec = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for workset in fec:
        if workset.Name == 'Model Spaces':
            workset_id = workset.Id.IntegerValue
            return workset_id


def _create_level_name_dct(doc):
    fec = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
    dct = {}
    for level in fec:
        dct[level.Name] = {}
        dct[level.Name]['instance'] = level
        dct[level.Name]['elevation'] = level.ProjectElevation
    return dct


def _create_link_document_name_dct(doc):
    dct = {}
    fec = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    for link_instance in fec:
        link_document = link_instance.GetLinkDocument()
        if link_document is not None:
            name = link_instance.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            dct[name] = link_document
    return dct


def _create_spaces_by_phase_dct(doc):
    dct = {}
    fec = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    for space in fec:
        phase_name = space.get_Parameter(BuiltInParameter.ROOM_PHASE).AsValueString()
        space_id = space.Id.IntegerValue
        if phase_name not in dct:
            dct[phase_name] = {}
        dct[phase_name].update({space_id: space})
    return dct


def _create_rooms_by_link_and_phase_dct(current_links):
    dct = {}
    for link_name, link_doc in current_links.items():
        if link_name not in dct:
            dct[link_name] = {}
        fec = FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
        for room in fec:
            phase_name = room.get_Parameter(BuiltInParameter.ROOM_PHASE).AsValueString()
            room_id = room.Id.IntegerValue
            if phase_name not in dct[link_name]:
                dct[link_name][phase_name] = {}
            dct[link_name][phase_name].update({room_id: room})
    return dct


def Main():
    logger.write_log('Launched successfully', Logger.INFO)
    workset_spaces_id = _find_workset_modelspaces_id(doc)
    if workset_spaces_id:
        current_levels = _create_level_name_dct(doc)
        current_links = _create_link_document_name_dct(doc)
        current_spaces_by_phase = _create_spaces_by_phase_dct(doc)
        rooms_by_link_and_phase = _create_rooms_by_link_and_phase_dct(current_links)
        active_view_phase = active_view.get_Parameter(BuiltInParameter.VIEW_PHASE).AsValueString()

        mw = MainWindow(doc, workset_spaces_id, current_spaces_by_phase, rooms_by_link_and_phase, current_levels, active_view_phase)
        mw.ShowDialog()
    else:
        logger.write_log('No "Model Spaces" workset. Create it.', Logger.ERROR)
        message = 'There is no "Model Spaces" workset in the Current model. Please create it and relaunch the Addin.'
        information_window = InformationWindow('Error', message)
        information_window.ShowDialog()        


if __name__ == '__main__':
    Main()
