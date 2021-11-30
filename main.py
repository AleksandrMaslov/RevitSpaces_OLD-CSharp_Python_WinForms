# -*- coding: utf-8 -*-
import clr
import sys
sys.path.append(r'C:\Program Files (x86)\IronPython 2.7\Lib')
import os
# sys.path.append(os.path.dirname(__file__))
sys.path.append(r'C:\Users\Admin\Desktop\Addins\CreateSpacesFromLinkedRooms')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialog  # noqa
from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, Level, RevitLinkInstance, UV, Transaction,  # noqa
                               FilteredWorksetCollector, WorksetKind, Phase, BuiltInParameter)
from lite_logging import Logger  # noqa
from main_window import MainWindow  # noqa

logger = Logger(parent_folders_path=os.path.join('Synergy Systems', 'Create Spaces From Linked Rooms'),
                file_name='test_log',
                default_status=Logger.WARNING)
doc = __revit__.ActiveUIDocument.Document  # noqa
transaction = Transaction(doc)  # noqa
# __window__.Close()


def _create_phase_name_dct(doc):
    dct = {}
    fec = FilteredElementCollector(doc).OfClass(Phase).ToElements()
    for phase in fec:
        dct[phase.Name] = phase
    return dct


def _create_phase_id_dct(doc):
    dct = {}
    fec = FilteredElementCollector(doc).OfClass(Phase).ToElements()
    for phase in fec:
        dct[phase.Id] = phase
    return dct


def _find_workset_modelspaces_id(doc):
    fec = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for workset in fec:
        if workset.Name == 'Model Spaces':
            workset_id = workset.Id.IntegerValue
            return workset_id
    return 0


def _create_level_name_dct(doc):
    fec = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
    dct = {}
    for level in fec:
        dct[level.Name] = {}
        dct[level.Name]['instance'] = level
        dct[level.Name]['id'] = level.Id
    return dct


def _create_level_id_dct(doc):
    fec = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
    dct = {}
    for level in fec:
        dct[level.Id] = {}
        dct[level.Id]['instance'] = level
        dct[level.Id]['name'] = level.Name
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


def _create_spaces_by_phase_dct(doc, current_phases):
    dct = {}
    fec = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    for space in fec:
        phase_name = space.get_Parameter(BuiltInParameter.ROOM_PHASE).AsValueString()
        space_id = space.Id.IntegerValue
        if phase_name not in dct:
            dct[phase_name] = {}
        dct[phase_name].update({space_id: space})
    return dct


def _create_rooms_by_link_and_phase_dct(doc, current_links):
    dct = {}
    for link_name, link_doc in current_links.items():
        if link_name not in dct:
            dct[link_name] = {}
        fec = FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
        for room in fec:
            phase_name = room.get_Parameter(BuiltInParameter.ROOM_PHASE).AsValueString()
            room_id = room.Id.IntegerValue
            if phase_name not in dct:
                dct[link_name][phase_name] = {}
            dct[link_name][phase_name].update({room_id: room})
    
    for link, phases in dct.items():
        print link
        for room in phases:
            print room
        print
    return dct


def _verification_model_spaces_workset():
    workset_model_spaces_id = _find_workset_modelspaces_id(doc)
    if workset_model_spaces_id != 0:
        status_workset = 'Verified'
    else:
        status_workset = 'Error'
    return status_workset


def _verification():
    status_workset = _verification_model_spaces_workset()
    status_phases = ''
    status_levels = ''

    statuses = [status_workset]
    return statuses


def create_new_instance():
    ws_id = workset_spaces()
    link_inst = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    unl_lnk_number = 0
    link_doc = None
    for lnk in link_inst:
        link_doc_ch = lnk.GetLinkDocument()
        if link_doc_ch is not None:
            link_elem_ch = FilteredElementCollector(link_doc_ch).OfCategory(BuiltInCategory.OST_Rooms) \
                                                                .WhereElementIsNotElementType() \
                                                                .ToElements()
            if ('A' in lnk.Name) and (len(link_elem_ch) > 2):
                link_doc = link_doc_ch
                link_elem = link_elem_ch
                link = doc.GetElement(lnk.GetTypeId())
                link.GetParameters('Room Bounding')[0].Set(1)
        else:
            unl_lnk_number += 1

    if ws_id > 0:
        if link_doc is None:
            logger.write_log('Architectural link is not found.', Logger.ERROR)
            logger.write_log('Plase check {} unloaded link(s).'.format(unl_lnk_number), Logger.DEBUG)
            TaskDialog.Show('ERROR', 'Architectural link is not found.\nPlase check {} unloaded link(s).'.format(unl_lnk_number))
        else:
            # dct_ph = create_phase_dct(doc,link_doc)
            # if len(dct_ph) == 0:
            #     logger.write_log('Different phases in the MEP and A models.', Logger.ERROR)
            #     TaskDialog.Show('ERROR', 'Different phases in the MEP and A models.\nPlease compare with the Architectural model.')
            # else:
            dct_lvl = create_level_dct(doc)
            dct_lvlid = create_levelid_dct(link_doc)
            name_err_fl = 0
            if len(dct_lvl) != len(dct_lvlid):
                logger.write_log('Different number of Levels.', Logger.ERROR)
                TaskDialog.Show('ERROR', 'Different number of levels.\nPlease compare with the Architectural model.')
            else:
                for _id, features in dct_lvlid.items():
                    if features['name'] not in dct_lvl:
                        name_err_fl = 1
                if name_err_fl != 0:
                    logger.write_log('Different Level names.', Logger.ERROR)
                    TaskDialog.Show('ERROR', 'Different Level names.\nPlease compare with the Architectural model.')
                else:
                    counter = 0
                    msg = 'Detected  A link - ' + link.GetParameters('Type Name')[0].AsString()
                    logger.write_log('Detected  A link - ' + link.GetParameters('Type Name')[0].AsString(), Logger.INFO)
                    act_viw_phid = doc.ActiveView.GetParameters('Phase')[0].AsElementId()
                    for element in link_elem:
                        if hasattr(element.Location, 'Point'):
                            # phase = dct_ph[element.GetParameters('Phase Id')[0].AsElementId()]
                            # if phase.Id == act_viw_phid:
                            loc_p = element.Location.Point
                            level = dct_lvl[dct_lvlid[element.LevelId]['name']]['inst']
                            bof = element.BaseOffset
                            lof = element.LimitOffset
                            num = element.Number
                            nam = element.GetParameters('Name')[0].AsString()
                            if element.UpperLimit is None:
                                upl = dct_lvl[dct_lvlid[element.LevelId]['name']]['lid']
                            else:
                                upl = dct_lvl[dct_lvlid[element.UpperLimit.Id]['name']]['lid']
                            try:
                                instance = doc.Create.NewSpace(level, UV(loc_p.X, loc_p.Y))
                                instance.GetParameters('Number')[0].Set(num)
                                instance.GetParameters('Name')[0].Set(nam)
                                instance.GetParameters('Base Offset')[0].Set(bof)
                                instance.GetParameters('Limit Offset')[0].Set(lof)
                                instance.GetParameters('Upper Limit')[0].Set(upl)
                                instance.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(ws_id)
                                counter += 1
                                c_p = instance.GetParameters('Phase')[0].AsValueString()
                                logger.write_log('Space {} have placed in phase - {}'.format(num, c_p), Logger.INFO)
                                # logger.write_log('Space {} have placed in phase - {} ({})'.format(num, c_p, phase.Name), Logger.INFO)
                            except Exception as e:
                                logger.write_log('Space Number: {}\n{}'.format(num, e), Logger.ERROR)
                                msg = msg + '\nSpace Number: {}\n{}'.format(num, e)
                        else:
                            logger.add_blank_line()
                            logger.write_log('Room Number "{}" is not placed in the A model'.format(element.Number))
                            logger.add_blank_line()
                            msg = msg + '\nRoom "{}" is not placed in the A model'.format(element.Number)
                    logger.write_log('Spaces Placed: {}'.format(counter), Logger.INFO)
                    msg = msg + '\n\nREPORT:'
                    msg = msg + '\nSpaces Placed: {}'.format(counter)
                    log_link = os.path.join(os.getenv('appdata'), 'Synergy Systems', 'Create Spaces From Linked Rooms')
                    dialog = TaskDialog('INFORMATION')
                    dialog.MainInstruction = msg
                    dialog.MainContent = '<a href=\"{} \">'.format(log_link) + 'Open Logs folder</a>'
                    dialog.Show()
                    doc.Regenerate()
    else:
        logger.write_log('No "Model Spaces" workset. Create it.', Logger.ERROR)
        TaskDialog.Show('ERROR', 'No "Model Spaces" workset. Please create it.')


def Main():
    logger.write_log('Launched successfully', Logger.INFO)
    statuses = _verification()
    current_phases = _create_phase_name_dct(doc)
    current_links = _create_link_document_name_dct(doc)
    current_spaces_by_phase = _create_spaces_by_phase_dct(doc, current_phases)
    rooms_by_link_and_phase = _create_rooms_by_link_and_phase_dct(doc, current_links)

    mw = MainWindow(doc, current_spaces_by_phase, rooms_by_link_and_phase)
    mw.ShowDialog()


if __name__ == '__main__':
    Main()
