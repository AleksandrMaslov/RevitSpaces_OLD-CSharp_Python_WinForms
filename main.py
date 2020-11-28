# -*- coding: utf-8 -*-
import clr
import sys
sys.path.append(r'C:\Program Files (x86)\IronPython 2.7\Lib')
import os
sys.path.append(os.path.dirname(__file__))
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialog  # noqa
from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, Level, RevitLinkInstance, UV, Transaction,  # noqa
                               FilteredWorksetCollector, WorksetKind, Phase, BuiltInParameter)
from lite_logging import Logger  # noqa

logger = Logger(parent_folders_path=os.path.join('Synergy Systems', 'Create Spaces From Linked Rooms'),
                file_name='test_log',
                default_status=Logger.WARNING)
transaction = Transaction(doc)  # noqa


def create_phase_dct(doc,link_doc):
    fec_ph_c = FilteredElementCollector(doc).OfClass(Phase).ToElements()
    fec_ph_l = FilteredElementCollector(link_doc).OfClass(Phase).ToElements()
    dct_phase = {}
    dct_phase_c = {}
    dct_phase_l = {}

    for phase in fec_ph_c:
        dct_phase_c[phase.Name] = phase
    dct_phase_l = {}
    for phase in fec_ph_l:
        dct_phase_l[phase.Id] = phase.Name
    if len(dct_phase_c) == len(dct_phase_l):
        for _id, nam in dct_phase_l.items():
            if nam not in dct_phase_c:
                dct_phase = {}
            else:
                dct_phase[_id] = dct_phase_c[nam]
    return dct_phase


def workset_spaces():
    sws = 0
    fec_works = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for ws in fec_works:
        if ws.Name == 'Model Spaces':
            sws = ws.Id.IntegerValue
    return sws


def create_level_dct(doc):
    fec_lvl = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
    dct_lvl = {}
    for lev in fec_lvl:
        dct_lvl[lev.Name] = {}
        dct_lvl[lev.Name]['inst'] = lev
        dct_lvl[lev.Name]['lid'] = lev.Id
    return dct_lvl


def create_levelid_dct(doc):
    fec_lvl = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
    dct_lvl = {}
    for lev in fec_lvl:
        dct_lvl[lev.Id] = {}
        dct_lvl[lev.Id]['inst'] = lev
        dct_lvl[lev.Id]['name'] = lev.Name
    return dct_lvl


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


def delete_instance():
    msg = ''
    fec_spc = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    del_num = 0
    for el in fec_spc:
        doc.Delete(el.Id)
        del_num += 1
    logger.write_log('Spaces Deleted: {}'.format(del_num), Logger.INFO)
    msg = msg + '\nSpaces Deleted: {}'.format(del_num)
    TaskDialog.Show('INFORMATION', msg)
    return


def main():
    logger.write_log('Launched successfully', Logger.INFO)
    transaction.Start('Delete Old Spaces')
    delete_instance()
    transaction.Commit()

    transaction.Start('Create Spaces')
    create_new_instance()
    transaction.Commit()


main()
