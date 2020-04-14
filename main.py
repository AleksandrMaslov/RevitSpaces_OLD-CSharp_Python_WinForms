# -*- coding: utf-8 -*-
import clr
import time
import sys
sys.path.append(r'C:\Program Files (x86)\IronPython 2.7\Lib')
import os
sys.path.append(os.path.dirname(__file__))
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Level, RevitLinkInstance, UV, Transaction
from lite_logging import Logger  # noqa


logger = Logger(parent_folders_path=os.path.join('Synergy Systems', 'Create Spaces From Linked Rooms'),
                file_name='test_log',
                default_status=Logger.WARNING)

# A_Link_Number = 1

transaction = Transaction(doc)  # noqa


def timer(f):
    def tmp(*args, **kwargs):
        t = time.time()
        res = f(*args, **kwargs)
        # print '"{}" function time is {}'.format(f.func_name, (time.time() - t))
        return res
    return tmp


t = time.time()


# @timer
def create_level_dct(doc):
    fec_lvl = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
    dct_lvl = {}
    for lev in fec_lvl:
        dct_lvl[lev.Name] = {}
        dct_lvl[lev.Name]['inst'] = lev
        dct_lvl[lev.Name]['lid'] = lev.Id
    return dct_lvl


# @timer
def create_levelid_dct(doc):
    fec_lvl = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
    dct_lvl = {}
    for lev in fec_lvl:
        dct_lvl[lev.Id] = {}
        dct_lvl[lev.Id]['inst'] = lev
        dct_lvl[lev.Id]['name'] = lev.Name
    return dct_lvl


# @timer
def create_base_dict():
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
        else:
            unl_lnk_number += 1

    dct_base_elem = {}
    if link_doc is None:
        # print 'WARNING: Architectural link is not found.'
        # print 'Plase check {} unloaded link(s).'.format(unl_lnk_number)
        logger.write_log('Architectural link is not found.', Logger.ERROR)
        logger.write_log('Plase check {} unloaded link(s).'.format(unl_lnk_number), Logger.DEBUG)
    else:
        dct_lvl = create_level_dct(doc)
        dct_lvlid = create_levelid_dct(link_doc)
        name_err_fl = 0
        if len(dct_lvl) != len(dct_lvlid):  # Проверка, что количество и имена уровней совпадают
            # print 'WARNING: Different number of Levels.'
            logger.write_log('Different number of Levels.', Logger.ERROR)
        else:
            for _id, features in dct_lvlid.items():
                if features['name'] not in dct_lvl:
                    name_err_fl = 1
            if name_err_fl != 0:
                # print 'WARNING: Different Level names.'
                logger.write_log('Different Level names.', Logger.ERROR)
            else:
                for element in link_elem:
                    if hasattr(element.Location, 'Point'):
                        _id = element.Id
                        dct_base_elem[_id] = {}
                        dct_base_elem[_id]['loc'] = element.Location.Point
                        dct_base_elem[_id]['lvl'] = dct_lvl[dct_lvlid[element.LevelId]['name']]['inst']
                        if element.UpperLimit is None:
                            dct_base_elem[_id]['upl'] = dct_lvl[dct_lvlid[element.LevelId]['name']]['lid']
                        else:
                            dct_base_elem[_id]['upl'] = dct_lvl[dct_lvlid[element.UpperLimit.Id]['name']]['lid']
                        dct_base_elem[_id]['bof'] = element.BaseOffset
                        dct_base_elem[_id]['lof'] = element.LimitOffset
                        dct_base_elem[_id]['num'] = element.Number
                        dct_base_elem[_id]['nam'] = element.GetParameters('Name')[0].AsString()
                        logger.write_log('Room Number "{}" have placed in the model'.format(element.Number), Logger.INFO)
                    else:
                        # print 'Room Number "{}" is not placed in the A model'.format(element.Number)
                        logger.add_blank_line()
                        logger.write_log('Room Number "{}" is not placed in the A model'.format(element.Number))
                        logger.add_blank_line()
    return dct_base_elem


# @timer
def create_new_instance(main_dct):
    counter = 0
    for _id, features in main_dct.items():
        loc_p = features['loc']
        try:
            instance = doc.Create.NewSpace(features['lvl'], UV(loc_p.X, loc_p.Y))
            instance.GetParameters('Number')[0].Set(features['num'])
            instance.GetParameters('Name')[0].Set(features['nam'])
            instance.GetParameters('Base Offset')[0].Set(features['bof'])
            instance.GetParameters('Limit Offset')[0].Set(features['lof'])
            instance.GetParameters('Upper Limit')[0].Set(features['upl'])

            counter += 1
            logger.write_log('Spaces Placed: {}'.format(counter, Logger.INFO))
        except Exception as e:
            logger.write_log('{}\nНомер помещения: {}'.format(e, features['num']))


# @timer
def delete_instance():
    fec_spc = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    del_num = 0
    for el in fec_spc:
        doc.Delete(el.Id)
        del_num += 1
    # print 'Spaces Deleted: {}\n'.format(del_num)
    logger.write_log('Spaces Deleted: {}'.format(del_num), Logger.INFO)
    return


logger.write_log('Запуск состоялся', Logger.INFO)
transaction.Start('Delete Old Spaces')
delete_instance()
transaction.Commit()

base_dct = create_base_dict()

transaction.Start('Create Spaces')
create_new_instance(base_dct)
transaction.Commit()
