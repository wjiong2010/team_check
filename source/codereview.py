import os
import xml.dom.minidom
from xml.dom import DOMException
from team import CR_ELEMENT_NODE


class ResultLocation:
    def __init__(self):
        self.filename = ''
        self.line = ''
        self.column = ''
        self.info = ''


class CodeReviewResult:
    def __init__(self):
        self.id = ''
        self.severity = ''
        self.msg = ''
        self.verbose = ''
        self.cwe = ''
        self.file0 = ''
        self.locations = []


def __arrange_results(cr_result, team):
    if len(cr_result.file0) != 0:
        code_file_path = cr_result.file0
        code_file_path = code_file_path.replace('/', '\\')
    elif len(cr_result.locations) != 0 and cr_result.locations[0].filename != '':
        code_file_path = cr_result.locations[0].filename
    else:
        code_file_path = ''

    is_pro_track = False
    fw_track_tag = "framework\\track\\"
    len_fw_track_tag = len(fw_track_tag)
    print("tag_path: " + code_file_path)
    module_index = code_file_path.find(fw_track_tag)
    if module_index == -1:
        fw_track_tag = "framework\\trackPro\\"
        len_fw_track_tag = len(fw_track_tag)
        print("pro_tag_path: " + code_file_path)
        module_index = code_file_path.find(fw_track_tag)
        is_pro_track = True
        if module_index == -1:
            return -1

    module_index += len_fw_track_tag
    # get module name from file0
    module_name = code_file_path[module_index:module_index + 6]
    module_name = module_name.upper()
    print("module_index: " + str(module_index) + "  module_name: " + module_name)
    if 'IO' == module_name[:2] or \
            'GPS' == module_name[:3] or \
            'MCU' == module_name[:3] or \
            'BLE' == module_name[:3] or \
            'DBG' == module_name[:3] or \
            'COMM' == module_name[:4] or \
            'ATCI' == module_name[:4] or \
            'PROT' == module_name[:4] or \
            'UTILS' == module_name[:5] or \
            'SERIAL' == module_name[:6] or \
            'SENSOR' == module_name[:6]:
        i = module_name.rfind('\\')
        if -1 != i:
            module_name = module_name[:i]
        print("proc module_name: " + module_name + " i:" + str(i))
        for mb in team.members:
            if module_name in mb.work_modules:
                mb.cr_result.append(cr_result)
    elif 'APPSGL' == module_name[:6]:
        app_i = module_index + 7  # skip 'appsGL/'
        app_name = code_file_path[app_i:app_i + 3].upper()
        print("gl_app_i: " + str(app_i) + " gl_app_name: " + app_name)
        for mb in team.members:
            if app_name in mb.work_gl_apps:
                mb.cr_result.append(cr_result)
    elif 'APPS' == module_name[:4]:
        app_i = module_index + 5  # skip 'apps\'
        app_j = code_file_path[app_i:].find("\\")
        app_name = code_file_path[app_i:app_i + app_j].upper()
        print("app_i,j: " + str(app_i) + "," + str(app_j) + " app_name: " + app_name)
        if is_pro_track:
            for mb in team.members:
                if app_name in mb.work_pro_apps:
                    mb.cr_result.append(cr_result)
        else:
            for mb in team.members:
                if app_name in mb.work_apps:
                    mb.cr_result.append(cr_result)
    else:
        print("module_name: " + module_name)

    return 0


# get attribute of node,
def __parse_attr(node, attr):
    value = node.getAttribute(attr)
    # not found will trigger exception, 'element xxx has no attributes.'
    if "" == value:
        print('Element \'' + node.nodeName + '\' has no \'' + attr + '\' attribute.')
    return value


def __parse_location(node, locs):
    for subNode in node.childNodes:
        if CR_ELEMENT_NODE == subNode.nodeType:
            loc = ResultLocation()
            if 'location' == subNode.nodeName:
                loc.filename = __parse_attr(subNode, 'file')
                loc.line = __parse_attr(subNode, 'line')
                loc.column = __parse_attr(subNode, 'column')
                loc.info = __parse_attr(subNode, 'info')
                locs.append(loc)


def __parse_errors(node, team):
    for subNode in node.childNodes:
        if CR_ELEMENT_NODE == subNode.nodeType:
            cr_res = CodeReviewResult()
            if 'error' == subNode.nodeName:
                cr_res.id = __parse_attr(subNode, 'id')
                cr_res.severity = __parse_attr(subNode, 'severity')
                cr_res.msg = __parse_attr(subNode, 'msg')
                cr_res.verbose = __parse_attr(subNode, 'verbose')
                cr_res.file0 = __parse_attr(subNode, 'file0')
                cr_res.cwe = __parse_attr(subNode, 'cwe')
                __parse_location(subNode, cr_res.locations)
                print("len of locations: " + str(len(cr_res.locations)))
                __arrange_results(cr_res, team)


def cr_parse_result(r_path, xml_list, team):
    for file in xml_list:
        f_xml = os.path.join(r_path, file)
        print(f"fpath: {f_xml}, fn: {xml_list}")

        try:
            dom = xml.dom.minidom.parse(str(f_xml))
        except DOMException:
            raise Exception(f_xml + ' is NOT a well-formed XML file.')

        root = dom.documentElement

        if root.nodeName != 'results':
            raise Exception('XML has no \'Project\' element.')

        # Get attributes of results
        results_ver = __parse_attr(root, 'version')
        print(results_ver)

        # for mb in app_tm.members:
        #     print(mb.name_cn + '(' + mb.name_en + ')' + ' ')

        for node in root.childNodes:
            if CR_ELEMENT_NODE == node.nodeType:  # 1 is Element
                if 'errors' == node.nodeName:
                    __parse_errors(node, team)
