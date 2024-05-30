import xml.dom.minidom
from xml.dom import DOMException
from codereview import CR_ELEMENT_NODE


def parse_work(m, node):
    for subNode in node.childNodes:
        if CR_ELEMENT_NODE == subNode.nodeType:
            if 'work_app' == subNode.nodeName:
                m.work_apps = subNode.getAttribute('item').split(',')
            elif 'work_module' == subNode.nodeName:
                m.work_modules = subNode.getAttribute('item').split(',')
            elif 'work_gl_app' == subNode.nodeName:
                m.work_gl_apps = subNode.getAttribute('item').split(',')
            elif 'work_pro_app' == subNode.nodeName:
                m.work_pro_apps = subNode.getAttribute('item').split(',')
            else:
                print("UNKNOWN node type: " + subNode.nodeName)


class Team:
    class Member:
        def __init__(self):
            self.name_en = ''
            self.name_cn = ''
            self.work_modules = []
            self.work_apps = []
            self.work_gl_apps = []
            self.work_pro_apps = []
            self.cr_result = []

    def save_as_text(self, txt):
        summary_in_text = ''

        if len(self.members) == 0:
            print('No members found.')
            return -1

        for mb in self.members:
            summary_in_text += f"{mb.name_cn}({mb.name_en}):\n"
            for cr in mb.cr_result:
                summary_in_text += f"\t{cr.severity.upper()}:\n"
                summary_in_text += f"\t{cr.msg}\n"
                if len(cr.msg) != len(cr.verbose):
                    summary_in_text += f"\t{cr.verbose}\n"
                for location in cr.locations:
                    if len(cr.locations) == 1:
                        summary_in_text += f"\t\t{cr.id}:\n"
                    else:
                        summary_in_text += f"\t\t{location.info}:\n"
                    summary_in_text += f"\t\t{location.filename} line:{location.line}, col:{location.column}\n"
                summary_in_text += f"\t修复结果：\n"
                summary_in_text += '\n'
            summary_in_text += '\n\n'

        with open(txt, 'w', encoding='utf-8') as f:
            f.write(summary_in_text)

    def parse_developer(self, node):
        for subNode in node.childNodes:
            if CR_ELEMENT_NODE == subNode.nodeType:
                if 'developer' == subNode.nodeName:
                    m = self.Member()
                    m.name_en = subNode.getAttribute('name_en')
                    m.name_cn = subNode.getAttribute('name_cn')
                    parse_work(m, subNode)
                    self.members.append(m)

    def print_members(self):
        for mb in self.members:
            print(mb.name_cn)
            print(mb.name_en)
            print(mb.work_apps)
            print(mb.work_modules)
            print(mb.work_gl_apps)
            print(mb.work_pro_apps)

    def init_members(self):
        file = "xml\\team.xml"
        try:
            dom = xml.dom.minidom.parse(file)
        except DOMException:
            raise Exception(file + ' is NOT a well-formed XML file.')

        root = dom.documentElement

        if root.nodeName != 'team':
            raise Exception('XML has no \'Team\' element.')

        # Get attributes of results
        results_ver = root.getAttribute('version')
        print(results_ver)

        for node in root.childNodes:
            if CR_ELEMENT_NODE == node.nodeType:  # 1 is Element
                if 'application' == node.nodeName or 'system' == node.nodeName:
                    self.parse_developer(node)

    def __init__(self):
        self.members = []
