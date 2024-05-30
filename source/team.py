import xml.dom.minidom
from xml.dom import DOMException
from kpi import KPIForOnePerson

CR_ELEMENT_NODE = 1


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


def kpi_head_text():
    report_head = '''任务量等级： 5:非常饱满， 4:饱满，3：适中，2：不足，1：严重不足'''
    report_main_seperator = '''================================================================================='''

    kpi_text = report_head + '\n'
    kpi_text += report_main_seperator + '\n\n'

    return kpi_text


class Team:
    class Member:
        def cr_result_in_text(self):
            s_text = ''
            s_text += f"{self.name_cn}({self.name_en}):\n"
            for cr in self.cr_result:
                s_text += f"\t{cr.severity.upper()}:\n"
                s_text += f"\t{cr.msg}\n"
                if len(cr.msg) != len(cr.verbose):
                    s_text += f"\t{cr.verbose}\n"
                for location in cr.locations:
                    if len(cr.locations) == 1:
                        s_text += f"\t\t{cr.id}:\n"
                    else:
                        s_text += f"\t\t{location.info}:\n"
                    s_text += f"\t\t{location.filename} line:{location.line}, col:{location.column}\n"
                s_text += f"\t修复结果：\n"
                s_text += '\n'
            s_text += '\n\n'

            return s_text

        def kpi_in_text(self):
            report_secondary_seperator = '''------------------------------------------------------------------------'''
            report_summary = '''总结：'''

            kpi_text = self.name_cn + "(" + self.name_en + "):" + '\n'
            kpi_text += self.kpi.fae_bug.get_info() + '\n'
            kpi_text += self.kpi.prot_dev.get_info() + '\n'
            kpi_text += self.kpi.requirement.get_info() + '\n'
            kpi_text += self.kpi.st_bug.get_info() + '\n\n'

            kpi_text += report_summary + '\n'
            kpi_text += "\tWorkload:" + '\n'
            kpi_text += self.kpi.fae_bug.summary + '\n'
            kpi_text += self.kpi.requirement.summary + '\n'
            kpi_text += self.kpi.st_bug.summary + '\n'
            kpi_text += self.kpi.prot_dev.summary + '\n\n'

            kpi_text += report_secondary_seperator + '\n'

            return kpi_text

        def __init__(self):
            self.name_en = ''
            self.name_cn = ''
            self.work_modules = []
            self.work_apps = []
            self.work_gl_apps = []
            self.work_pro_apps = []
            self.cr_result = []
            self.kpi = KPIForOnePerson(self.name_en, self.name_cn)

    def save_as_text(self, select, txt):
        summary_in_text = ''

        if len(self.members) == 0:
            print('No members found.')
            return summary_in_text

        if 'kpi' == select:
            summary_in_text += kpi_head_text()

        for mb in self.members:
            if 'cr_result' == select:
                summary_in_text += mb.cr_result_in_text()
            if 'kpi' == select:
                summary_in_text += mb.kpi_in_text()

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
                    m.kpi = KPIForOnePerson(m.name_en, m.name_cn)
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
