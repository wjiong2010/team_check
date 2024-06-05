import os
import csv

row_attr_index = {
    "id": 0,
    "work_item_type": 0,
    "severity": 0,
    "title": 0,
    "status": 0,
    "person_in_charge": 0,
    "dead_line": 0,
    "estimated_man_hours": 0,
    "registered_man_hours": 0,
    "rest_man_hours": 0,
    "reopen_times": 0,
    "creator": 0,
    "creation_time": 0,
    "project": ""
}


class KPIItem:
    def __init__(self):
        self.name_list = []
        self.status_counter = {}
        self.total = 0

    def reset(self):
        # self.name_list.clear()
        self.total = 0
        keys_list = list(self.status_counter.keys())
        for k in keys_list:
            self.status_counter[k] = 0

    def get_info(self):
        # line_format = "{0:>40},{1:>40}"
        ts = " " * 4
        for s in self.name_list:
            ts += s
            ts += '&'
        ts = ts.strip('&')
        ts += ', total {0}:\n'.format(self.total)

        keys_list = list(self.status_counter.keys())
        for k in keys_list:
            if self.status_counter[k] != 0:
                ts += " " * 8 + k + ': ' + str(self.status_counter[k]) + '\n'
        # print(ts)
        return ts

    def do_proc(self, id_v, st, ty):
        print(f"ID: {id_v}, Status: {st}, Type: {ty} in {self.name_list}")
        i = st.find('（')
        if i != -1:
            st = st[:i]
        self.total += 1
        keys_list = list(self.status_counter.keys())
        for k in keys_list:
            print(f"checking {st} in {k}")
            if -1 != k.find(st):
                print(f"{st} in {k}")
                self.status_counter[k] += 1
                break

    def rate_calculater(self, rt_list, pre, count=0):
        if self.total == 0:
            return " " * 4 + pre + " null"

        if len(rt_list) == 0:
            c = count
        else:
            d = self.status_counter
            c = 0
            for r in rt_list:
                c += d[r]

        ratio = "{:.1f}".format(float(c / self.total) * 100.0)
        return " " * 4 + pre + str(ratio) + "%"

    def get_status_count(self, rt_list):
        d = self.status_counter
        c = 0
        for r in rt_list:
            c += d[r]

        return c


class itemFAEBUG(KPIItem):
    def __init__(self):
        super().__init__()
        self.name_list = ["FAE_BUG"]
        self.status_counter = {
            "NO_FEEDBACK": 0,
            "NO_RESPONSE": 0,
            "RESOLVED": 0,
            "REOPEN": 0,
            "WAIT_FEEDBACK": 0,
            "NEW": 0,
            "DOING": 0,
            "ASSIGNED": 0,
            "PENDING": 0,
            "TESTING": 0,
            "REJECTED": 0,
            "WAIT_RELEASE": 0,
            "CLOSED-关闭": 0
        }
        self.fix_pre = "FAE_BUG Fixed: "
        self.reopen_pre = "FAE_BUG Reopened: "
        self.summary = " " * 4 + self.fix_pre + "null\n"
        self.summary += " " * 4 + self.reopen_pre + "null"
        self.reopen_times = 0

    def do_proc_fae(self, id_v, st, ty, rop_t):
        super().do_proc(id_v, st, ty)
        if len(rop_t) != 0:
            self.reopen_times += int(rop_t)

    def calcu_summary(self):
        rt_list = ["NO_FEEDBACK", "RESOLVED", "REJECTED", "WAIT_RELEASE", "CLOSED-关闭", "NO_RESPONSE"]
        self.summary = super().rate_calculater(rt_list, self.fix_pre) + "\n"

        # "REOPEN"
        # reopen率 = 总reopen次数/FAEBUG总数 x 100%
        self.summary += super().rate_calculater([], self.reopen_pre, self.reopen_times) + "\n"

        if self.reopen_times == 0:
            self.summary += "    FAE_BUG Reopened Times: 0"
        else:
            self.summary += "    FAE_BUG Reopened Times: " + str(self.reopen_times)


class itemREQUIREMENT(KPIItem):
    def __init__(self):
        super().__init__()
        self.name_list = ["REQUIREMENT", "需求"]
        self.status_counter = {
            "RESOLVED": 0,
            "REOPEN": 0,
            "WAIT_FEEDBACK": 0,
            "NEW": 0,
            "DOING-进行中": 0,
            "ASSIGNED": 0,
            "PENDING": 0,
            "TESTING": 0,
            "REJECTED": 0,
            "WAIT_RELEASE": 0,
            "关闭": 0,
            "NO_FEEDBACK": 0,
            "NO_RESPONSE": 0
        }
        self.fix_pre = "REQUIREMENT Completed: "
        self.reopen_pre = "REQUIREMENT Reopened: "
        self.summary = " " * 4 + self.fix_pre + "null\n"
        self.summary += " " * 4 + self.reopen_pre + "null"
        self.reopen_times = 0

    def do_proc_req(self, id_v, st, ty, rop_t):
        super().do_proc(id_v, st, ty)
        if len(rop_t) != 0:
            self.reopen_times += int(rop_t)

    def calcu_summary(self):
        rt_list = ["NO_FEEDBACK", "RESOLVED", "REJECTED", "WAIT_RELEASE", "关闭", "WAIT_FEEDBACK", "NO_RESPONSE"]
        self.summary = super().rate_calculater(rt_list, self.fix_pre) + "\n"

        # "REOPEN"
        # reopen率 = 总reopen次数/REQUIREMENT总数 x 100%
        self.summary += super().rate_calculater([], self.reopen_pre, self.reopen_times) + "\n"

        if self.reopen_times == 0:
            self.summary += "    REQUIREMENT Reopened Times: 0"
        else:
            self.summary += "    REQUIREMENT Reopened Times: " + str(self.reopen_times)


class itemPROT_DEV(KPIItem):
    def __init__(self):
        super().__init__()
        self.name_list = ["PROTOCOL", "HF_PROTOCOL", "DEVELOP", "任务", "子任务"]
        self.status_counter = {
            "NO_FEEDBACK": 0,
            "RESOLVED": 0,
            "CLOSED-关闭": 0,
            "WAIT_FEEDBACK": 0,
            "NEW-未开始": 0,
            "DOING-进行中": 0,
            "ASSIGNED": 0,
            "PENDING": 0,
            "TESTING": 0,
            "REJECTED": 0,
            "WAIT_RELEASE": 0,
            "FILED-已完成": 0
        }
        self.fix_pre = "Protocol & Develop Complete: "
        self.summary = " " * 4 + self.fix_pre + "null"

    def calcu_summary(self):
        rt_list = ["NO_FEEDBACK", "RESOLVED", "REJECTED", "WAIT_RELEASE", "CLOSED-关闭", "WAIT_FEEDBACK",
                   "FILED-已完成"]
        self.summary = super().rate_calculater(rt_list, self.fix_pre)


class itemST_BUG(KPIItem):
    def __init__(self):
        super().__init__()
        self.name_list = ["ST_BUG", "HF_ST_BUG", "缺陷"]
        self.status_counter = {
            "NEW-待处理": 0,
            "UNCONFIRMED-待讨论": 0,
            "WAIT_SOLVE（待解决）": 0,
            "RESOLVED-已修复": 0,
            "LOG_Req": 0,
            "拒绝": 0,
            "CLOSED-关闭": 0,
            "REOPENED-重新打开": 0,
            "验证中": 0
        }
        self.reopen_times = 0
        self.fix_pre = "ST_BUG Fixed: "
        self.reopen_pre = "ST_BUG Reopened: "
        self.summary = " " * 4 + self.fix_pre + "null\n"
        self.summary += " " * 4 + self.reopen_pre + "null"

    def do_proc_req(self, id_v, st, ty, rop_t):
        super().do_proc(id_v, st, ty)
        if len(rop_t) != 0:
            self.reopen_times += int(rop_t)

    def calcu_summary(self):
        rt_list = ["拒绝", "RESOLVED-已修复", "CLOSED-关闭", "验证中"]
        self.summary = super().rate_calculater(rt_list, self.fix_pre) + '\n'

        # "REOPENED-重新打开"
        # reopen率 = 总reopen次数/ST_BUG总数 x 100%
        self.summary += super().rate_calculater([], self.reopen_pre, self.reopen_times)


def get_csv_filename(r, fn):
    t_fn = fn.replace(" ", "_") + ".csv"
    return os.path.join(r, t_fn)


class KPIForOnePerson:

    def reset(self, r, name):
        self.fae_bug.reset()
        self.requirement.reset()
        self.st_bug.reset()
        self.prot_dev.reset()
        p = get_csv_filename(r, name)
        if os.path.exists(p):
            os.remove(p)
            print(f"{p} has been removed")

    def parse_kpi_row(self, kpi_row):
        work_type = kpi_row[row_attr_index["work_item_type"]]
        id_v = kpi_row[row_attr_index["id"]]
        st = kpi_row[row_attr_index["status"]]
        print("work type: ", work_type)

        if work_type in self.fae_bug.name_list:
            if row_attr_index["reopen_times"] == 0:
                rop_t = ""
            else:
                rop_t = kpi_row[row_attr_index["reopen_times"]]
            print(f"fae_bug rop_t: {rop_t}")
            self.fae_bug.do_proc_fae(id_v, st, work_type, rop_t)
        elif work_type in self.prot_dev.name_list:
            self.prot_dev.do_proc(id_v, st, work_type)
        elif work_type in self.st_bug.name_list:
            if row_attr_index["reopen_times"] == 0:
                rop_t = ""
            else:
                rop_t = kpi_row[row_attr_index["reopen_times"]]
            print(f"st_bug rop_t: {rop_t}")
            self.st_bug.do_proc_req(id_v, st, work_type, rop_t)
        elif work_type in self.requirement.name_list:
            if row_attr_index["reopen_times"] == 0:
                rop_t = ""
            else:
                rop_t = kpi_row[row_attr_index["reopen_times"]]
            print(f"requirement rop_t: {rop_t}")
            self.requirement.do_proc_req(id_v, st, work_type, rop_t)
        else:
            raise Exception(f"unknown work type: {work_type}")

    def kpi_summary(self):
        self.fae_bug.calcu_summary()
        self.prot_dev.calcu_summary()
        self.st_bug.calcu_summary()
        self.requirement.calcu_summary()

    def __init__(self, name_en, name_cn):
        self.name_en = name_en
        self.name_cn = name_cn
        self.work_type = ''
        self.work_load = ""
        self.work_hour = 0
        self.fae_bug = itemFAEBUG()
        self.requirement = itemREQUIREMENT()
        self.prot_dev = itemPROT_DEV()
        self.st_bug = itemST_BUG()
        # self.report = ""


# ones
# ['\ufeffID', '工作项类型', '严重程度', '标题', '状态', '负责人', '截止日期', '预估工时（小时）', '已登记工时（小时）', '剩余工时（小时）', '重新打开-停留次数', '创建者',
# '创建时间', '所属项目']
# redmin
# ['\ufeff#', '跟踪', 'Severity', '主题', '状态', '指派给', '计划完成日期', '预估工时统计', '耗时', '预期时间', '作者', '创建于', '项目']
def init_row_index(row):
    row_attr_index["id"] = 0
    try:
        row_attr_index["work_item_type"] = row.index("工作项类型")
    except ValueError:
        print("error!!")
        row_attr_index["work_item_type"] = row.index("跟踪")
    r = row_attr_index["work_item_type"]
    print(f"work_item_type: {r}")

    try:
        row_attr_index["severity"] = row.index("严重程度")
    except ValueError:
        print("error!!")
        row_attr_index["severity"] = row.index("Severity")
    r = row_attr_index["severity"]
    print(f"severity: {r}")

    try:
        row_attr_index["title"] = row.index("标题")
    except ValueError:
        print("error!!")
        row_attr_index["title"] = row.index("主题")
    r = row_attr_index["title"]
    print(f"title: {r}")

    row_attr_index["status"] = row.index("状态")
    r = row_attr_index["status"]
    print(f"status: {r}")

    try:
        row_attr_index["person_in_charge"] = row.index("负责人")
    except ValueError:
        print("error!!")
        row_attr_index["person_in_charge"] = row.index("指派给")
    r = row_attr_index["person_in_charge"]
    print(f"person_in_charge: {r}")

    try:
        row_attr_index["dead_line"] = row.index("截止日期")
    except ValueError:
        print("error!!")
        row_attr_index["dead_line"] = row.index("计划完成日期")
    r = row_attr_index["dead_line"]
    print(f"dead_line: {r}")

    try:
        row_attr_index["estimated_man_hours"] = row.index("预估工时（小时）")
    except ValueError:
        print("error!!")
        row_attr_index["estimated_man_hours"] = row.index("预估工时统计")
    r = row_attr_index["estimated_man_hours"]
    print(f"estimated_man_hours: {r}")

    try:
        row_attr_index["registered_man_hours"] = row.index("已登记工时（小时）")
    except ValueError:
        print("error!!")
        row_attr_index["registered_man_hours"] = row.index("耗时")
    r = row_attr_index["registered_man_hours"]
    print(f"registered_man_hours: {r}")

    try:
        row_attr_index["creator"] = row.index("创建者")
    except ValueError:
        print("error!!")
        row_attr_index["creator"] = row.index("作者")
    r = row_attr_index["creator"]
    print(f"creator: {r}")

    try:
        row_attr_index["creation_time"] = row.index("创建时间")
    except ValueError:
        print("error!!")
        row_attr_index["creation_time"] = row.index("创建于")
    r = row_attr_index["creation_time"]
    print(f"creation_time: {r}")

    try:
        row_attr_index["project"] = row.index("所属项目")
    except ValueError:
        print("error!!")
        row_attr_index["project"] = row.index("项目")
    r = row_attr_index["project"]
    print(f"project: {r}")

    try:
        row_attr_index["reopen_times"] = row.index("重新打开-停留次数")
    except ValueError:
        print("error!!")
        # row_attr_index["project"] = row.index("项目")
    r = row_attr_index["reopen_times"]
    print(f"reopen_times: {r}")

    print(row_attr_index)


def __get_name(row_list):
    """
    :param row_list:
    :return: name in lower case, each word are seperated by space, like: john wang
    """
    name = row_list[row_attr_index["person_in_charge"]]
    if name.endswith("HFSW"):
        name = name[:-5]
        name = name.replace(".", " ")

    name = name.lower()

    return name


def __get_item_type(row_list):
    return row_list[row_attr_index["work_item_type"]]


def row_parser(row_list, first_row, members):
    name = ""
    if first_row:
        init_row_index(row_list)
    else:
        name = __get_name(row_list)
        print(f"name: {name}")
        for mb in members:
            if mb.name_en.lower() == name:
                mb.kpi.parse_kpi_row(row_list)

    return name


def row_save(r_path, name, row):
    p = get_csv_filename(r_path, name)
    print(f"saving {p}")
    print(f"row: {row}")
    with open(p, 'a', encoding="utf-8") as csv_f:
        writer = csv.writer(csv_f, quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
        writer.writerow(row)


def kpi_process(r_path, csv_list, members):
    for mb in members:
        mb.kpi.reset(r_path, mb.name_en)

    for csv_file in csv_list:
        first_row = True
        fpath = os.path.join(r_path, csv_file)
        print(f"fpath: {fpath}, fn: {csv_file}")

        with open(fpath, 'r', encoding="utf-8") as csv_f:
            reader = csv.reader(csv_f)
            for r_list in reader:
                row_name = row_parser(r_list, first_row, members)
                first_row = False
                row_save(r_path, row_name, r_list)

    for mb in members:
        mb.kpi.kpi_summary()
