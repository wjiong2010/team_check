import os
import csv
from datetime import datetime



def get_csv_filename(r, fn):
    t_fn = fn.replace(" ", "_") + ".csv"
    return os.path.join(r, t_fn)


class KPIRow:
    def get_name(self, row_list):
        """
        :param row_list: pick up name from row_list, 
                eg, #27411,任务,,测试FRI ERI命令及报文,Normal,进行中,Easy,aidan.chen-HFSW,,0h,QPDC,8.5h,0h,aidan.chen-HFSW,2025-01-24 16:56:31,,,,,,,,,
                the name of person in charge is aidan.chen-HFSW, we should covert it to 'aidan chen'
        :return: name in lower case, each word are seperated by space
        """
        name = row_list[self.row_index["person_in_charge"]]
        if name.endswith("HFSW"):
            name = name[:-5]
            name = name.replace(".", " ")

        name = name.lower()

        return name
    
    def init_folder(self, r, name):
        self._path = r
        p = get_csv_filename(r, name)
        if os.path.exists(p):
            os.remove(p)
            print("Clear exist file: " + p)

    def get_kpi_row(self):
        """
        :return: the kpi row
        """
        return [self.id, self.work_item_type, self.severity, self.title, self.priority, self.status, self.difficulty, self.person_in_charge, self.dead_line, self.complete_time, self.bug_fix_cycle, self.bug_fixed_duration, self.estimated_workinghours, self.registered_workinghours, self.rest_workinghours, self.project_name, self.creator, self.creation_time, self.reopen_times, self.reopen_duration, self.reopen_confirm_times, self.reopen_confirm_duration, self.state_detail, self.evaluating_duration, self.assigned_duration, self.assigned_times, self.pending_times, self.pending_duration]
    
    # ones
    # ID, 工作项类型, 严重程度, 标题, 优先级, 状态, Difficulty Degree, 负责人, 截止日期, 预估工时（小时）, 所属项目, 已登记工时（小时）, 剩余工时（小时）, 创建者, 创建时间, 
    # 实际完成时间, REOPEN-停留次数, REOPEN-停留时间, REOPEN_CONFIRM-停留次数, REOPEN_CONFIRM-停留时间, 已修复-停留时间, 缺陷修复周期, State Details, EVALUATING-停留时间
    # ,ASSIGNED-停留时间,ASSIGNED-停留次数,PENDING-停留次数,PENDING-停留时间
    # redmin
    # ['#', '跟踪', 'Severity', '主题', '状态', '指派给', '计划完成日期', '预估工时统计', '耗时', '预期时间', '作者', '创建于', '项目', '结束日期']
    def init_row_index(self, row):
        print(str(row))
        for attr in self.row_attribut_list:
            _name = attr["name"]
            _keys = attr["key_words"]
            _def  = attr["default_index"]
            try:
                _index = row.index(_keys[0])
                self.head_line.append(_keys[0])
            except ValueError:
                print("error!!")
                if len(_keys) > 1:
                    _index = row.index(_keys[1])
                else:
                    _index = _def
            print("{}: {}".format(_name, _index))
            self.row_index.update({_name: _index})
        
        print(self.row_index)
        print(self.head_line)
    
    def save_csv(self, row):
        path = os.path.join(self._path, self.member_name + ".csv")
        with open(path, 'a', encoding="utf-8-sig") as csv_f:
            writer = csv.writer(csv_f, quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
            if self.is_first_row:
                print("create the first line for {}".format(csv_f))
                writer.writerow(self.head_line)
            writer.writerow(row)

    def save_kpi_row(self, mb, format = "csv"):
        final_row = [mb.kpi.oper_res]
        final_row += self.get_kpi_row()
        print(f"final_row: {final_row}")
        if format == "csv":
            self.save_csv(final_row)
        else:
            print("save as text or push to database")

    def pre_proc(self, row_list, members):

        print("pre_proc: {}".format(self.is_first_row))
        if self.is_first_row:
            self.init_row_index(row_list)
        else:
            # split the row into different items and save them to the corresponding parameters
            print("sssss:   " + str(row_list))            

            # get member from members by name
            self.member_name = self.get_name(row_list)
            member = None
            print(f"name: {self.member_name}")
            for mb in members:
                if mb.name_en.lower() == self.member_name:
                    member = mb
                    break
            
            if member is not None:
                print(f"member {self.member_name} found")
                return member
            else:
                print(f"member {self.member_name} not found")
                return None

    def __init__(self):
        self.id = ""                    
        self.work_item_type = ""        
        self.severity = ""              
        self.priority = ""              
        self.title = ""                 
        self.status = ""                
        self.difficulty = ""            
        self.person_in_charge = ""      
        self.dead_line = ""             
        self.complete_time = ""         
        self.bug_fix_cycle = ""         
        self.bug_fixed_duration = ""    
        self.estimated_workinghours = ""
        self.registered_workinghours = ""
        self.rest_workinghours = ""     
        self.project_name = ""          
        self.creator = ""               
        self.creation_time = ""         
        self.reopen_times = ""          
        self.reopen_duration = ""       
        self.reopen_confirm_times = ""  
        self.reopen_confirm_duration = ""
        self.state_detail = ""          
        self.evaluating_duration = ""   
        self.assigned_duration = ""     
        self.assigned_times = ""        
        self.pending_times = ""         
        self.pending_duration = ""      
        self.row_attribut_list = [
            {"name": "id",                      "key_words": ['ID'],                            "default_index": 0},
            {"name": "work_item_type",          "key_words": ['工作项类型', '跟踪'],            "default_index": -1},
            {"name": "severity",                "key_words": ['严重程度', 'Severity'],          "default_index": -1},
            {"name": "priority",                "key_words": ['优先级'],                        "default_index": -1},
            {"name": "title",                   "key_words": ['标题', '主题'],                  "default_index": -1},
            {"name": "status",                  "key_words": ['状态'],                          "default_index": -1},
            {"name": "difficulty",              "key_words": ['Difficulty Degree'],             "default_index": -1},
            {"name": "person_in_charge",        "key_words": ['负责人', '指派给'],                  "default_index": -1},
            {"name": "dead_line",               "key_words": ['计划完成日期','截止日期'],           "default_index": -1},
            {"name": "complete_time",           "key_words": ['实际完成时间', '结束日期'],          "default_index": -1},
            {"name": "bug_fix_cycle",           "key_words": ['缺陷修复周期'],                      "default_index": -1},
            {"name": "bug_fixed_duration",      "key_words": ['已修复-停留时间'],                   "default_index": -1},
            {"name": "estimated_workinghours",  "key_words": ['预估工时（小时）', '预估工时统计'],   "default_index": -1},
            {"name": "registered_workinghours", "key_words": ['已登记工时（小时）', '耗时'],        "default_index": -1},
            {"name": "rest_workinghours",       "key_words": ['剩余工时（小时）', '预期时间'],      "default_index": -1},
            {"name": "project_name",            "key_words": ['所属项目', '项目'],                   "default_index": -1},
            {"name": "creator",                 "key_words": ['创建者', '作者'],                    "default_index": -1},
            {"name": "creation_time",           "key_words": ['创建时间', '创建于'],                "default_index": -1},
            {"name": "reopen_times",            "key_words": ['REOPEN-停留次数'],                    "default_index": -1},
            {"name": "reopen_duration",         "key_words": ['REOPEN-停留时间'],                    "default_index": -1},
            {"name": "reopen_confirm_times",    "key_words": ['REOPEN_CONFIRM-停留次数'],             "default_index": -1},
            {"name": "reopen_confirm_duration", "key_words": ['REOPEN_CONFIRM-停留时间'],             "default_index": -1},
            {"name": "state_detail",            "key_words": ['State Details'],                     "default_index": -1},
            {"name": "evaluating_duration",     "key_words": ['EVALUATING-停留时间'],             "default_index": -1},
            {"name": "assigned_duration",       "key_words": ['ASSIGNED-停留时间'],                 "default_index": -1},
            {"name": "assigned_times",          "key_words": ['ASSIGNED-停留次数'],                "default_index": -1},
            {"name": "pending_times",           "key_words": ['PENDING-停留次数'],                "default_index": -1},
            {"name": "pending_duration",        "key_words": ['PENDING-停留时间'],                "default_index": -1}
        ]
        self.head_line = []
        self.row_index = {}
        self.is_first_row = True
        self.member_name = ""
        self._path = ""

kpi_row = KPIRow()


def date_diff_days(start_date, end_date):
    date_format = "%Y-%m-%d"
    f_start = start_date.replace("/","-")
    start_date = datetime.strptime(f_start, date_format)
    f_end = end_date.replace("/","-")
    end_date = datetime.strptime(f_end, date_format)

    # 计算相差天数
    delta = end_date - start_date
    return delta.days


def minutes_to_dhm(minutes):
    """
    :param minutes:
    :return: 2 天18 小时1 分钟
            3 小时44 分钟
            2 分钟
            6 天27 分钟
    """
    ret_str = ""
    days = int(minutes // 1440)
    if days != 0:
        ret_str += "{} 天".format(days)
    hours = int(minutes % 1440)
    hours = int(hours // 60)
    if hours != 0:
        ret_str += " {} 小时".format(hours)
    minutes = int(minutes % 60)
    if minutes != 0:
        ret_str += " {} 分钟".format(minutes)

    return ret_str


def dhm_to_minutes(dhm):
    """
    :param dhm: 2 天18 小时1 分钟
                3 小时44 分钟
                2 分钟
                6 天27 分钟
    :return: time in minutes
    """
    days = 0
    hours = 0
    minutes = 0
    d = dhm.find("天")
    if d != -1:
        days = dhm[:d].strip()
        print(f"days: {days}")
    h = dhm.find("小时")
    if h != -1:
        if d == -1:
            hours = dhm[:h].strip()
        else:
            hours = dhm[d + 1:h].strip()
        print(f"hours: {hours}")
    m = dhm.find("分钟")
    if m != -1:
        if h == -1 and d == -1:
            minutes = dhm[:m].strip()
        elif h == -1 and d != 1:
            minutes = dhm[d + 1:m].strip()
        else:
            minutes = dhm[h + 2:m].strip()
        print(f"minutes: {minutes}")

    return int(days) * 1440 + int(hours) * 60 + int(minutes)


class KPIItem:
    def __init__(self):
        self.name_list = []
        self.status_counter = {}
        self.total = 0
        self.total_complete = 0

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

    def do_status_count(self, st):
        i = st.find('（')
        if i != -1:
            st = st[:i]

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

        self.total_complete = c

        ratio = "{:.1f}".format(float(c / self.total) * 100.0)
        return " " * 4 + pre + str(ratio) + "%"

    def get_status_count(self, rt_list):
        d = self.status_counter
        c = 0
        for r in rt_list:
            c += d[r]

        return c
    
    def get_reopen_time(self, row_list):
        if row_attr_index["reopen_times"] != 0:
            rop_t = row_list[row_attr_index["reopen_times"]]
            print(f"kpi item reopen times: {rop_t}")
            if len(rop_t) != 0:
                return int(rop_t)

        return 0


class itemFAEBUG(KPIItem):
    def __init__(self):
        super().__init__()
        self.name_list = ["FAE_BUG"]
        self.status_counter = {
            "NO_FEEDBACK": 0,
            "NO_RESPONSE": 0,
            "RESOLVED": 0,
            "REOPEN": 0,
            "REOPEN_CONFIRM": 0,
            "WAIT_FEEDBACK": 0,
            "NEW-未开始": 0,
            "DOING-进行中": 0,
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

    def do_proc_fae(self, id_v, st, ty, row_list):
        self.total += 1
        super().do_status_count(st)
        # super().do_proc(id_v, st, ty)
        self.reopen_times += super().get_reopen_time(row_list)

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
            "REOPEN_CONFIRM": 0,
            "WAIT_FEEDBACK": 0,
            "NEW-未开始": 0,
            "DOING-进行中": 0,
            "ASSIGNED": 0,
            "PENDING": 0,
            "TESTING": 0,
            "REJECTED": 0,
            "WAIT_RELEASE": 0,
            "关闭": 0,
            "NO_FEEDBACK": 0,
            "EVALUATING": 0,
            "NO_RESPONSE": 0
        }
        self.fix_pre = "REQUIREMENT Completed: "
        self.reopen_pre = "REQUIREMENT Reopened: "
        self.summary = " " * 4 + self.fix_pre + "null\n"
        self.summary += " " * 4 + self.reopen_pre + "null"
        self.reopen_times = 0
        self.diff_days = 0
        self.in_time = 0
        self.out_time = 0

    def do_proc_req(self, id_v, st, ty, row_list):
        super().do_proc(id_v, st, ty)
        if row_attr_index["reopen_times"] == 0:
            rop_t = ""
        else:
            rop_t = row_list[row_attr_index["reopen_times"]]
        print(f"requirement rop_t: {rop_t}")
        if len(rop_t) != 0:
            self.reopen_times += int(rop_t)

        # complete in time
        if row_attr_index["complete_time"] == 0:
            cmp_t = ""
        else:
            cmp_t = row_list[row_attr_index["complete_time"]]
            if cmp_t != "":
                l = cmp_t.split()
                print(f"complete time: {l}, {cmp_t}")
                cmp_t = "".join(l[0])
        print(f"requirement cmp_t: {cmp_t}")

        # deadline
        if row_attr_index["dead_line"] == 0:
            dd_line = ""
        else:
            dd_line = row_list[row_attr_index["dead_line"]]
            
        print(f"requirement dd_line: {dd_line}")

        # diff_days = deadline - complete_time, < 0 timeout. >=0: in time
        if len(dd_line) != 0 and len(cmp_t) != 0:
            self.diff_days = date_diff_days(cmp_t, dd_line)
            if self.diff_days < 0:
                self.out_time += 1
            else:
                self.in_time += 1

    def calcu_summary(self):
        rt_list = ["NO_FEEDBACK", "RESOLVED", "REJECTED", "WAIT_RELEASE", "关闭", "WAIT_FEEDBACK", "NO_RESPONSE"]
        self.summary = super().rate_calculater(rt_list, self.fix_pre) + "\n"
        print(f"requirement total complete: {self.total_complete}/{self.total}")

        # COMPLETE IN TIME
        # rate of complete in time = in_time/total_complete x 100%
        if self.in_time == 0:
            self.summary += "    REQUIREMENT complete in time: 0"
        else:
            rate = "{:.1f}%".format(float(self.in_time / self.total_complete) * 100.0)
            self.summary += "    REQUIREMENT complete in time({}/{}): {}".format(self.in_time, self.total_complete, rate)
        self.summary += "\n"
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
            "ASSIGNED-已分配": 0,
            "PENDING-暂停": 0,
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
            "REOPEN": 0,
            "验证中": 0
        }
        self.reopen_times = 0
        self.fix_pre = "ST_BUG Fixed: "
        self.reopen_pre = "ST_BUG Reopened: "
        self.summary = " " * 4 + self.fix_pre + "null\n"
        self.summary += " " * 4 + self.reopen_pre + "null"
        self.redmin_bug = {"Critical": 0, "Major": 0, "Normal": 0}
        self.fixed = {"Critical": 0, "Major": 0, "Normal": 0}
        self.reopened = {"Critical": 0, "Major": 0, "Normal": 0}
        self.finish_time = {"Critical": 0, "Major": 0, "Normal": 0}
        self.diff_total = {"Critical": 0, "Major": 0, "Normal": 0}

    def do_proc_req(self, id_v, st, ty, row_list):
        super().do_proc(id_v, st, ty)
        # severity
        if row_attr_index["severity"] == 0:
            severity = ""
        else:
            severity = row_list[row_attr_index["severity"]]
            print(f"st_bug severity: {severity}, value:{self.diff_total[severity]}")
            self.diff_total[severity] += 1

        # diff days
        bug_ft_mins = 0
        r = row_attr_index["bug_fix_cycle"]
        print(f"st_bug bug_fix_cycle: {r}/{len(row_list)}")
        if r != 0:
            bug_ft = row_list[r]

            if bug_ft != "":
                bug_ft_mins = dhm_to_minutes(bug_ft)
                print(f"st_bug bug_ft: {bug_ft}")
        else:
            self.redmin_bug[severity] += 1

        print(f"st_bug bug_fix_cycle: {bug_ft_mins}")
        self.finish_time[severity] += bug_ft_mins

        rop_t = super().get_reopen_time(row_list)
        self.reopen_times += rop_t
        self.reopened[severity] += rop_t

        if st == 'REOPENED' or st == 'REOPEN':
            self.reopen_times += 1
            self.reopened[severity] += 1

    def calcu_summary(self):
        rt_list = ["拒绝", "RESOLVED-已修复", "CLOSED-关闭", "验证中"]
        self.summary = super().rate_calculater(rt_list, self.fix_pre) + '\n'

        severity_list = ["Critical", "Major", "Normal"]
        avg_finish_time = ""
        reopened_time = ""
        for severity in severity_list:
            if self.diff_total[severity] == 0:
                continue
            total = self.diff_total[severity] - self.redmin_bug[severity]
            avg = int(self.finish_time[severity] / total)

            pre_fix = " " * 4 + severity.title() + " ST_BUG Average Finish Time:"
            avg_finish_time += pre_fix + " {} \n".format(minutes_to_dhm(avg))
            pre_fix = " " * 4 + severity.title() + " ST_BUG Reopened "
            reopened_time += pre_fix + "({}/{}) = {:.1f}%\n".format(self.reopened[severity], self.diff_total[severity],
                                                                     (self.reopened[severity] / self.diff_total[severity]) * 100)

        self.summary += avg_finish_time
        self.summary += reopened_time



class KPIForOnePerson:
    def reset(self, r, name):
        self.fae_bug.reset()
        self.requirement.reset()
        self.st_bug.reset()
        self.prot_dev.reset()

    def parse_kpi_row(self, kpi_row):
        work_type = kpi_row[row_attr_index["work_item_type"]]
        id_v = kpi_row[row_attr_index["id"]]
        st = kpi_row[row_attr_index["status"]]
        print("work type: ", work_type)

        self.oper_res = self.oper_result_dict['normal']

        if work_type in self.fae_bug.name_list:
            self.fae_bug.do_proc_fae(id_v, st, work_type, kpi_row)
        elif work_type in self.prot_dev.name_list:
            self.prot_dev.do_proc(id_v, st, work_type)
        elif work_type in self.st_bug.name_list:
            self.st_bug.do_proc_req(id_v, st, work_type, kpi_row)
        elif work_type in self.requirement.name_list:
            self.requirement.do_proc_req(id_v, st, work_type, kpi_row)
        else:
            raise Exception(f"unknown work type: {work_type}")

    def kpi_summary(self):
        self.fae_bug.calcu_summary()
        self.prot_dev.calcu_summary()
        self.st_bug.calcu_summary()
        self.requirement.calcu_summary()

    def __init__(self, name_en, name_cn):
        self.oper_result_dict = {
            'normal' : "正常",
            'revised' : "已修订",
            'ignore' : "忽略"
        }
        self.oper_res = ""
        self.name_en = name_en
        self.name_cn = name_cn
        self.work_type = ''
        self.work_load = ""
        self.work_hour = 0
        self.fae_bug = itemFAEBUG()
        self.requirement = itemREQUIREMENT()
        self.prot_dev = itemPROT_DEV()
        self.st_bug = itemST_BUG()
        self._path = ''
        # self.report = ""


def row_parser(row_list, first_row, members):
    if first_row:
        #init_row_index(row_list)
        kpi_row.init_row_index(row_list)
    else:
        kpi_row.pre_proc(row_list)
        name = kpi_row.get_name(row_list)
        member = None
        print(f"name: {name}")
        for mb in members:
            if mb.name_en.lower() == name:
                member = mb
                break
        
        if member is not None:
            member.parse_kpi_row(kpi_row)
            return member
        else:
            print(f"member {name} not found")
            return None


def kpi_process(r_path, csv_list, members):
    sep_kpi_fold = csv_list[0]
    sep_kpi_fold = sep_kpi_fold[sep_kpi_fold.find("_") + 1:sep_kpi_fold.rfind(".")]
    sep_kpi_fold = os.path.join(r_path, sep_kpi_fold)
    if not os.path.exists(sep_kpi_fold):
        os.makedirs(sep_kpi_fold)
    else:
        for mb in members:
            kpi_row.init_folder(sep_kpi_fold, mb.name_en)
    
    for csv_file in csv_list:
        fpath = os.path.join(r_path, csv_file)
        print(f"fpath: {fpath}, fn: {csv_file}")

        with open(fpath, 'r', encoding="utf-8-sig") as csv_f:
            # set the first row as True when open an new file
            kpi_row.is_first_row = True
            reader = csv.reader(csv_f)
            for r_list in reader:
                mb = kpi_row.pre_proc(r_list, members)
                if mb is not None:
                    # mb.parse_kpi_row(kpi_row)
                    kpi_row.save_kpi_row(mb)
                
                kpi_row.is_first_row = False

    #for mb in members:
    #    mb.kpi.kpi_summary()
