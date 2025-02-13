import os
import csv
from datetime import datetime
# from team import TeamMember

# GROUP_APPLICATION = TeamMember.GROUP_APPLICATION
#GROUP_SYSTEM = TeamMember.GROUP_SYSTEM


def get_csv_filename(r, fn):
    '''
    Get the csv file name.
    '''
    fn = fn.title()
    t_fn = fn.replace(" ", "_") + ".csv"
    return os.path.join(r, t_fn)


def season_date(season):
    '''
    Get the start and end date of the season.
    '''
    _season_dict = {
        "Q1": ["0101", "0331"],
        "Q2": ["0401", "0630"],
        "Q3": ["0701", "0930"],
        "Q4": ["1001", "1231"]
    }
    
    return _season_dict[season][0], _season_dict[season][1]


def get_kpi_csv_file_list(season):
    '''
    Get the KPI csv file list.
    '''
    kpi_row.ss_start, kpi_row.ss_end = season_date(season)
    
    return ["PMS_" + kpi_row.ss_start + "-" + kpi_row.ss_end + ".csv", "redmin_" + kpi_row.ss_start + "-" + kpi_row.ss_end + ".csv"]


class KPIRow:
    '''
    KPI row class.
    '''
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
        self.path = r
        p = get_csv_filename(r, name)
        print("exist file: " + p)
        if os.path.exists(p):
            os.remove(p)
            print("Clear exist file: " + p)

    def get_kpi_row(self, head_line):
        """
        :return: the kpi row
        """
        final_row = []
        for attr in self.row_attribut_list:
            if head_line:
                final_row.append(attr["key_words"][0])
                continue
            _name = attr["name"]
            exec("final_row.append(self." + _name + ")")
        return final_row
    
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
            except ValueError:
                print("error!!")
                if len(_keys) > 1:
                    _index = row.index(_keys[1])
                else:
                    _index = _def
            print("{}: {}".format(_name, _index))
            self.row_index.update({_name: _index})
        
        print(self.row_index)
    
    def save_csv(self):
        '''
        Save the KPI row to csv file.
        '''
        csv_file = get_csv_filename(self.path, self.member_name)
        if not os.path.exists(csv_file):
            # if the file does not exist, create a new file and write the head line
            _head_line = self.get_kpi_row(True)
        else:
            _head_line = []

        row = self.get_kpi_row(False)
        print(f"final_row: {row}")

        with open(csv_file, 'a', encoding="utf-8-sig") as csv_f:
            writer = csv.writer(csv_f, quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
            if len(_head_line) != 0:
                writer.writerow(_head_line)
            writer.writerow(row)

    def save_kpi_row(self, mb, fmt = "csv"):
        '''
        Save the KPI row as given format.
        '''
        if fmt == "csv":
            self.save_csv()
        else:
            print("save as text or push to database")

    def split_row_value(self, row_list):
        """
        :param row_list: the row list
        :return: the value of each item in the row list
        """
        for attr in self.row_attribut_list:
            _name = attr["name"]
            _index = self.row_index[_name]
            if -1 == _index:
                _tmp_v = ""
            else:
                _tmp_v = row_list[_index]
            print("name_index: {}, {}, {}".format(_name, _index, _tmp_v))
            exec("self." + _name + " = _tmp_v")
            
    def get_member_by_name(self, members):
        """
        :param members: the members list
        :return: the member object
        """
        member = None
        for mb in members:
            if mb.name_en.lower() == self.member_name:
                member = mb
                break
        return member
    
    def case_valid_check(self):
        """
        If the status of a task record is only modified and there is no substantive change, ignore it
        remark the ignored case as "ignore"
        """
        # if a case is finished before this season, ignore it        
        if self.complete_time != "":
            c = self.complete_time.split()
            if c[0].find("-") != -1:
                _time = c[0].split("-")
            elif c[0].find("/") != -1:
                _time = c[0].split("/")
            else:
                raise ValueError("Invalid date format")

            c_time = "{:02d}{:02d}".format(int(_time[1]), int(_time[2]))
            print("c_time: " + c_time + " y: " + _time[0] + " self.year: " + self.year)
            if c_time < self.ss_start and _time[0] <= self.year:
                self.remark = "ignore"

    def pre_proc(self, row_list, members):
        '''
        Pre-process the row list.
        '''
        if self.is_first_row:
            self.init_row_index(row_list)
            return None
        else:
            # split the row into different items and save them to the corresponding parameters
            print("pre_proc:   " + str(row_list))
            self.split_row_value(row_list)
            
            # valid check, 
            self.case_valid_check()
            
            # get member from members by name
            self.member_name = self.get_name(row_list)
            return self.get_member_by_name(members)

    def __init__(self):
        self.remark = ""
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
            {"name": "remark",                  "key_words": ['评注'],                          "default_index": -1},
            {"name": "id",                      "key_words": ['ID'],                            "default_index": 0},
            {"name": "work_item_type",          "key_words": ['工作项类型', '跟踪'],            "default_index": -1},
            {"name": "severity",                "key_words": ['严重程度', 'Severity'],          "default_index": -1},
            {"name": "priority",                "key_words": ['优先级'],                        "default_index": -1},
            {"name": "title",                   "key_words": ['标题', '主题'],                  "default_index": -1},
            {"name": "status",                  "key_words": ['状态'],                          "default_index": -1},
            {"name": "difficulty",              "key_words": ['Difficulty Degree'],             "default_index": -1},
            {"name": "person_in_charge",        "key_words": ['负责人', '指派给'],                  "default_index": -1},
            {"name": "dead_line",               "key_words": ['截止日期', '计划完成日期'],           "default_index": -1},
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
        self.row_index = {}
        self.is_first_row = True
        self.member_name = ""
        self.path = ""
        self.ss_start = ""
        self.ss_end = ""
        self.year = ""

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
    if minutes == 0:
        return "0 分钟"
    
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
    if len(dhm) == 0:
        return 0

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

    def do_status_count(self, st):
        # remove the content in '（'
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

    def parser(self, kpi_row):
        # total count
        self.total += 1
        
        # count status
        super().do_status_count(kpi_row.status)
        
        # reopen times
        self.reopen_times += int(kpi_row.reopen_times)

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

    def parser(self, kpi_row):
        # total count
        self.total += 1

        # status count
        super().do_status_count(kpi_row.status)
        
        # reopen times
        rop_t = kpi_row.reopen_times
        print(f"requirement rop_t: {rop_t}")
        if len(rop_t) != 0:
            self.reopen_times += int(rop_t)

        # complete in time
        cmp_t = kpi_row.complete_time
        if cmp_t != "":
            l = cmp_t.split()
            print(f"complete time: {l}, {cmp_t}")
            cmp_t = "".join(l[0])
        print(f"requirement cmp_t: {cmp_t}")

        # deadline
        dd_line = kpi_row.dead_line
        print(f"requirement dd_line: {dd_line}")

        # diff_days = deadline - complete_time, < 0 timeout. >=0: in time
        if len(dd_line) != 0 and len(cmp_t) != 0:
            self.diff_days = date_diff_days(cmp_t, dd_line)
            if self.diff_days < 0:
                self.out_time += 1
            else:
                self.in_time += 1
        # for requirement without deadline, consider it as out time
        elif len(dd_line) == 0:
            self.out_time += 1

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
    
    def parser(self, kpi_row):
        # total count
        self.total += 1
        
        # status count
        super().do_status_count(kpi_row.status)

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

    def parser(self, kpi_row):
        # total count
        self.total += 1

        # item type
        if kpi_row.work_item_type != "缺陷":
            is_redmin = True
        else:
            is_redmin = False
        
        # status
        super().do_status_count(kpi_row.status)
        
        # severity: critical, major, normal
        severity = kpi_row.severity
        print(f"st_bug severity: {severity}, value:{self.diff_total[severity]}")
        self.diff_total[severity] += 1

        # diff days
        bug_ft_mins = 0
        if not is_redmin:
            bug_ft_mins = dhm_to_minutes(kpi_row.bug_fix_cycle)
            print(f"st_bug bug_fix_cycle: {kpi_row.bug_fix_cycle}")
        else:
            self.redmin_bug[severity] += 1

        print(f"st_bug bug_fix_cycle: {bug_ft_mins}")
        self.finish_time[severity] += bug_ft_mins

        # reopen times
        rop_t = kpi_row.reopen_times
        if rop_t != "":
            self.reopen_times += int(rop_t)
            self.reopened[severity] += int(rop_t)

        if kpi_row.status == 'REOPEN':
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
            if total != 0:
                avg = int(self.finish_time[severity] / total)
            else:
                avg = 0

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
        if kpi_row.is_first_row:
            return

        work_type = kpi_row.work_item_type
        print("work type: ", work_type)

        kpi_row.remark = self.oper_result_dict['normal']

        if work_type in self.fae_bug.name_list:
            self.fae_bug.parser(kpi_row)
        elif work_type in self.prot_dev.name_list:
            self.prot_dev.parser(kpi_row)
        elif work_type in self.st_bug.name_list:
            self.st_bug.parser(kpi_row)
        elif work_type in self.requirement.name_list:
            self.requirement.parser(kpi_row)
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
        self.path = ''
        # self.report = ""


def kpi_folder_init(r_path, csv_list, members, clear_folder = False):
    sep_kpi_fold = csv_list[0]
    sep_kpi_fold = sep_kpi_fold[sep_kpi_fold.find("_") + 1:sep_kpi_fold.rfind(".")]
    sep_kpi_fold = os.path.join(r_path, sep_kpi_fold)
    if not os.path.exists(sep_kpi_fold):
        os.makedirs(sep_kpi_fold)
        return ''
    elif clear_folder:
        for mb in members:
            kpi_row.init_folder(sep_kpi_fold, mb.name_en)
    
    return sep_kpi_fold


def kpi_pre_process(r_path, csv_list, members, fmt):
    
    for csv_file in csv_list:
        fpath = os.path.join(r_path, csv_file)
        print(f"fpath: {fpath}, fn: {csv_file}")

        with open(fpath, 'r', encoding="utf-8-sig") as csv_f:
            # set the first row as True when open an new file
            kpi_row.is_first_row = True
            reader = csv.reader(csv_f)
            for r_list in reader:
                mb = kpi_row.pre_proc(r_list, members)
                if mb is None:
                    kpi_row.is_first_row = False
                    continue
                kpi_row.save_kpi_row(mb, fmt)


def kpi_analyze_process(r_path, members, fmt):
    for mb in members:
        if mb.group != 'application_development_group': continue
        csv_file = get_csv_filename(kpi_row.path, mb.name_en)
        csv_file = os.path.join(r_path, csv_file)
        
        try:
            csv_f = open(csv_file, 'r', encoding="utf-8-sig")
        except FileNotFoundError:
            print("File not found: " + csv_file)
            continue

        kpi_row.is_first_row = True
        reader = csv.reader(csv_f)
        for r_list in reader:
            kpi_row.pre_proc(r_list, members)
            mb.kpi.parse_kpi_row(kpi_row)
            kpi_row.is_first_row = False
        
        mb.kpi.kpi_summary()


def kpi_process(kpi_path, year, season, members, option, fmt):
    csv_list = get_kpi_csv_file_list(season)
    kpi_row.year = year
    
    if option != "kpi_analyze":
        kpi_folder_init(kpi_path, csv_list, members, True)
        kpi_pre_process(kpi_path, csv_list, members, fmt)

    if option != "kpi_pre":
        p = kpi_folder_init(kpi_path, csv_list, members)
        if p == '':
            print("No kpi files for analyze.")
            return
        kpi_analyze_process(p, members, fmt)
