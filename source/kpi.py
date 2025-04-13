import os
import csv
from datetime import datetime
from team_check_util import season_date
from team_check_util import date_diff
from team_check_util import minutes_to_dhm
from team_check_util import dhm_to_minutes
from team_check_util import is_chinese
import xlrd
import xlwt


def get_csv_filename(r, fn):
    '''
    Get the csv file name.
    '''
    fn = fn.title()
    t_fn = fn.replace(" ", "_") + ".csv"
    return os.path.join(r, t_fn)


def get_kpi_csv_file_list(year, season):
    '''
    Get the KPI csv file list.
    '''
    kpi_row.ss_start, kpi_row.ss_end = season_date(year, season)
    
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
                #print("error!!")
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
            # print("name_index: {}, {}, {}".format(_name, _index, _tmp_v))
            exec("self." + _name + " = _tmp_v")
            
    def get_member_by_name(self, members, _name):
        """
        :param members: the members list
        :return: the member object
        """
        print("name: " + _name)
        if is_chinese(_name):
            _en = False
        else:
            _en = True
        
        member = None
        for mb in members:
            if _en and mb.name_en.lower() == _name.lower():
                member = mb
                break
            if not _en and mb.name_cn.encode() == _name.encode():
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
            print("c_time: " + self.complete_time + " ss_start: " + self.ss_start + " self.year: " + self.year)
            
            # ss_start - complete_time >= 0, means the case is finished before this season
            if date_diff(self.complete_time, self.ss_start) >= 0:
                self.remark = "ignore"
                print("remark as ignore2")
            else:
                print("remark as not ignore")

    def pre_proc(self, row_list):
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
            #return self.get_member_by_name(members, self.member_name)

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
        self.planned_start_date = ""      
        self.dead_line = ""             
        self.planned_finish_date = ""             
        self.complete_time = ""         
        self.bug_fix_cycle = ""         
        self.iteration = ""         
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
        self.evaluating_times = ""   
        self.assigned_duration = ""     
        self.assigned_times = ""        
        self.pending_times = ""         
        self.pending_duration = ""
        self.pending_times = ""         
        self.wait_feedback_times = ""
        self.wait_feedback_duration = ""         
        self.wait_release_duration = ""
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
            {"name": "planned_start_date",      "key_words": ['计划开始日期'],                      "default_index": -1},
            {"name": "dead_line",               "key_words": ['截止日期'],                          "default_index": -1},
            {"name": "planned_finish_date",     "key_words": ['计划完成日期'],                      "default_index": -1},
            {"name": "complete_time",           "key_words": ['实际完成时间', '结束日期'],          "default_index": -1},
            {"name": "bug_fix_cycle",           "key_words": ['缺陷修复周期'],                      "default_index": -1},
            {"name": "iteration",               "key_words": ['所属迭代'],                        "default_index": -1},
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
            {"name": "evaluating_times",        "key_words": ['EVALUATING-停留次数'],             "default_index": -1},
            {"name": "assigned_duration",       "key_words": ['ASSIGNED-停留时间'],                 "default_index": -1},
            {"name": "assigned_times",          "key_words": ['ASSIGNED-停留次数'],                "default_index": -1},
            {"name": "pending_duration",        "key_words": ['PENDING-停留时间'],                "default_index": -1},
            {"name": "pending_times",           "key_words": ['PENDING-停留次数'],                "default_index": -1},
            {"name": "wait_feedback_duration",  "key_words": ['WAIT_FEEDBACK-停留时间'],                "default_index": -1},
            {"name": "wait_feedback_times",     "key_words": ['WAIT_FEEDBACK-停留次数'],                "default_index": -1},
            {"name": "wait_release_duration",   "key_words": ['WAIT_RELEASE-停留时间'],                "default_index": -1},
            {"name": "wait_release_times",      "key_words": ['WAIT_RELEASE-停留次数'],                "default_index": -1},
        ]
        self.row_index = {}
        self.is_first_row = True
        self.member_name = ""
        self.path = ""
        self.ss_start = ""
        self.ss_end = ""
        self.year = ""

kpi_row = KPIRow()


class KPIItem:
    def __init__(self):
        self.name_list = []
        self.status_counter = {}
        self.total = 0
        self.total_complete = 0
        self.STATUS_UNKNOWN = -1
        self.RESOLVED = 0
        self.REOPEN = 1
        self.REOPEN_CONFIRM = 2
        self.WAIT_FEEDBACK = 3
        self.NEW = 4
        self.DOING = 5
        self.ASSIGNED = 6
        self.PENDING = 7
        self.TESTING = 8
        self.REJECTED = 9
        self.WAIT_RELEASE = 10
        self.CLOSED = 11
        self.NO_FEEDBACK = 12
        self.EVALUATING = 13
        self.NO_RESPONSE = 14
        self.FINISHED = 15
        self.UNCONFIRMED = 16
        self.LOG_REQ = 17
        self.REFUSED = 18
        self.WAIT_SOLVE = 19
        self.CHN_RESOLVED = 20
        self.CHN_RESOLVED_RDM = 21
        self.CHN_REOPEN = 22
        self.REOPEND = 23
        self.CHN_WAIT_SOLVE = 24
        self.CHN_UNSTART = 25
        self.CHN_WAIT_SOLEV_RDM = 26
        self.CHN_DOING = 27
        self.CHN_ASSIGNED = 28
        self.CHN_PAUSE = 29
        self.CHN_TESTING = 30
        self.CHN_CLOSED = 31
        self.CHN_FINISHED = 32
        self.CHN_UNCONFIRMED = 33
        self.CHN_REFUSED = 34
        self.status_dict = {
            "NO_FEEDBACK": self.NO_FEEDBACK,
            "NO_RESPONSE": self.NO_RESPONSE,
            "RESOLVED": self.RESOLVED,
            "已修复": self.CHN_RESOLVED,
            "RESOLVED（已解决）": self.CHN_RESOLVED_RDM,
            "REOPEN": self.REOPEN,
            "REOPENED": self.REOPEND,
            "重新打开": self.CHN_REOPEN,
            "REOPEN_CONFIRM": self.REOPEN_CONFIRM,
            "WAIT_FEEDBACK": self.WAIT_FEEDBACK,
            "NEW": self.NEW,
            "待处理": self.CHN_WAIT_SOLVE,
            "未开始": self.CHN_UNSTART,
            "WAIT_SOLVE（待解决）": self.CHN_WAIT_SOLEV_RDM,
            "DOING": self.DOING,
            "进行中": self.CHN_DOING,
            "ASSIGNED": self.ASSIGNED,
            "已分配": self.CHN_ASSIGNED,
            "PENDING": self.PENDING,
            "暂停": self.CHN_PAUSE,
            "TESTING": self.TESTING,
            "验证中": self.CHN_TESTING,
            "REJECTED": self.REJECTED,
            "WAIT_RELEASE": self.WAIT_RELEASE,
            "CLOSED": self.CLOSED,
            "关闭": self.CHN_CLOSED,
            "EVALUATING": self.EVALUATING,
            "FILED": self.FINISHED,
            "已完成": self.CHN_FINISHED,
            "UNCONFIRMED": self.UNCONFIRMED,
            "待讨论": self.CHN_UNCONFIRMED,
            "LOG_Req": self.LOG_REQ,
            "拒绝": self.CHN_REFUSED
        }
        self.status_id = self.STATUS_UNKNOWN
    
    def get_status_key(self, status_id):
        for k, v in self.status_dict.items():
            if status_id == v:
                return k
        return ""
    
    def get_status_id(self, status):
        try:
            return self.status_dict[status]
        except KeyError:
            return self.STATUS_UNKNOWN

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
                _stat = self.get_status_key(k)
                if _stat == "": continue
                ts += " " * 8 + _stat + ': ' + str(self.status_counter[k]) + '\n'
        # print(ts)
        return ts

    def do_status_count(self, status_id):
        if status_id in self.status_counter.keys():
            self.status_counter[status_id] += 1
        else:
            self.status_counter.update({status_id: 1})


    def rate_calculater(self, rt_list, pre, count=0):
        # print("rt_list: {}, stat: {}".format(str(rt_list), str(self.status_counter)))
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
        self.job_done_flags = [self.NO_FEEDBACK, self.RESOLVED, self.REJECTED, self.WAIT_RELEASE, self.CHN_CLOSED, self.CLOSED, self.NO_RESPONSE]
        self.status_counter = {
            self.NO_FEEDBACK: 0,
            self.NO_RESPONSE: 0,
            self.RESOLVED: 0,
            self.REOPEN: 0,
            self.REOPEN_CONFIRM: 0,
            self.WAIT_FEEDBACK: 0,
            self.NEW: 0,
            self.CHN_UNSTART: 0,
            self.DOING: 0,
            self.CHN_DOING: 0,
            self.ASSIGNED: 0,
            self.PENDING: 0,
            self.TESTING: 0,
            self.REJECTED: 0,
            self.WAIT_RELEASE: 0,
            self.CHN_CLOSED: 0,
            self.CLOSED: 0
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
        self.status_id = super().get_status_id(kpi_row.status)
        super().do_status_count(self.status_id)

        # reopen times
        self.reopen_times += int(kpi_row.reopen_times)

    def calcu_summary(self):
        self.summary = super().rate_calculater(self.job_done_flags, self.fix_pre) + "\n"

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
        self.job_done_flags = [self.NO_FEEDBACK, self.RESOLVED, self.REJECTED, self.WAIT_RELEASE, self.CHN_CLOSED, self.NO_RESPONSE]
        self.status_counter = {
            self.RESOLVED: 0,
            self.REOPEN: 0,
            self.REOPEN_CONFIRM: 0,
            self.WAIT_FEEDBACK: 0,
            self.NEW: 0,
            self.CHN_UNSTART: 0,
            self.DOING: 0,
            self.CHN_DOING: 0,
            self.ASSIGNED: 0,
            self.PENDING: 0,
            self.TESTING: 0,
            self.REJECTED: 0,
            self.WAIT_RELEASE: 0,
            self.CHN_CLOSED: 0,
            self.NO_FEEDBACK: 0,
            self.EVALUATING: 0,
            self.NO_RESPONSE: 0
        }
        self.fix_pre = "REQUIREMENT Completed: "
        self.reopen_pre = "REQUIREMENT Reopened: "
        self.summary = " " * 4 + self.fix_pre + "null\n"
        self.summary += " " * 4 + self.reopen_pre + "null"
        self.reopen_times = 0
        self.diff_minutes = 0   # total time consumed for all requirements
        self.in_time = 0
        self.out_time = 0
        self.out_time_list = []
        self.rejects = 0
        self.reject_list = []
        self._ddl = "" # deadline or planned finish date

    def _parser_in_out_time(self, deadline, plan_fin_date, complete_time):
        # complete time
        if complete_time != "":
            l = complete_time.split()
            print(f"complete time: {l}, {complete_time}")
            complete_time = "".join(l[0])
        print(f"requirement complete_time: {complete_time}")

        # deadline and planned finish date
        print(f"requirement dd_line: {deadline}, pf: {plan_fin_date}")
        if len(deadline) != 0:
            self._ddl = deadline
        elif len(plan_fin_date) != 0:
            self._ddl = plan_fin_date
        else:
            self._ddl = ""
            # for requirement without deadline, consider it as out time
            self.out_time += 1

        # date difference between deadline and complete_time
        # _diff_days = deadline - complete_time, < 0 timeout. >=0: in time
        if len(self._ddl) != 0 and len(complete_time) != 0:
            if self._ddl.find(":") == -1:
                self._ddl += " 23:59:59"
            if complete_time.find(":") == -1:
                complete_time += " 23:59:59"
            
            print(f"requirement _ddl: {self._ddl}, complete_time: {complete_time}")
            
            if date_diff(complete_time, self._ddl, "minutes") < 0:
                self.out_time += 1
                self.out_time_list.append(kpi_row.id)
            else:
                self.in_time += 1

    def _parser_time_consumed(self, kpi_row):
        '''
        Calculate the time consumed for all requirements.
        '''
        print(f"requirement status: {self.status_id}, _ddl: {self._ddl}")
        if self._ddl == "" or self.status_id not in self.job_done_flags:
            # skip the requirement without deadline or not in job done flags
            print(f"requirement {kpi_row.id} skip")
            return

        # date difference between creation_time and _ddl
        assert len(kpi_row.creation_time) != 0
        self.diff_minutes += date_diff(kpi_row.creation_time, self._ddl, "minutes")
        print(f"requirement diff_minutes: {self.diff_minutes}")

    def parser(self, kpi_row):
        # total count
        self.total += 1

        # status count
        self.status_id = super().get_status_id(kpi_row.status)
        super().do_status_count(self.status_id)
        
        if self.status_id == self.REJECTED:
            self.rejects += 1
            self.reject_list.append(kpi_row.id)

        # reopen times
        rop_t = kpi_row.reopen_times
        print(f"requirement rop_t: {rop_t}")
        if len(rop_t) != 0:
            self.reopen_times += int(rop_t)

        # calculate the in time or out time by deadline, planned finish date and complete time
        self._parser_in_out_time(kpi_row.dead_line, kpi_row.planned_finish_date, kpi_row.complete_time)
        
        # time consume
        self._parser_time_consumed(kpi_row)
            

    def calcu_summary(self):
        # COMPLETE rate
        self.summary = super().rate_calculater(self.job_done_flags, self.fix_pre) + "\n"
        print(f"requirement total complete: {self.total_complete}/{self.total}")

        # COMPLETE IN TIME
        # rate of complete in time = in_time/total_complete x 100%
        if self.in_time == 0 or self.total_complete == 0:
            self.summary += "    REQUIREMENT complete in time: 0"
        else:
            rate = "{:.1f}%".format(float(self.in_time / self.total_complete) * 100.0)
            self.summary += "    REQUIREMENT complete in time({}/{}): {}".format(self.in_time, self.total_complete, rate)
        self.summary += "  TIMEOUT: " + str(self.out_time_list) + " REJECTED: " + str(self.reject_list)
        self.summary += "\n"

        # "REOPEN"
        # reopen率 = 总reopen次数/REQUIREMENT总数 x 100%
        self.summary += super().rate_calculater([], self.reopen_pre, self.reopen_times) + "\n"

        if self.reopen_times == 0:
            self.summary += "    REQUIREMENT Reopened Times: 0"
        else:
            self.summary += "    REQUIREMENT Reopened Times: " + str(self.reopen_times)
            
        # time consumed
        aver_diff_mins = 0 if self.total == 0 or self.diff_minutes == 0 else int(self.diff_minutes / self.total)
        _time_consumed = minutes_to_dhm(aver_diff_mins)
        self.summary += "\n    REQUIREMENT Time Consumed in Average: " + _time_consumed


class itemPROT_DEV(KPIItem):
    def __init__(self):
        super().__init__()
        self.name_list = ["PROTOCOL", "HF_PROTOCOL", "DEVELOP", "任务", "子任务"]
        self.job_done_flags = [self.NO_FEEDBACK, self.RESOLVED, self.REJECTED, self.WAIT_RELEASE, self.CLOSED, self.CHN_CLOSED]
        self.status_counter = {
            self.NO_FEEDBACK: 0,
            self.RESOLVED: 0,
            self.CLOSED: 0,
            self.CHN_CLOSED: 0,
            self.WAIT_FEEDBACK: 0,
            self.NEW: 0,
            self.CHN_UNSTART: 0,
            self.DOING: 0,
            self.CHN_DOING: 0,
            self.ASSIGNED: 0,
            self.CHN_ASSIGNED: 0,
            self.PENDING: 0,
            self.CHN_PAUSE: 0,
            self.TESTING: 0,
            self.REJECTED: 0,
            self.WAIT_RELEASE: 0,
            self.FINISHED: 0,
            self.CHN_FINISHED: 0
        }
        self.fix_pre = "Protocol & Develop Complete: "
        self.summary = " " * 4 + self.fix_pre + "null"
    
    def parser(self, kpi_row):
        # total count
        self.total += 1
        
        # status count
        self.status_id = super().get_status_id(kpi_row.status)
        super().do_status_count(self.status_id)

    def calcu_summary(self):
        self.summary = super().rate_calculater(self.job_done_flags, self.fix_pre)


class itemST_BUG(KPIItem):
    def __init__(self):
        super().__init__()
        self.name_list = ["ST_BUG", "HF_ST_BUG", "缺陷"]
        self.job_done_flags = [self.CHN_REFUSED, self.RESOLVED, self.CHN_RESOLVED, self.CHN_RESOLVED_RDM, self.CLOSED, self.CHN_CLOSED, self.CHN_TESTING]
        self.status_counter = {
            self.NEW: 0,
            self.CHN_UNSTART: 0,
            self.UNCONFIRMED: 0,
            self.CHN_UNCONFIRMED: 0,
            self.CHN_WAIT_SOLEV_RDM: 0,
            self.RESOLVED: 0,
            self.CHN_RESOLVED: 0,
            self.LOG_REQ: 0,
            self.CHN_REFUSED: 0,
            self.CLOSED: 0,
            self.CHN_CLOSED: 0,
            self.REOPEN: 0,
            self.CHN_REOPEN: 0,
            self.REOPEND: 0,
            self.CHN_TESTING: 0,
            self.CHN_RESOLVED_RDM: 0
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
        self.status_id = super().get_status_id(kpi_row.status)
        super().do_status_count(self.status_id)
        
        # severity: critical, major, normal
        severity = kpi_row.severity
        print(f"name: {kpi_row.person_in_charge}, st_bug severity: {severity}, value:{self.diff_total[severity]}, id: {kpi_row.id}")
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
        self.summary = super().rate_calculater(self.job_done_flags, self.fix_pre) + '\n'

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

class KPIPerformance:
    
    def __init__(self):
        self.year = ""
        self.name_cn = ""
        self.customer_type = ""
        self.org = ""
        self.season = ""
        self.pm_score = 0
        self.supervisor_score = 0
        self.total_score = 0
        self.rank = 0
        self.level = ""
        self.comment = ""
        self.opinion = ""


kpi_perf = KPIPerformance()

class KPIForOnePerson:
    def reset(self):
        self.fae_bug.reset()
        self.requirement.reset()
        self.st_bug.reset()
        self.prot_dev.reset()

    def parse_kpi_row(self, kpi_row):
        if kpi_row.remark == "ignore":
            return
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
        self.perf = KPIPerformance()


def kpi_folder_init(r_path, csv_list, members, clear_folder = False):
    sep_kpi_fold = csv_list[0]
    sep_kpi_fold = sep_kpi_fold[sep_kpi_fold.find("_") + 1:sep_kpi_fold.rfind(".")]
    sep_kpi_fold = os.path.join(r_path, sep_kpi_fold)
    if not os.path.exists(sep_kpi_fold):
        os.makedirs(sep_kpi_fold)
        return ''
    elif clear_folder:
        for mb in members:
            # clear if exist
            kpi_row.init_folder(sep_kpi_fold, mb.name_en)
    
    return sep_kpi_fold


def kpi_pre_process(r_path, csv_list, members, fmt):
    
    for csv_file in csv_list:
        fpath = os.path.join(r_path, csv_file)
        print(f"fpath: {fpath}, fn: {csv_file}")
        
        if not os.path.exists(fpath):
            print("File not found: " + fpath)
            continue

        with open(fpath, 'r', encoding="utf-8-sig") as csv_f:
            # set the first row as True when open an new file
            kpi_row.is_first_row = True
            reader = csv.reader(csv_f)
            for r_list in reader:
                kpi_row.pre_proc(r_list)
                mb = kpi_row.get_member_by_name(members, kpi_row.member_name)
                kpi_row.is_first_row = False
                if mb is None:
                    continue
                kpi_row.save_kpi_row(mb, fmt)


def kpi_analyze_process(r_path, members, mb_groups, fmt):
    for mb in members:
        if mb.group not in mb_groups: continue
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
            kpi_row.pre_proc(r_list)
            mb.kpi.parse_kpi_row(kpi_row)
            kpi_row.is_first_row = False

        mb.kpi.kpi_summary()
        
        csv_f.close()
        if csv_f.closed:
            print("File closed: " + csv_file)
        else:
            print("Failed to closed: " + csv_file)


def kpi_perf_row(row_list):
    """
    :param row_list: the row list ['用户名', '用户类型', '组织机构', '季度', 'PM得分', '主管得分', '总分']
    :param first_row: the first row flag
    :return: the row list
    """
    return row_list[0], row_list[4], row_list[5], row_list[6]


def kpi_interview_process(kpi_path, year, season, members, fmt):
    # parse the 2024Q4KPI考核成绩汇总.xls
    kpi_perf_file = os.path.join(kpi_path, year + season + "KPI考核成绩汇总.xls")
    kperf_book = xlrd.open_workbook(kpi_perf_file)
    src_sheet = kperf_book.sheet_by_index(0)

    first_row = True
    for row in range(1, src_sheet.nrows):
        row_list = src_sheet.row_values(row)
        if len(row_list) == 0:
            continue
        print("row_list: ", row_list)
        if first_row:
            first_row = False
            continue
        
        _name, _pm_score, _supervisor_score, _total_score = kpi_perf_row(row_list)
        print(f"row_list: {_name}, {_pm_score}, {_supervisor_score}, {_total_score}")
        mb = kpi_row.get_member_by_name(members, _name)
        if mb is None:
            print(f"member not found: {_name}")
            continue
        mb_perf = mb.kpi.perf
        mb_perf.name_cn = _name
        mb_perf.pm_score = _pm_score
        mb_perf.supervisor_score = _supervisor_score
        mb_perf.total_score = _total_score
        mb_perf.year = year
        mb_perf.season = season
        mb_perf.rank = ""
        mb_perf.level = ""
        mb_perf.comment = ""
        mb_perf.opinion = ""
    

def kpi_process(kpi_path, year, season, members, member_groups, option, fmt):
    if option == "kpi_interview":
        kpi_interview_process(kpi_path, year, season, members, fmt)
        return
    
    csv_list = get_kpi_csv_file_list(year, season)
    kpi_row.year = year
    
    if option != "kpi_analyze":
        kpi_folder_init(kpi_path, csv_list, members, True)
        kpi_pre_process(kpi_path, csv_list, members, fmt)

    if option != "kpi_pre":
        p = kpi_folder_init(kpi_path, csv_list, members)
        if p == '':
            print("No kpi files for analyze.")
            return
        kpi_row.ss_start, kpi_row.ss_end = season_date(year, season, "YYYY-MM-DD HH:MM:SS")
        kpi_analyze_process(p, members, member_groups, fmt)
