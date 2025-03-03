import re
from datetime import datetime

season_cvt = {
    'Q1': '一',
    'Q2': '二',
    'Q3': '三',
    'Q4': '四'
}

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


def date_cvt(date):
    '''
    Convert the date from mmdd to mm/dd.
    '''
    print("Enter date:{}".format(date))
    pattern = r'^(\d{4}\/\d{1,2}\/\d{1,2})( \d{1,2})?(:\d{2})?(:\d{2})?$'
    date_match = re.match(pattern, date)
    if date_match is None:
        print("date format error! date:{}".format(date))
        return date
    
    print("date_match:{}, len: {}".format(date_match.groups(), len(date_match.groups())))
    if date_match.group(2) is None:
        date += " 00:00:00" # add hours, minutes and seconds if pattern
    elif date_match.group(3) is None:
        date += ":00:00" # add minutes and seconds if pattern
    elif date_match.group(4) is None:
        date += ":00" # add seconds if pattern
    else:
        print("date group:{}, {}, {}, {}".format(date_match.group(1), date_match.group(2), date_match.group(3), date_match.group(4)))

    return date.replace("/","-")


def date_diff(start_date, end_date, date_type = "days"):
    '''
    Calculate the difference between two dates.
    start_date: 2025/2/10  09:10:32
    end_date: 2024/3/29  00:00:00
    '''
    assert(len(start_date) != 0)
    assert(len(end_date) != 0)
    
    # format the date, convert the date from YYYY/MM/DD to YYYY-MM-DD
    f_start = date_cvt(start_date)
    f_end = date_cvt(end_date)
    print("f_start:{} f_end:{}".format(f_start, f_end))

    date_format = "%Y-%m-%d %H:%M:%S"
    start_date = datetime.strptime(f_start, date_format)
    end_date = datetime.strptime(f_end, date_format)

    # calculate the difference between two dates
    delta = end_date - start_date
    
    if date_type == "days":
        return delta.days
    elif date_type == "hours":
        return delta.days * 24 + delta.seconds // 3600
    elif date_type == "minutes":
        return delta.days * 24 * 60 + delta.seconds // 60
    else: # seconds
        return delta.seconds


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


def is_chinese(s):
    pattern = re.compile(r'^[\u4e00-\u9fa5]+$')
    if pattern.match(s):
        return True
    else:
        return False