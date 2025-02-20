import os
import argparse
from kpi import kpi_process
from team import Team
from codereview import cr_parse_result
from database import data_base as db


software_develop_team = Team()


def args_init():
    '''
    Initialize the arguments for the teamcheck.
    when -op is cr, only --cr_date is necessary. the arguments should be like: python teamcheck.py --cr_date 1230 -op cr
    when -op is kpi_pre, kpi_analyze or kpi, --cr_date is not a necessary argument. the arguments should be like: python teamcheck.py -s Q4 -y 2024 -op kpi_pre
    '''
    parser = argparse.ArgumentParser(description='Arguments for teamcheck. Include KPI and Code Review')
    parser.add_argument('-s', '--season', type=str,  choices=['Q1','Q2','Q3','Q4'], help='KPI season')
    parser.add_argument('--cr_date', type=str, default='0000', help='date of code review, formate: mmdd')
    parser.add_argument('-y', '--year', type=str, default='0000', help='The year of the KPI season')
    parser.add_argument('-a', '--archive', type=str, choices=['csv', 'docx', 'database'], default='csv', help='How to archive the data')
    parser.add_argument('-t', '--type', type=str, choices=['cr','kpi_pre','kpi_analyze', 'kpi_interview', 'kpi'], 
                        help='There are 4 types: cr, kpi_pre, kpi_analyze, kpi' +
                        'cr: code review, kpi_pre: KPI preprocess, kpi_analyze: KPI analyze, kpi: Full KPI process')
    parser.add_argument('-rel', '--release', type=str, default='docx', help='Build the KPI Interview Form.')
    args = parser.parse_args()
    return args


def get_root_path():
    '''
    Get the root path of the project.
    '''
    path = os.getcwd()
    root_path = path.replace("source", "raw_data")
    return root_path


def code_review_process(r_path):
    '''
    Process the code review.
    '''
    cr_path = os.path.join(r_path, "cr_" + cr_week)
    cr_xml_file_list = ["Code_Review.xml", "Code_Review_pro.xml"]
    cr_parse_result(cr_path, cr_xml_file_list, software_develop_team)
    cr_xlsx = os.path.join(cr_path, "Code_Review_" + cr_week + ".xlsx")
    software_develop_team.save_as_excel("cr_result", cr_xlsx)


def team_kpi_process(kpi_path, year, season, option, archive):
    '''
    Process the team KPI.
    '''
    kpi_process(kpi_path, year, season, software_develop_team.members, option, archive)

    if option == 'kpi' or option == 'kpi_analyze':
        output_file = os.path.join(kpi_path, year + season + "-KPI_Report.txt")
        software_develop_team.save_as_text("kpi", output_file)


def main():
    '''
    Main function for teamcheck.
    '''
    args = args_init()
    print(f"season: {args.season}, cr_week: {args.cr_date}, year: {args.year}, type: {args.type}, archive: {args.archive}, release: {args.release}")
    root_path = get_root_path()
    
    software_develop_team.init_members()
    # db.init_prot_info()
    # db.mysql_proc("team_members", software_develop_team.member_db_list)
    # software_develop_team.print_members()
    work_type = ''
    work_path = ''
    match args.type:
        case 'cr':
            if args.cr_date == '0000':
                print("Please input the cr_date for code review")
            else:
                work_type = 'cr'
                code_review_process(root_path)

        case 'kpi_pre' | 'kpi_analyze' | 'kpi' | 'kpi_interview':
            if args.season == None or args.year == '0000':
                print("Please input the season and year for KPI")
            else:
                work_path = os.path.join(root_path, args.season)
                team_kpi_process(work_path, args.year, args.season, args.type, args.archive)
                
                if args.type == 'kpi_interview':
                    software_develop_team.team_preformance(args.year, args.season, work_path, 'docx')

        case _: 
            print("Please input the correct type: cr, kpi_pre, kpi_analyze, kpi")


# Do not execute when imported as a module
if __name__ == '__main__':
    main()
