import os
from codereview import cr_parse_result
from kpi import kpi_process
from team import Team


software_develop_team = Team()


def main():
    season = 'Q2'
    path = os.getcwd()
    root_path = path.replace("source", "raw_data")
    root_path = os.path.join(root_path, season)
    print(f"root_path: {root_path}")

    software_develop_team.init_members()
    # software_develop_team.print_members()

    # cr_xml_file_list = ["Code_Review.xml", "Code_Review_pro.xml"]
    # cr_parse_result(root_path, cr_xml_file_list, software_develop_team)
    # # software_develop_team.save_as_text("cr_result", "Code_Review.txt")
    # software_develop_team.save_as_excel("cr_result", "Code_Review.xlsx")

    print("KPI processing start")
    kpi_csv_file_list = ["PMS_0101-0630.csv"]
    kpi_process(root_path, kpi_csv_file_list, software_develop_team.members)
    output_file = os.path.join(root_path, "2024Q2-REQ-BUG_Report.txt")
    software_develop_team.save_as_text("kpi", output_file)
    print("KPI processing end")


# 外部调用的时候不执行
if __name__ == '__main__':
    main()
