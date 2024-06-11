import os
from codereview import cr_parse_result
from kpi import kpi_process
from team import Team


software_develop_team = Team()


def main():
    path = os.getcwd()
    root_path = path.replace("source", "raw_data")
    print(f"root_path: {root_path}")

    software_develop_team.init_members()
    # software_develop_team.print_members()

    cr_xml_file_list = ["Code_Review.xml", "Code_Review_pro.xml"]
    cr_parse_result(root_path, cr_xml_file_list, software_develop_team)
    # software_develop_team.save_as_text("cr_result", "Code_Review.txt")
    software_develop_team.save_as_excel("cr_result", "Code_Review.xlsx")

    print("KPI processing start")
    kpi_csv_file_list = ["redmin_0401-0605.csv", "PMS_0401-0605.csv"]
    kpi_process(root_path, kpi_csv_file_list, software_develop_team.members)
    software_develop_team.save_as_text("kpi", "2024Q2-KPI_Report.txt")
    print("KPI processing end")


# 外部调用的时候不执行
if __name__ == '__main__':
    main()
