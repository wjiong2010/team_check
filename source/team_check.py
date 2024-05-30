from codereview import cr_parse_result
from team import Team


software_develop_team = Team()


def main():
    software_develop_team.init_members()
    # software_develop_team.print_members()
    cr_parse_result("Code_Review.xml", software_develop_team)
    cr_parse_result("Code_Review_pro.xml", software_develop_team)
    software_develop_team.save_as_text("Code_Review.txt")


# 外部调用的时候不执行
if __name__ == '__main__':
    main()
