from docx import Document
from team import Team
from team import TeamMember
import datetime
from docx.enum.text import WD_ALIGN_PARAGRAPH


replace_list = [
    {'<<Name>>': '王炯'},
    {'<<Year>>': '2024'},
    {'<<Season>>': 'Q4'},
    {'<<S_Month>>': '222'},
    {'<<S_Day>>': ''},
    {'<<E_Month>>': ''},
    {'<<E_Day>>': ''},
    {'<<P_Score>>': ''},
    {'<<M_Score>>': ''},
    {'<<Key_Score>>': '0'},
    {'<<Total_Score>>': ''},
    {'<<Rank>>': ''},
    {'<<Level>>': ''},
    {'<<Comment>>': ''},
    {'<<Opinion>>': ''}
]


def update_doc_info(document):
    '''
    Update the document information.
    '''
    _temp_list = []
    for rp_dict in replace_list:
        for i,j in rp_dict.items():
            _temp_list.append(tuple([i, j]))
    print(_temp_list)
    
    document.core_properties.author = "John Wang"
    # doc.core_properties.category = "Queclink"
    # doc.core_properties.comments = "Queclink"
    # doc.core_properties.content_status = "Queclink"
    document.core_properties.created = datetime.datetime.now()
    # doc.core_properties.identifier = "Queclink"
    # doc.core_properties.keywords = "Queclink"
    document.core_properties.language = "zh-CN"
    document.core_properties.last_modified_by = "John Wang"
    # doc.core_properties.revision = 1
    # doc.core_properties.subject = "Queclink"
    # doc.core_properties.title = "Queclink"
    # doc.core_properties.version = "Queclink"
    table_count = 0
    for table in document.tables:
        table_count += 1
        for row in table.rows:
            for cell in row.cells:
                print(cell.text)
                for tl in _temp_list:
                    if tl[0] in cell.text:
                        cell.text = cell.text.replace(tl[0], tl[1])
                        print(f"Table Cell1 Replace {tl[0]} with {tl[1]}")
                for para in cell.paragraphs:
                    if table_count <= 2:
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for para in document.paragraphs:
        for tl in _temp_list:
            if tl[0] in para.text:
                para.text = para.text.replace(tl[0], tl[1])
                print(f"Replace {tl[0]} with {tl[1]}")


def build_docx(doc_path = "kpi_interview_form_template.docx"):
    '''     
    Build the KPI Interview Form.
    '''
    doc = Document(doc_path)  # Create a document object.

    try:
        update_doc_info(doc)

        # The following steps are just for testing. They can be commented.
        # doc.add_paragraph('Revision History', style="Queclink Title 01")
        # doc.add_paragraph('Introduction', style="Queclink Title 01")
        # p = doc.add_paragraph('This is a test text to use the original style.', style="Queclink text paragraph")
        # doc.add_paragraph('Reference', style="Queclink Title 02")

        doc.save("new_kpi_template.docx")
    finally:
        if doc:
            doc.save("new_kpi_template.docx")
