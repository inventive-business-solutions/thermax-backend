from openpyxl import Workbook

def create_design_basis_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Design Basis"
    ws.append(["Design Basis Revision History"])
    ws.append(["Project OC Number", "Project Name", "Project Owner", "Approver", "Review Status", "Review Date", "Review Comments", "Review Attachments"])
    wb.save("Design Basis Revision History.xlsx")
    return "Design Basis Excel created successfully"