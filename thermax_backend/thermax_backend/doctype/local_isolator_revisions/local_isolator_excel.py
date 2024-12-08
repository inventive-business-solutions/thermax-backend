from openpyxl import Workbook

def create_design_basis_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Local Isolator"
    ws.append(["Local Isolator Revisions"])
    # below line needs to be changed
    ws.append(["Project OC Number", "Project Name", "Project Owner", "Approver", "Review Status", "Review Date", "Review Comments", "Review Attachments"])
    wb.save("Local Isolator Revisions.xlsx")
    return "Local Isolator Excel created successfully"