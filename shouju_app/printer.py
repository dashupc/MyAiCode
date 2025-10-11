import win32api
import pdf_export

def print_pdf(data):
    pdf_export.export(data)
    filename = f"{data['receipt_no']}.pdf"
    win32api.ShellExecute(
        0,
        "print",
        filename,
        None,
        ".",
        0
    )
