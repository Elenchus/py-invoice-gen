import os
from datetime import date, time
from comtypes import client
import pandas as pd
from mailmerge import MailMerge

def create_invoice(timesheet, invoice_no=1, rate=20, template='./assets/InvoiceTemplate.docx'):
    ()
    dte = '{:%d-%b-%Y}'.format(date.today())
    data = pd.read_excel(f"{timesheet}.xlsx", usecols=list(range(3)))
    data["Date"] = data["Date"].apply(lambda x: x.strftime('%d-%b-%Y'))
    data["Total"] = data["Hours"] * rate
    due = data["Total"].sum()
    hours = data["Hours"].sum()
    data = data.astype(str)

    with MailMerge(template) as document:
        document.merge(
            date=dte,
            hrs=str(hours),
            due=f"${due:.2f}",
            rate=f"${rate}/hour",
            inv_num=str(invoice_no)
        )

        rows = data.to_dict('records')
        document.merge_rows('Date', rows)
        word_file = f"{timesheet}_temp.docx"
        document.write(word_file)

        try:
            word = client.CreateObject('Word.Application')
            doc = word.Documents.Open(word_file)
            doc.SaveAs(f"{timesheet}.pdf", FileFormat=17)
        finally:
            try:
                doc.Close()
            except NameError:
                pass

            word.Quit()


        os.remove(word_file)

if __name__ == "__main__":
    import sys
    from argparse import ArgumentParser
    from pathlib import Path
    path_to_script = os.path.realpath(__file__)
    path_dir = os.path.dirname(path_to_script)
    path = Path(path_dir)
    defaults = {
        'timesheet': path / "assets/SampleTimesheet",
        'invoice_no': 1,
        'rate': 20,
        'template': path / "assets/InvoiceTemplate.docx"
    }

    parser = ArgumentParser(description="get inputs for create_invoice")
    for key, val in defaults.items():
        parser.add_argument(f"--{key}", metavar=key, type=type(val), default=val)

    args = parser.parse_args()
    create_invoice(**vars(args))