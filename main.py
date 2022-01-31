# pyinstaller main.py --onefile -w -nDES.LM.Reports
import sys

import myconstants
import mainform
import reportcreater
from myutils import get_files_list


if __name__ == "__main__":
    curr_reporter = reportcreater.ReportCreater()
    app, ui, mainwindow = mainform.get_app_and_mainwindow()
    ui.setup_form(
        curr_reporter.get_reports_list(),
        get_files_list(reportcreater.get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)))

    mainwindow.show()
    sys.exit(app.exec_())

