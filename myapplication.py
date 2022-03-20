import sys

import myconstants
import mainform
import reportcreater
import myutils


class MyApplication:
    def __init__(self):
        self._curr_reporter = reportcreater.ReportCreater()
        
        self._mainwindow = mainform.MyWindow()
        
        self._mainwindow.ui.setup_form(self._curr_reporter.get_reports_list(),
                                        myutils.get_files_list(reportcreater.get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)))

        self._mainwindow.show()
        sys.exit(self._mainwindow._app.exec_())
