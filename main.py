# pyinstaller main.py --onefile -w -nDES.LM.Reports
import datetime
import logging
import os
import myapplication as myapp
import myutils

if __name__ == "__main__":
    if os.path.isfile("debug.cfg"):
        debug_logs_path = os.path.join(myutils.get_home_dir(), "debug.logs")

        if not os.path.isdir(debug_logs_path):
            os.mkdir(debug_logs_path)

        cdt = datetime.datetime.now()
        creation_str = f"{cdt.year:04}-{cdt.month:02}-{cdt.day:02} {cdt.hour:02}-{cdt.minute:02}-{cdt.second:02}"
        log_file_name = os.path.join(debug_logs_path, f"debug {creation_str}.log")
        logging.basicConfig(filename=log_file_name, level=logging.DEBUG)
        logging.debug(' DEBUG mode is on.')

    app_handle = myapp.MyApplication()
