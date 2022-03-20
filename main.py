# pyinstaller main.py --onefile -w -nDES.LM.Reports
import myapplication as myapp

app_handle = None

if __name__ == "__main__":
    app_handle = myapp.MyApplication()
