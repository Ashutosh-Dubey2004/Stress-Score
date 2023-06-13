import os
DEBUG_MODE = False
DBNAME      = "entries.db"
DBLOC       = f"{os.environ['USERPROFILE']}\\.survey-sheets\\"
DB          = DBLOC + DBNAME
REPORTSFOL  = f"{os.environ['USERPROFILE']}\\.survey-sheets\\reports"
RAWFOL      = f"{os.environ['USERPROFILE']}\\.survey-sheets\\raw"
TEMPDIR     = f"{os.environ['USERPROFILE']}\\.survey-sheets\\TEMP"
LOGDIR      = f"{os.environ['USERPROFILE']}\\.survey-sheets\\logs"
LOGFILE     = f"{LOGDIR}\\simple-log.txt"