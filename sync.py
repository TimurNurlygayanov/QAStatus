import gspread
from datetime import date
from dateutil import relativedelta, parser
from dateutil.parser import parse
from jira.client import JIRA
from launchpadlib.launchpad import Launchpad


def get_date(parameter):
    return parser.parse(parameter.ctime())

EMAIL = 'email@google.com'
PASSWORD = 'secure_password'

## This is Google Spredsheet with v2 API (we can not modify docs with V3 API)
DOCUMENT_NAME = 'Input Table For Mirantis OpenStack Product QA Status'

PROJECTS = [{'name': 'murano', 'start_cell': 6},
            {'name': 'sahara', 'start_cell': 19},
            {'name': 'mistral', 'start_cell': 32},
            {'name': 'ceilometer', 'start_cell': 45},
            {'name': 'magnetodb', 'start_cell': 58},
            {'name': 'heat', 'start_cell': 71},
            {'name': 'rally', 'start_cell': 84},
            {'name': 'fuel', 'start_cell': 97}]

months_ago = date.today() - relativedelta.relativedelta(months=1)
week_ago = date.today() - relativedelta.relativedelta(weeks=1)

gs = gspread.login(EMAIL, PASSWORD)
wks = gs.open(DOCUMENT_NAME).sheet1

for pr in PROJECTS:
    launchpad = Launchpad.login_with(pr['name'], 'production')
    project = launchpad.projects[pr['name']]

    launchpad_bugs = project.searchTasks(importance=["Critical",],
                                         status=["New", "Fix Committed",
                                                 "Won't Fix", "Confirmed",
                                                 "Triaged", "In Progress",
                                                 "Incomplete"])

    wks.update_acell("D{0}".format(pr['start_cell'] + 4),
                     str(len(launchpad_bugs)))

    launchpad_bugs = project.searchTasks(status=["New", "Won't Fix",
                                                 "Confirmed", "Triaged",
                                                 "In Progress","Incomplete"])

    wks.update_acell("D{0}".format(pr['start_cell'] + 5),
                     str(len(launchpad_bugs)))

    launchpad_bugs = project.searchTasks(status=["New", "Fix Committed",
                                                 "Won't Fix", "Confirmed",
                                                 "Triaged", "In Progress",
                                                 "Incomplete"])

    wks.update_acell("D{0}".format(pr['start_cell'] + 6),
                     str(len(launchpad_bugs)))

    fixed_on_the_last_month = 0
    created_on_the_last_month = 0
    fixed_on_the_last_week = 0
    created_on_the_last_week = 0
    for bug in launchpad_bugs:
        if get_date(months_ago) < get_date(bug.date_created):
            created_on_the_last_month += 1
        if get_date(week_ago) < get_date(bug.date_created):
            created_on_the_last_week += 1
        if bug.date_fix_committed is not None:
            if get_date(months_ago) < get_date(bug.date_fix_committed):
                fixed_on_the_last_month += 1
            if get_date(week_ago) < get_date(bug.date_fix_committed):
                fixed_on_the_last_week += 1

    wks.update_acell("D{0}".format(pr['start_cell']),
                     str(created_on_the_last_week))
    wks.update_acell("D{0}".format(pr['start_cell'] + 1),
                     str(fixed_on_the_last_week))
    wks.update_acell("D{0}".format(pr['start_cell'] + 2),
                     str(created_on_the_last_month))
    wks.update_acell("D{0}".format(pr['start_cell'] + 3),
                     str(fixed_on_the_last_month))
