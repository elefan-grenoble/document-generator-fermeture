import datetime, calendar
import locale

def set_locale(locale_):
    locale.setlocale(category=locale.LC_ALL, locale=locale_)

def datestring_fr(date):
    set_locale('fr_FR.utf8')
    datestring = calendar.day_name[date.weekday()][0:3].upper()
    datestring += ' ' + str(date.strftime("%d.%m"))
    return datestring

def get_dates_for_month(year, month):
    num_days = calendar.monthrange(year, month)[1]
    days = [datestring_fr(datetime.date(year, month, day)) for day in range(1, num_days+1)]
    return days