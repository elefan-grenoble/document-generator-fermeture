import datetime
import calendar

weekday = [
    'Lundi',
    'Mardi',
    'Mercredi',
    'Jeudi',
    'Vendredi',
    'Samedi',
    'Dimanche'
]

months = [
    'Janvier',
    'Février',
    'Mars',
    'Avril',
    'Mai',
    'Juin',
    'Juillet',
    'Août',
    'Septembre',
    'Octobre',
    'Novembre',
    'Décembre'
]


def datestring_fr(date):
    datestring = weekday[date.weekday()][0:3].upper()
    datestring += ' ' + str(date.strftime("%d.%m"))
    return datestring


def get_dates_for_month(year, month):
    num_days = calendar.monthrange(year, month)[1]
    days = [datestring_fr(datetime.date(year, month, day)) for day in range(1, num_days+1)]
    return days