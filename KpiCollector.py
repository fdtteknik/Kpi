import pypyodbc
import logging
import sys
import os

import datetime
from dateutil.relativedelta import *

# create logger
logger = logging.getLogger(os.path.basename(sys.argv[0]))
logger.setLevel(logging.DEBUG)

_path = "C:\\KPI\\"
#_path = "http://10.124.0.81/"
#_path = "http://10.124.0.81/"
#_path = "C:\\inetpub\\wwwroot\"

# KStalle mapping
# SELECT TOP (1000) * FROM [boka100].[dbo].[Kställe]


def _thekpi(_cursor):
    _targets = [('Dec-18',4501760),('Nov-18',3820593),('Oct-18',4361450),('Sep-18',4310718),('Aug-18',3406248),('Jul-18',2354235),
                ('Jun-18',3931161),('May-18',3806196),('Apr-18',3639404),('Mar-18',3833141),('Feb-18',2891047),('Jan-18',2765652),
                ('Dec-17',4501760),('Nov-17',3820593),('Oct-17',4361450),('Sep-17',4310718),('Aug-17',3406248),('Jul-17',2354235),
                ('Jun-17',3931161),('May-17',3806196),('Apr-17',3639404),('Mar-17',3833141),('Feb-17',2891047),('Jan-17',2765652),
                ('Dec-16',2765652),('Nov-16',2765652)]

    grand_result = []
    today = datetime.datetime.today()
    current_month = datetime.datetime(today.year, today.month, 1)
    for i in range(0, 12):
        end = current_month + relativedelta(years=0, months=-i+1)
        start = current_month + relativedelta(years=0, months=-i)
        str = "select [Fakt100].[dbo].fakthstk.kställe as avdelning,sum(beloppexkl) as belopp from [Fakt100].[dbo].fakthstk left join [Fakt100].[dbo].fakthsth on [Fakt100].[dbo].fakthsth.FakturaNr = [Fakt100].[dbo].fakthstk.FakturaNr and [Fakt100].[dbo].fakthsth.ordernr=[Fakt100].[dbo].fakthstk.ordernr where [Fakt100].[dbo].fakthsth.datum>='{}' and datum<'{}' group by [Fakt100].[dbo].fakthstk.kställe".format(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))
        _cursor.execute(str)
        dbl = _cursor.fetchall()
        result = [start.strftime("%b-%y")]
        salj = 0
        supp = 0
        kons = 0
        tekn = 0
        utv = 0
        squ = 0
        misc = 0
        for target in _targets:
            if result[0] == target[0]:
                target_value = target[1]
                for row in dbl:
                    if row[0] == '2':
                        salj = salj + row[1]
                    elif row[0] == '3':
                        supp = supp + row[1]
                    elif row[0] == '4':
                        utv = utv + row[1]
                    elif row[0] == '5':
                        tekn = tekn + row[1]
                    elif row[0] == '7':
                        kons = kons + row[1]
                    elif row[0] == '9':
                        squ = squ + row[1]
                    else:
                        misc = misc + row[1]

                result.append("%d, %d, %d, %d, %d, %d, %d, %d" % (misc, salj, supp, utv, tekn, kons, squ, target_value))

                logger.warning(result)
                grand_result.append(result)

    # Write the result to the kpi1.js file
    f = open(_path+ "thekpi.js", "w")

    part1 = """
    function drawTheKpiChart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'FDT', 'SAL', 'SUP', 'UTV', 'TEK', 'KON', 'SQU', 'BUDGET'],"""

    part2 = ""
    j = 0
    for e in reversed(grand_result):
        part2 = part2 + "["
        i = 0;
        for v in e:
            if i == 0:
                part2 = part2 + "'%s'" % v
            else:
                part2 = part2 + v
            if i < len(e) - 1:
                part2 = part2 + ", "
            i = i + 1
        part2 = part2 + "]"
        if j < len(grand_result) - 1:
            part2 = part2 + ','
        j = j + 1
    part2 = part2 + "]"
    part3 = """
    	);

        var options = {
        title : 'Resultat / FDT / Månad',
        titleTextStyle: { color: 'grey' },
	    width:2048,
	    height:1152,
        vAxis: {
            title: 'Månad',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        hAxis: {
            title: 'Sek',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        legend: {
            textStyle: { color: 'grey' }
        },
        backgroundColor: '#000',
        isStacked: true,
        seriesType: 'bars',
        series: {
            7: { type: 'line', color: 'red'}
        }
        };

        var chart = new google.visualization.ComboChart(document.getElementById('thekpi_chart_div'));
        chart.draw(data, options);
    }
    """
    print(part1, file=f)
    print(part2, file=f)
    print(part3, file=f)
    f.close()


#
# KPI 1
#
def _kpi1(_cursor):
    _targets = [('Dec-18',3911760),('Nov-18',2960593),('Oct-18',3501450),('Sep-18',3450718),('Aug-18',2716248),('Jul-18',2071735),
                ('Jun-18',3541161),('May-18',3133696),('Apr-18',2779404),('Mar-18',2973141),('Feb-18',2031047),('Jan-18',1945652),
                ('Dec-17',3911760),('Nov-17',2960593),('Oct-17',3501450),('Sep-17',3450718),('Aug-17',2716248),('Jul-17',2071735),
                ('Jun-17',3541161),('May-17',3133696),('Apr-17',2779404),('Mar-17',2973141),('Feb-17',2031047),('Jan-17',1945652),
                ('Dec-16',2263152),('Nov-16',1905652)]

    grand_result = []
    today = datetime.datetime.today()
    current_month = datetime.datetime(today.year, today.month, 1)
    for i in range(0, 12):
        end = current_month + relativedelta(years=0, months=-i+1)
        start = current_month + relativedelta(years=0, months=-i)
        str = "select [Fakt100].[dbo].fakthstk.kställe as avdelning,sum(beloppexkl) as belopp from [Fakt100].[dbo].fakthstk left join [Fakt100].[dbo].fakthsth on [Fakt100].[dbo].fakthsth.FakturaNr = [Fakt100].[dbo].fakthstk.FakturaNr and [Fakt100].[dbo].fakthsth.ordernr=[Fakt100].[dbo].fakthstk.ordernr where [Fakt100].[dbo].fakthsth.datum>='{}' and datum<'{}' and [Fakt100].[dbo].fakthstk.kställe = '5' AND [Fakt100].[dbo].fakthstk.ArtikelNr != 'Ö9100' GROUP BY [Fakt100].[dbo].fakthstk.kställe".format(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))

        _cursor.execute(str)
        dbl = _cursor.fetchall()
        result = [start.strftime("%b-%y")]
        total = 0
        for target in _targets:
            if result[0] == target[0]:
                target_value = target[1]
                for row in dbl:
                    total = total + row[1]
                result.append("%d, %d" % (total, target_value))
                logger.warning(result)
                grand_result.append(result)

    # Write the result to the kpi1.js file
    f = open(_path + "kpi1.js", "w")

    part1 = """
    function drawKpi1Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Teknik', 'Budget'],"""

    part2 = ""
    j = 0
    for e in reversed(grand_result):
        part2 = part2 + "["
        i = 0;
        for v in e:
            if i == 0:
                part2 = part2 + "'%s'" % v
            else:
                part2 = part2 + v
            if i < len(e) - 1:
                part2 = part2 + ", "
            i = i + 1
        part2 = part2 + "]"
        if j < len(grand_result) - 1:
            part2 = part2 + ','
        j = j + 1
    part2 = part2 + "]"
    part3 = """
    	);

        var options = {
        title : 'Resultat / Teknik NOT Konsult Ö9100 / Månad',
        titleTextStyle: { color: 'grey' },
	    width:620,
	    height:500,
        vAxis: {
            title: 'Månad',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        hAxis: {
            title: 'Sek',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        legend: {
            textStyle: { color: 'grey' }
        },    
        backgroundColor: 'black',
        seriesType: 'bars',
        series: {1: {type: 'line'}}
        };

        var chart = new google.visualization.ComboChart(document.getElementById('kpi1_chart_div'));
        chart.draw(data, options);
    }
    """
    print(part1, file=f)
    print(part2, file=f)
    print(part3, file=f)
    f.close()



#
# KPI 2
#
def _kpi2(_cursor):
    _targets = [('Dec-18',60000),('Nov-18',80000),('Oct-18',80000),('Sep-18',80000),('Aug-18',40000),('Jul-18',20000),
                ('Jun-18',40000),('May-18',80000),('Apr-18',80000),('Mar-18',80000),('Feb-18',80000),('Jan-18',60000),
                ('Dec-17',80000),('Nov-17',80000),('Oct-17',160000),('Sep-17',160000),('Aug-17',120000),('Jul-17',40000),
                ('Jun-17',120000),('May-17',160000),('Apr-17',160000),('Mar-17',160000),('Feb-17',160000),('Jan-17',120000),
                ('Dec-16',120000),('Nov-16',160000)]
    grand_result = []
    today = datetime.datetime.today()
    current_month = datetime.datetime(today.year, today.month, 1)
    for i in range(0, 12):
        end = current_month + relativedelta(years=0, months=-i+1)
        start = current_month + relativedelta(years=0, months=-i)
        str = "select [Fakt100].[dbo].fakthstk.kställe as avdelning,sum(beloppexkl) as belopp from [Fakt100].[dbo].fakthstk left join [Fakt100].[dbo].fakthsth on [Fakt100].[dbo].fakthsth.FakturaNr = [Fakt100].[dbo].fakthstk.FakturaNr and [Fakt100].[dbo].fakthsth.ordernr=[Fakt100].[dbo].fakthstk.ordernr where [Fakt100].[dbo].fakthsth.datum>='{}' and datum<'{}' and [Fakt100].[dbo].fakthstk.kställe = '5' AND [Fakt100].[dbo].fakthstk.ArtikelNr = 'Ö9100' GROUP BY [Fakt100].[dbo].fakthstk.kställe".format(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))

        _cursor.execute(str)
        dbl = _cursor.fetchall()
        result = [start.strftime("%b-%y")]
        total = 0
        for target in _targets:
            if result[0] == target[0]:
                target_value = target[1]
                for row in dbl:
                    total = total + row[1]
                result.append("%d, %d" % (total, target_value))
                logger.warning(result)
                grand_result.append(result)

    # Write the result to the kpi1.js file
    f = open(_path+"kpi2.js", "w")

    part1 = """
    function drawKpi2Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Teknik', 'Budget'],"""

    part2 = ""
    j = 0
    for e in reversed(grand_result):
        part2 = part2 + "["
        i = 0;
        for v in e:
            if i == 0:
                part2 = part2 + "'%s'" % v
            else:
                part2 = part2 + v
            if i < len(e) - 1:
                part2 = part2 + ", "
            i = i + 1
        part2 = part2 + "]"
        if j < len(grand_result) - 1:
            part2 = part2 + ','
        j = j + 1
    part2 = part2 + "]"
    part3 = """
    	);

        var options = {
        title : 'Resultat / Teknik Konsult Ö9100 / Månad',
        titleTextStyle: { color: 'grey' },
	    width:620,
	    height:500,
        vAxis: {
            title: 'Månad',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        hAxis: {
            title: 'Sek',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        legend: {
            textStyle: { color: 'grey' }
        },    
        backgroundColor: 'black',
        seriesType: 'bars',
        series: {1: {type: 'line'}}
        };

        var chart = new google.visualization.ComboChart(document.getElementById('kpi2_chart_div'));
        chart.draw(data, options);
    }
    """
    print(part1, file=f)
    print(part2, file=f)
    print(part3, file=f)
    f.close()

#
# KPI 3
#
def _kpi3(_cursor):
    _targets = [('Dec-18',40000),('Nov-18',40000),('Oct-18',40000),('Sep-18',40000),('Aug-18',40000),('Jul-18',40000),
                ('Jun-18',40000),('May-18',40000),('Apr-18',40000),('Mar-18',40000),('Feb-18',40000),('Jan-18',40000),
                ('Dec-17',40000),('Nov-17',40000),('Oct-17',40000),('Sep-17',40000),('Aug-17',40000),('Jul-17',40000),
                ('Jun-17',40000),('May-17',40000),('Apr-17',40000),('Mar-17',40000),('Feb-17',40000),('Jan-17',40000),
                ('Dec-16',40000),('Nov-16',40000)]
    grand_result = []
    today = datetime.datetime.today()
    current_month = datetime.datetime(today.year, today.month, 1)
    for i in range(0, 12):
        end = current_month + relativedelta(years=0, months=-i+1)
        start = current_month + relativedelta(years=0, months=-i)
        str = "select [Fakt100].[dbo].fakthstk.kställe as avdelning,sum(beloppexkl) as belopp from [Fakt100].[dbo].fakthstk left join [Fakt100].[dbo].fakthsth on [Fakt100].[dbo].fakthsth.FakturaNr = [Fakt100].[dbo].fakthstk.FakturaNr and [Fakt100].[dbo].fakthsth.ordernr=[Fakt100].[dbo].fakthstk.ordernr where [Fakt100].[dbo].fakthsth.datum>='{}' and datum<'{}' and [Fakt100].[dbo].fakthstk.kställe = '2'  GROUP BY [Fakt100].[dbo].fakthstk.kställe".format(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))

        _cursor.execute(str)
        dbl = _cursor.fetchall()
        result = [start.strftime("%b-%y")]
        total = 0
        for target in _targets:
            if result[0] == target[0]:
                target_value = target[1]
                for row in dbl:
                    total = total + row[1]
                result.append("%d, %d" % (total, target_value))
                logger.warning(result)
                grand_result.append(result)

    # Write the result to the kpi1.js file
    f = open(_path+"kpi3.js", "w")

    part1 = """
    function drawKpi3Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Sälj', 'Budget'],"""

    part2 = ""
    j = 0
    for e in reversed(grand_result):
        part2 = part2 + "["
        i = 0;
        for v in e:
            if i == 0:
                part2 = part2 + "'%s'" % v
            else:
                part2 = part2 + v
            if i < len(e) - 1:
                part2 = part2 + ", "
            i = i + 1
        part2 = part2 + "]"
        if j < len(grand_result) - 1:
            part2 = part2 + ','
        j = j + 1
    part2 = part2 + "]"
    part3 = """
    	);

        var options = {
        title : 'Resultat / Sälj / Månad',
        titleTextStyle: { color: 'grey' },
	    width:620,
	    height:500,
        vAxis: {
            title: 'Månad',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        hAxis: {
            title: 'Sek',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        legend: {
            textStyle: { color: 'grey' }
        },    
        backgroundColor: 'black',
        seriesType: 'bars',
        series: {1: {type: 'line'}}
        };

        var chart = new google.visualization.ComboChart(document.getElementById('kpi3_chart_div'));
        chart.draw(data, options);
    }
    """
    print(part1, file=f)
    print(part2, file=f)
    print(part3, file=f)
    f.close()

#
# KPI 4
# Konsult
#
def _kpi4(_cursor):
    _targets = [('Dec-18',200000),('Nov-18',300000),('Oct-18',300000),('Sep-18',300000),('Aug-18',112500),('Jul-18',0),
                ('Jun-18',112500),('May-18',300000),('Apr-18',300000),('Mar-18',300000),('Feb-18',300000),('Jan-18',300000),
                ('Dec-17',200000),('Nov-17',300000),('Oct-17',300000),('Sep-17',300000),('Aug-17',112500),('Jul-17',0),
                ('Jun-17',112500),('May-17',300000),('Apr-17',300000),('Mar-17',300000),('Feb-17',300000),('Jan-17',300000),
                ('Dec-16',112500),('Nov-16',300000)]
    grand_result = []
    today = datetime.datetime.today()
    current_month = datetime.datetime(today.year, today.month, 1)
    for i in range(0, 12):
        end = current_month + relativedelta(years=0, months=-i+1)
        start = current_month + relativedelta(years=0, months=-i)
        str = "select [Fakt100].[dbo].fakthstk.kställe as avdelning,sum(beloppexkl) as belopp from [Fakt100].[dbo].fakthstk left join [Fakt100].[dbo].fakthsth on [Fakt100].[dbo].fakthsth.FakturaNr = [Fakt100].[dbo].fakthstk.FakturaNr and [Fakt100].[dbo].fakthsth.ordernr=[Fakt100].[dbo].fakthstk.ordernr where [Fakt100].[dbo].fakthsth.datum>='{}' and datum<'{}' and [Fakt100].[dbo].fakthstk.kställe = '7' GROUP BY [Fakt100].[dbo].fakthstk.kställe".format(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))

        _cursor.execute(str)
        dbl = _cursor.fetchall()
        result = [start.strftime("%b-%y")]
        total = 0
        for target in _targets:
            if result[0] == target[0]:
                target_value = target[1]
                for row in dbl:
                    total = total + row[1]
                result.append("%d, %d" % (total, target_value))
                logger.warning(result)
                grand_result.append(result)

    # Write the result to the kpi1.js file
    f = open(_path+"kpi4.js", "w")

    part1 = """
    function drawKpi4Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Konsult', 'Budget'],"""

    part2 = ""
    j = 0
    for e in reversed(grand_result):
        part2 = part2 + "["
        i = 0;
        for v in e:
            if i == 0:
                part2 = part2 + "'%s'" % v
            else:
                part2 = part2 + v
            if i < len(e) - 1:
                part2 = part2 + ", "
            i = i + 1
        part2 = part2 + "]"
        if j < len(grand_result) - 1:
            part2 = part2 + ','
        j = j + 1
    part2 = part2 + "]"
    part3 = """
    	);

        var options = {
        title : 'Resultat / Konsult / Månad',
        titleTextStyle: { color: 'grey' },
	    width:620,
	    height:500,
        vAxis: {
            title: 'Månad',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        hAxis: {
            title: 'Sek',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        legend: {
            textStyle: { color: 'grey' }
        },    
        backgroundColor: 'black',
        seriesType: 'bars',
        series: {1: {type: 'line'}}
        };

        var chart = new google.visualization.ComboChart(document.getElementById('kpi4_chart_div'));
        chart.draw(data, options);
    }
    """
    print(part1, file=f)
    print(part2, file=f)
    print(part3, file=f)
    f.close()

#
# KPI 5
# UTV
#
def _kpi5(_cursor):
    _targets = [('Dec-18',240000),('Nov-18',360000),('Oct-18',360000),('Sep-18',360000),('Aug-18',240000),('Jul-18',120000),
                ('Jun-18',240000),('May-18',360000),('Apr-18',360000),('Mar-18',360000),('Feb-18',360000),('Jan-18',360000),
                ('Dec-17',240000),('Nov-17',360000),('Oct-17',360000),('Sep-17',360000),('Aug-17',240000),('Jul-17',120000),
                ('Jun-17',240000),('May-17',360000),('Apr-17',360000),('Mar-17',360000),('Feb-17',360000),('Jan-17',360000),
                ('Dec-16',240000),('Nov-16',360000)]
    grand_result = []
    today = datetime.datetime.today()
    current_month = datetime.datetime(today.year, today.month, 1)
    for i in range(0, 12):
        end = current_month + relativedelta(years=0, months=-i+1)
        start = current_month + relativedelta(years=0, months=-i)
        str = "select [Fakt100].[dbo].fakthstk.kställe as avdelning,sum(beloppexkl) as belopp from [Fakt100].[dbo].fakthstk left join [Fakt100].[dbo].fakthsth on [Fakt100].[dbo].fakthsth.FakturaNr = [Fakt100].[dbo].fakthstk.FakturaNr and [Fakt100].[dbo].fakthsth.ordernr=[Fakt100].[dbo].fakthstk.ordernr where [Fakt100].[dbo].fakthsth.datum>='{}' and datum<'{}' and [Fakt100].[dbo].fakthstk.kställe = '4' GROUP BY [Fakt100].[dbo].fakthstk.kställe".format(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))

        _cursor.execute(str)
        dbl = _cursor.fetchall()
        result = [start.strftime("%b-%y")]
        total = 0
        for target in _targets:
            if result[0] == target[0]:
                target_value = target[1]
                for row in dbl:
                    total = total + row[1]
                result.append("%d, %d" % (total, target_value))
                logger.warning(result)
                grand_result.append(result)

    # Write the result to the kpi1.js file
    f = open(_path+"kpi5.js", "w")

    part1 = """
    function drawKpi5Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Utveckling', 'Budget'],"""

    part2 = ""
    j = 0
    for e in reversed(grand_result):
        part2 = part2 + "["
        i = 0;
        for v in e:
            if i == 0:
                part2 = part2 + "'%s'" % v
            else:
                part2 = part2 + v
            if i < len(e) - 1:
                part2 = part2 + ", "
            i = i + 1
        part2 = part2 + "]"
        if j < len(grand_result) - 1:
            part2 = part2 + ','
        j = j + 1
    part2 = part2 + "]"
    part3 = """
    	);

        var options = {
        title : 'Resultat / Utveckling / Månad',
        titleTextStyle: { color: 'grey' },
	    width:620,
	    height:500,
        vAxis: {
            title: 'Månad',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        hAxis: {
            title: 'Sek',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        legend: {
            textStyle: { color: 'grey' }
        },    
        backgroundColor: 'black',
        seriesType: 'bars',
        series: {1: {type: 'line'}}
        };

        var chart = new google.visualization.ComboChart(document.getElementById('kpi5_chart_div'));
        chart.draw(data, options);
    }
    """
    print(part1, file=f)
    print(part2, file=f)
    print(part3, file=f)
    f.close()

#
# KPI 6
# Squid
#
def _kpi6(_cursor):
    _targets = [('Dec-18',30000),('Nov-18',40000),('Oct-18',40000),('Sep-18',40000),('Aug-18',30000),('Jul-18',10000),
                ('Jun-18',30000),('May-18',40000),('Apr-18',40000),('Mar-18',40000),('Feb-18',40000),('Jan-18',40000),
                ('Dec-17',30000),('Nov-17',40000),('Oct-17',40000),('Sep-17',40000),('Aug-17',30000),('Jul-17',10000),
                ('Jun-17',30000),('May-17',40000),('Apr-17',40000),('Mar-17',40000),('Feb-17',40000),('Jan-17',40000),
                ('Dec-16',30000),('Nov-16',40000)]
    grand_result = []
    today = datetime.datetime.today()
    current_month = datetime.datetime(today.year, today.month, 1)
    for i in range(0, 12):
        end = current_month + relativedelta(years=0, months=-i+1)
        start = current_month + relativedelta(years=0, months=-i)
        str = "select [Fakt100].[dbo].fakthstk.kställe as avdelning,sum(beloppexkl) as belopp from [Fakt100].[dbo].fakthstk left join [Fakt100].[dbo].fakthsth on [Fakt100].[dbo].fakthsth.FakturaNr = [Fakt100].[dbo].fakthstk.FakturaNr and [Fakt100].[dbo].fakthsth.ordernr=[Fakt100].[dbo].fakthstk.ordernr where [Fakt100].[dbo].fakthsth.datum>='{}' and datum<'{}' and [Fakt100].[dbo].fakthstk.kställe = '9' GROUP BY [Fakt100].[dbo].fakthstk.kställe".format(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))

        _cursor.execute(str)
        dbl = _cursor.fetchall()
        result = [start.strftime("%b-%y")]
        total = 0
        for target in _targets:
            if result[0] == target[0]:
                target_value = target[1]
                for row in dbl:
                    total = total + row[1]
                result.append("%d, %d" % (total, target_value))
                logger.warning(result)
                grand_result.append(result)

    # Write the result to the kpi1.js file
    f = open(_path+"kpi6.js", "w")

    part1 = """
    function drawKpi6Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Squid', 'Budget'],"""

    part2 = ""
    j = 0
    for e in reversed(grand_result):
        part2 = part2 + "["
        i = 0;
        for v in e:
            if i == 0:
                part2 = part2 + "'%s'" % v
            else:
                part2 = part2 + v
            if i < len(e) - 1:
                part2 = part2 + ", "
            i = i + 1
        part2 = part2 + "]"
        if j < len(grand_result) - 1:
            part2 = part2 + ','
        j = j + 1
    part2 = part2 + "]"
    part3 = """
    	);

        var options = {
        title : 'Resultat / Squid/ Månad',
        titleTextStyle: { color: 'grey' },
	    width:620,
	    height:500,
        vAxis: {
            title: 'Månad',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        hAxis: {
            title: 'Sek',
            textStyle:{color: 'grey'},
            titleTextStyle:{color: 'grey'}
        },
        legend: {
            textStyle: { color: 'grey' }
        },    
        backgroundColor: 'black',
        seriesType: 'bars',
        series: {1: {type: 'line'}}
        };

        var chart = new google.visualization.ComboChart(document.getElementById('kpi6_chart_div'));
        chart.draw(data, options);
    }
    """
    print(part1, file=f)
    print(part2, file=f)
    print(part3, file=f)
    f.close()

# Main fdtvm01
sqlserver = pypyodbc.connect("Driver={SQL Server};Server=192.168.1.204;Database=Fakt100;uid=FDT;pwd=kab10mvTDF") # TODO - usr,pwd from config file
cursor = sqlserver.cursor()

_thekpi(cursor)
_kpi1(cursor)
_kpi2(cursor)
_kpi3(cursor)
_kpi4(cursor)
_kpi5(cursor)
_kpi6(cursor)
