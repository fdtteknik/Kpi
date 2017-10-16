import pypyodbc
import logging
import sys
import os

import datetime
from dateutil.relativedelta import *

# create logger
logger = logging.getLogger(os.path.basename(sys.argv[0]))
logger.setLevel(logging.DEBUG)


# KStalle mapping
# SELECT TOP (1000) * FROM [boka100].[dbo].[Kställe]
def _thekpi(_cursor):
    _targets = [('Dec-17', 40000), ('Nov-17', 40000), ('Oct-17', 40000), ('Sep-17', 40000), ('Aug-17', 40000), ('Jul-17', 40000), ('Jun-17', 40000), ('May-17', 40000), ('Apr-17', 40000), ('Mar-17', 40000), ('Feb-17', 40000), ('Jan-17', 40000), ('Dec-17', 40000), ('Nov-16', 40000)]
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
    f = open("C:\\Users\\MikaelE\\Desktop\\KPI\\thekpi.js", "w")

    part1 = """
    function drawTheKpiChart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Sälj', 'Target'],"""

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
	    width:620,
	    height:500,
        vAxis: {title: 'Manad'},
        hAxis: {title: 'Sek'},
        backgroundColor: '#E4E4E4',
        seriesType: 'bars',
        series: {1: {type: 'line'}}
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
    kpi1_target = [1000000, 1500000, 1000000, 1500000, 1000000, 1500000,
                   1000000, 1500000, 1000000, 1500000, 1000000, 1500000]
    grand_result = []
    today = datetime.datetime.today()
    current_month = datetime.datetime(today.year, today.month, 1)
    for i in range(0, 6):
        end = current_month + relativedelta(years=0, months=-i+1)
        start = current_month + relativedelta(years=0, months=-i)
        str = "select [Fakt100].[dbo].fakthstk.kställe as avdelning,sum(beloppexkl) as belopp from [Fakt100].[dbo].fakthstk left join [Fakt100].[dbo].fakthsth on [Fakt100].[dbo].fakthsth.FakturaNr = [Fakt100].[dbo].fakthstk.FakturaNr and [Fakt100].[dbo].fakthsth.ordernr=[Fakt100].[dbo].fakthstk.ordernr where [Fakt100].[dbo].fakthsth.datum>='{}' and datum<'{}' group by [Fakt100].[dbo].fakthstk.kställe".format(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))

        _cursor.execute(str)
        dbl = _cursor.fetchall()
        result = [start.strftime("%b-%y")]
        total = 0
        salj = 0
        supp = 0
        kons = 0
        tekn = 0
        utv = 0
        squ = 0
        misc = 0
        for row in dbl:
            total = total + row[1]
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

        result.append("%d, %d, %d, %d, %d, %d, %d, %d" % (misc, salj, supp, utv, tekn, kons, squ, total))
        logger.warning(result)
        grand_result.append(result)

    # Write the result to the kpi1.js file
    f = open("C:\\Users\\MikaelE\\Desktop\\KPI\\kpi1.js", "w")

    part1 = """
    // Callback that draws the pie chart for Anthony's pizza.
    function drawKpi1Chart() {
        // Create the data table for Anthony's pizza.
        // Some raw data (not necessarily accurate)
        var data = google.visualization.arrayToDataTable([
        ['Month', 'FDT', 'SAL', 'SUP', 'UTV', 'TEK', 'KON', 'SQU', 'TOT'],"""

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
        title : 'Resultat / Avdelning / Månad',
	    width:620,
	    height:500,
        vAxis: {title: 'Manad'},
        backgroundColor: '#E4E4E4',
        hAxis: {title: 'Sek'},
        };

        var chart = new google.visualization.LineChart(document.getElementById('kpi1_chart_div'));
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
    _targets = [('Dec-17', 40000), ('Nov-17', 40000), ('Oct-17', 40000), ('Sep-17', 40000), ('Aug-17', 40000), ('Jul-17', 40000), ('Jun-17', 40000), ('May-17', 40000), ('Apr-17', 40000), ('Mar-17', 40000), ('Feb-17', 40000), ('Jan-17', 40000), ('Dec-17', 40000), ('Nov-16', 40000)]
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
    f = open("C:\\Users\\MikaelE\\Desktop\\KPI\\kpi2.js", "w")

    part1 = """
    function drawKpi2Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Teknik', 'Target'],"""

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
	    width:620,
	    height:500,
        vAxis: {title: 'Manad'},
        hAxis: {title: 'Sek'},
        backgroundColor: '#E4E4E4',
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
    _targets = [('Dec-17', 40000), ('Nov-17', 40000), ('Oct-17', 40000), ('Sep-17', 40000), ('Aug-17', 40000), ('Jul-17', 40000), ('Jun-17', 40000), ('May-17', 40000), ('Apr-17', 40000), ('Mar-17', 40000), ('Feb-17', 40000), ('Jan-17', 40000), ('Dec-17', 40000), ('Nov-16', 40000)]
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
    f = open("C:\\Users\\MikaelE\\Desktop\\KPI\\kpi3.js", "w")

    part1 = """
    function drawKpi3Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Sälj', 'Target'],"""

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
	    width:620,
	    height:500,
        vAxis: {title: 'Manad'},
        hAxis: {title: 'Sek'},
        backgroundColor: '#E4E4E4',
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
    _targets = [('Dec-17', 40000), ('Nov-17', 40000), ('Oct-17', 40000), ('Sep-17', 40000), ('Aug-17', 40000), ('Jul-17', 40000), ('Jun-17', 40000), ('May-17', 40000), ('Apr-17', 40000), ('Mar-17', 40000), ('Feb-17', 40000), ('Jan-17', 40000), ('Dec-17', 40000), ('Nov-16', 40000)]
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
    f = open("C:\\Users\\MikaelE\\Desktop\\KPI\\kpi4.js", "w")

    part1 = """
    function drawKpi4Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Konsult', 'TARGET'],"""

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
	    width:620,
	    height:500,
        vAxis: {title: 'Manad'},
        hAxis: {title: 'Sek'},
        backgroundColor: '#E4E4E4',
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
    _targets = [('Dec-17', 40000), ('Nov-17', 40000), ('Oct-17', 40000), ('Sep-17', 40000), ('Aug-17', 40000), ('Jul-17', 40000), ('Jun-17', 40000), ('May-17', 40000), ('Apr-17', 40000), ('Mar-17', 40000), ('Feb-17', 40000), ('Jan-17', 40000), ('Dec-17', 40000), ('Nov-16', 40000)]
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
    f = open("C:\\Users\\MikaelE\\Desktop\\KPI\\kpi5.js", "w")

    part1 = """
    function drawKpi5Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Utveckling', 'Target'],"""

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
	    width:620,
	    height:500,
        vAxis: {title: 'Manad'},
        hAxis: {title: 'Sek'},
        backgroundColor: '#E4E4E4',
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
    _targets = [('Dec-17', 40000), ('Nov-17', 40000), ('Oct-17', 40000), ('Sep-17', 40000), ('Aug-17', 40000), ('Jul-17', 40000), ('Jun-17', 40000), ('May-17', 40000), ('Apr-17', 40000), ('Mar-17', 40000), ('Feb-17', 40000), ('Jan-17', 40000), ('Dec-17', 40000), ('Nov-16', 40000)]
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
    f = open("C:\\Users\\MikaelE\\Desktop\\KPI\\kpi6.js", "w")

    part1 = """
    function drawKpi6Chart() {
        var data = google.visualization.arrayToDataTable([
        ['Month', 'Squid', 'Target'],"""

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
	    width:620,
	    height:500,
        vAxis: {title: 'Manad'},
        hAxis: {title: 'Sek'},
        backgroundColor: '#E4E4E4',
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

# Main
sqlserver = pypyodbc.connect("Driver={SQL Server};Server=fdtvm01;Database=Fakt100;uid=FDT;pwd=kab10mvTDF") # TODO - usr,pwd from config file
cursor = sqlserver.cursor()

_thekpi(cursor)
_kpi1(cursor)
_kpi2(cursor)
_kpi3(cursor)
_kpi4(cursor)
_kpi5(cursor)
_kpi6(cursor)
