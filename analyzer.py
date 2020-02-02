import datetime
import os
import openpyxl
import matplotlib.pyplot as pl
import numpy
from scipy.interpolate import splint


DAILY_REPORT_IN_EXCEL_PATH = './huanggang/'


def extract_data_from_official_daily_report_in_excel():
	now = datetime.datetime.now()
	day = now - datetime.timedelta(days=30)
	case_list = list()
	first_day = None

	while day <= now:
		try:
			str_first_day = datetime.datetime.strftime(day, '%Y%m%d')
			report_file = DAILY_REPORT_IN_EXCEL_PATH + str_first_day + '.xlsx'
			if os.path.exists(report_file):
				if not first_day:
					first_day = day
				work_book = openpyxl.load_workbook(report_file)
				sheet_obj = work_book.active
				case = dict()
				case[str_first_day] = {
					'newly_added': {
						'confirmed': sheet_obj.cell(row=13, column=2).value,
						'cured': sheet_obj.cell(row=13, column=3).value,
						'dead': sheet_obj.cell(row=13, column=4).value,
					}
				}
				case_list.append(case)

			day += datetime.timedelta(days=1)
		except Exception as e:
			print(e)
	return first_day, case_list
	print(first_day, case_list)


def mock_early_case(end_date, str_start_date='20200119'):
	dates = list()
	cases = list()
	start_date = datetime.datetime.strptime(str_start_date, '%Y%m%d')
	n = (end_date - start_date).days
	for i in range(n):
		dates.append(datetime.datetime.strftime(start_date + datetime.timedelta(days=i), '%Y%m%d'))
		if i == n - 1:
			cases.append(64) # 122 - 58, data from official report of Jan 25
		else:
			cases.append(0)
	return dates, cases


def draw_daily_case_figure(date_list, case_number_list, title='黄冈全市疫情新增趋势图'):
	pl.rcParams['font.sans-serif'] = ['STHeiti']
	pl.rcParams['font.serif'] = ['STHeiti']
	pl.rcParams["figure.figsize"] = (8, 4)

	# pl.plot(date_list, case_number_list, 'r', markevery=100)
	pl.plot(date_list, case_number_list, 'r', markevery=100)
	pl.grid(color='grey', axis='y')
	# pl.scatter(date_list, case_number_list)
	pl.title(title)
	# pl.show()
	t = datetime.datetime.strftime(datetime.datetime.now(), '%Y%m%d-%H%M%S')
	pl.savefig(F'reports/{title}-{t}.png')
	pl.close()


def draw_fff_daily_case_figure(date_list, case_number_list, title='黄冈全市疫情新增趋势图'):
	int_date_list = [int(i) for i in date_list]
	t = numpy.array(int_date_list)
	power = numpy.array(case_number_list)
	xnew = numpy.linspace(t.min(), t.max(), 100)
	power_smooth = spline(t, power, xnew)

	pl.rcParams['font.sans-serif'] = ['STHeiti']
	pl.rcParams['font.serif'] = ['STHeiti']

	pl.plot(xnew, power_smooth, 'r', markevery=100)
	pl.grid(color='grey', axis='y')
	pl.scatter(date_list, case_number_list)
	pl.title(title)
	pl.show()
	pl.savefig('huanggang_daily_added_confirmed_case.png')


def write_data_to_excel(date_list, case_number_list, excel_name='data.xlsx'):
	wb = openpyxl.Workbook()
	ws = wb.active
	ws.cell(row=1, column=1, value='日期')
	ws.cell(row=2, column=1, value='人数')
	for col in range(1, len(date_list) + 1):
		ws.cell(row=1, column=col + 1, value=date_list[col-1])
		ws.cell(row=2, column=col + 1, value=case_number_list[col - 1])
	wb.save(excel_name)


def sum_daily_added_cases(newly_added_case_number_list):
	accumulated_cases = list()

	accumulated_cases.append(newly_added_case_number_list[0])
	i = 1
	while i < len(newly_added_case_number_list):
		accumulated_case = accumulated_cases[i-1] + newly_added_case_number_list[i]
		accumulated_cases.append(accumulated_case)
		i += 1
	return accumulated_cases


if __name__ == '__main__':
	first_day, case_list = extract_data_from_official_daily_report_in_excel()

	date_list, newly_added_case_number_list = mock_early_case(first_day)
	for c in case_list:
		for i in c:
			date_list.append(i)
			newly_added_case_number_list.append(c.get(i).get('newly_added').get('confirmed'))

	accumulated_case_list = sum_daily_added_cases(newly_added_case_number_list)

	simplified_date_list = list(map(lambda d: d.split('2020')[1], date_list))

	write_data_to_excel(simplified_date_list, newly_added_case_number_list)
	draw_daily_case_figure(simplified_date_list, newly_added_case_number_list, '黄冈全市疫情新增趋势图')
	draw_daily_case_figure(simplified_date_list, accumulated_case_list, '黄冈全市疫情确诊累计趋势图')

	# draw_daily_case_figure(simplified_date_list, newly_added_case_number_list)


