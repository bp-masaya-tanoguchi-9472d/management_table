import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import calendar
from datetime import timedelta, datetime
import glob

def get_first_last_month(year,month):
    first_date = datetime(year,month,1)
    last_day = calendar.monthrange(first_date.year, first_date.month)[1]
    last_date = first_date.replace(day=last_day)
    return {'first_date':first_date,'last_date':last_date}

def daterange(_start, _end):
    for n in range((_end - _start).days):
        yield _start + timedelta(n)


if __name__ == '__main__':
	# input current month
	
	print("Please input 'year'")
	while True:
		try:
			year = int(input())
			print('good')
			break
		except ValueError:
			print('not int try again...')
	print("Please input 'month'")
	while True:
		try:
			month = int(input())
			print('good')
			break
		except ValueError:
			print('not int try again...')

	# make working day file
	dates = get_first_last_month(year,month)
	start = dates['first_date']
	end = dates['last_date'] + timedelta(days=1)

	Date = []
	Day_of_the_week = []
	for i in daterange(start, end):
	    Date.append(i)
	    Day_of_the_week.append(i.weekday())
	df_d = pd.DataFrame({'Date':Date,'Day of the week':Day_of_the_week})

	df_d = df_d[(df_d['Day of the week']<=4)]
	table = {0:'月',1:'火',2:'水',3:'木',4:'金',5:'土',6:'日'}
	df_d['Day of the week'] = df_d['Day of the week'].replace(table)
	df_d = df_d.reset_index(drop=True)

	N = df_d.shape[0]
	overtime = (40/N) + 8
	nan_list = [np.nan if i==0 else i for i in np.zeros(N)]


	#add information
	df_add = pd.DataFrame({'出社':nan_list,'退勤':nan_list,'勤務形態':nan_list,'(終了-開始)-休憩':nan_list,'規定労働時間(8hr)からの超過(hr)':nan_list,'平均残業時間込みからの超過(hr)':nan_list})
	df = pd.concat([df_d,df_add],axis=1)


	#insert excel code
	Work = []
	Over = []
	More_Over = []
	for i in range(2,N+2):
	    work = '=(D%s-C%s)*24-1'%(i,i)
	    over = '=F%s-8'%i
	    more_over = '=F%s-%s'%(i,overtime)
	    
	    Work.append(work)
	    Over.append(over)
	    More_Over.append(more_over)


	df['(終了-開始)-休憩'] = Work
	df['規定労働時間(8hr)からの超過(hr)'] = Over
	df['平均残業時間込みからの超過(hr)'] = More_Over


	df_emp = pd.DataFrame([],columns=df.columns,index=['SUM'])
	df_emp.loc['SUM','規定労働時間(8hr)からの超過(hr)'] = '=SUM(G2:G%s)'%(N+1)
	df_emp.loc['SUM','平均残業時間込みからの超過(hr)'] = '=SUM(H2:H%s)'%(N+1)

	df = pd.concat([df,df_emp])

	# file save
	filename = '勤務管理_%s年%s月.xlsx'%(year,month)
	FNs = glob.glob('./*')
	chk = '.\\'+filename in FNs
	if chk == True:
	    print('same excel exist...')
	else:
	    df.to_excel(filename,encoding='cp932',index=None)
	    print('finish')