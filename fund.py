import requests
from bs4 import BeautifulSoup
import re
import numpy as np
import matplotlib
import xlsxwriter

# 处理乱码
matplotlib.rcParams['font.sans-serif'] = ['SimHei']
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['axes.unicode_minus'] = False


# 基金字典   这里是我们的基金代码已经名称
def my_data():
    datas = {
        '001511': '兴全新视野定开混合（001511）',
        '160119': '南方中证500ETF联接A（160119）',
        # '460300': '华泰柏瑞沪深300ETF联接A（460300）',
        # '110011': '易方达中小盘混合（110011）',
        # '161005': '富国天惠成长混合AB(LOF)（161005）',
        # '005827': '易方达蓝筹精选混合（005827）',
        # # '010849': '易方达竞争优势企业混合C(010849)',
        # '163406': '兴全合润混合(LOF)(163406)',
        # '166002': '中欧新蓝筹混合A(166002)',
        # '003095': '中欧医疗健康混合A(003095)',
        # '001714': '工银文体产业股票A(001714)',
        # '163402': '兴全趋势投资混合(LOF)(163402)',
     }
    return datas


# 数据保存到execl
def save_excel(file_dir, data_list, sheet_name):
    workbook = xlsxwriter.Workbook(file_dir)

    for i in range(len(sheet_name)):
        worksheet = workbook.add_worksheet(sheet_name[i])
        bold = workbook.add_format({'bold': 1})
        headings = ['净值日期', '单位净值', '累计净值', '日增长率', '申购状态', '赎回状态', '分红送配']
        worksheet.write_row('A1', headings, bold)
        for h in range(len(data_list[i])):
            worksheet.write_row('A' + str(h + 2), data_list[i][h])

        chart_col = workbook.add_chart({'type': 'line'})

        # 配置第一个系列数据
        chart_col.add_series({
            # 如果新建sheet时设置了sheet名，这里就要设置成相应的值，如果没有的话，那就是Sheet1
            'name': '={}!$B$1'.format(sheet_name[i]),
            'categories': '={}!$A$2:$A${}'.format(sheet_name[i], len(data_list[i])),
            'values': '={}!$B$2:$B${}'.format(sheet_name[i], len(data_list[i])),
            'line': {'color': 'red'},
        })
        # 配置第二个系列数据
        chart_col.add_series({
            'name': '={}!$C$1'.format(sheet_name[i]),
            'categories': '={}!$A$2:$A${}'.format(sheet_name[i], len(data_list[i]) + 1),  # 展示名字
            'values': '={}!$C$2:$C${}'.format(sheet_name[i], len(data_list[i]) + 1),  # 展示数据
            'line': {'color': 'yellow'},
        })
        # 设置图表的title 和 x，y轴信息
        chart_col.set_title({'name': '基金走势'})
        chart_col.set_x_axis({'name': '时间'})
        chart_col.set_y_axis({'name': '价值'})
        # 设置图表的风格
        chart_col.set_style(1)
        # 把图表插入到worksheet并设置偏移
        worksheet.insert_chart('A2', chart_col, {'x_offset': 500, 'y_offset': 1})

    workbook.close()


# 数据保存到execl，集合在一个sheet
def save_excel_coll(file_dir, data_list, sheet_name):
    workbook = xlsxwriter.Workbook(file_dir)
    worksheet = workbook.add_worksheet("fund")
    bold = workbook.add_format({'bold': 1})
    # headings = ['净值日期', '单位净值', '累计净值', '日增长率', '申购状态', '赎回状态', '分红送配']
    worksheet.write_row('A1', headings, bold)
    worksheet.write(0, 0, '净值日期')
    # 写第一列（日期）
    for h in range(len(data_list[0])):
        worksheet.write(h + 1, 0, data_list[0][h][0])

    for i in range(len(sheet_name)):
        # worksheet = workbook.add_worksheet(sheet_name[i])
        # bold = workbook.add_format({'bold': 1})
        # headings = ['净值日期', '单位净值', '累计净值', '日增长率', '申购状态', '赎回状态', '分红送配']
        # worksheet.write_row('A1', headings, bold)

        for h in range(len(data_list[i])):
            # worksheet.write_row('A' + str(h + 2), data_list[i][h])
            # print("h=%d, i=%d, value=%s" % (h, i, data_list[i][h]))
            # 净值
            worksheet.write(h + 1, i+1, data_list[i][h][1])
            # print(data_list[i][h])

    workbook.close()

# 页面解析
def get_html(code, start_date, end_date, page=1, per=20):
    url = 'http://fund.eastmoney.com/f10/F10DataApi.aspx?type=lsjz&code={0}&page={1}&sdate={2}&edate={3}&per={4}'.format(
        code, page, start_date, end_date, per)
    rsp = requests.get(url)
    html = rsp.text
    return html


# 获取数据
def get_fund(code, start_date, end_date, page=1, per=20):
    # 获取html
    html = get_html(code, start_date, end_date, page, per)
    soup = BeautifulSoup(html, 'html.parser')
    # 获取总页数
    pattern = re.compile('pages:(.*),')
    result = re.search(pattern, html).group(1)
    total_page = int(result)
    # 获取表头信息
    heads = []
    for head in soup.findAll("th"):
        heads.append(head.contents[0])

    # 数据存取列表
    records = []
    # 获取每一页的数据
    current_page = 1
    while current_page <= total_page:
        html = get_html(code, start_date, end_date, current_page, per)
        soup = BeautifulSoup(html, 'html.parser')
        # 获取数据
        for row in soup.findAll("tbody")[0].findAll("tr"):
            row_records = []
            for record in row.findAll('td'):
                val = record.contents
                # 处理空值
                if val == []:
                    row_records.append(np.nan)
                else:
                    row_records.append(val[0])
            # 记录数据
            records.append(row_records)
        # 下一页
        current_page = current_page + 1

    return records


# 次程序
def collect(start_time, end_time):
    datas = []
    names = []
    data = my_data()
    for i in data:
        fund_code = i
        name = data[i]
        names.append(name)
        print(fund_code, name)
        fund_df = get_fund(fund_code, start_date=start_time, end_date=end_time)
        # 20210418 调整顺序；采集的是按日期倒叙，变为按日期正序
        fund_df.sort(reverse=False)

        for i in fund_df:  # 为了方便制图，写入到execl表内时，数字转换成浮点型
            i[1] = float(i[1])
            i[2] = float(i[2])
            i[6] = str(i[6])  # 存入到execl表前，要转换成str
        datas.append(fund_df)
    return datas, names


if __name__ == '__main__':
    datas, names = collect('2020-10-21', '2021-04-16')  # 这里是获取基金数据的时间
    # save_excel('20210418.xlsx', datas, names)  # 名称
    save_excel_coll('20210418_coll.xlsx', datas, names)  # 名称
    # save_excel('fund.xlsx', datas, names)  # 名称
