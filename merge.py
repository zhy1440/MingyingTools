import pandas as pd
import datetime
import logging
import numpy as np
from decimal import Decimal

NUMBER_CN = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
             "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
             "二十一", "二十二", "二十三", "二十四", "二十五", "二十六", "二十七", "二十八", "二十九", "三十",
             "三十一", "三十二", "三十三", "三十四", "三十五", "三十六", "三十七", "三十八", "三十九", "四十",
             "四十一", "四十二", "四十三", "四十四", "四十五", "四十六", "四十七", "四十八", "四十九", "五十",
             "五十一", "五十二", "五十三", "五十四", "五十五", "五十六", "五十七", "五十八", "五十九", "六十",
             "六十一", "六十二", "六十三", "六十四", "六十五", "六十六", "六十七", "六十八", "六十九", "七十"]


def config_logger():
    log_formatter = '[%(asctime)s] %(levelname)s : %(message)s'
    # create formatter
    formatter = logging.Formatter(log_formatter)
    log_level = logging.DEBUG

    # create logger
    logger = logging.getLogger('')
    logger.setLevel(logging.DEBUG)

    # create console handler and set level to debug
    ch = logging.StreamHandler()
    ch.setLevel(log_level)

    # add formatter to ch
    ch.setFormatter(formatter)

    fh = logging.FileHandler(
        filename='merge.log', encoding='utf-8')
    fh.setLevel(log_level)
    # tell the handler to use this format
    fh.setFormatter(formatter)

    # add the handler to the root logger
    logger.addHandler(ch)
    logger.addHandler(fh)
    return logger


SPLIT_MARK = '====================== {} ========================='
pd.set_option('display.float_format', lambda x: '%.2f' %
              x)  # 为了直观的显示数字，不采用科学计数法
logger = config_logger()


def decimal_from_value(value):
    if value:
        return Decimal(str(value).replace(',', ''))


def read_excel_file(file_name):

    def conver_date(date):
        return pd.to_datetime('2019'+date.zfill(4))

    columns = ['id', 'project_id', 'project_name', 'unit', 'project_amount',
               'unit_price', 'sum_price']
    df = pd.read_excel(file_name,
                       #   converters={'project_amount': decimal_from_value, 'unit_price': decimal_from_value, 'sum_price': decimal_from_value},
                       names=columns, header=1, na_values=['NA'])

    logger.debug(df)
    # quit()
    return df


def print_split(title):
    logger.info(SPLIT_MARK.format(title))


def process_toubiao(df_toubiao, df_jiesuan):
    ##################### Solution 1 #####################
    # obj = {}
    # sub_sum = []
    # sub_titles = []
    # unknow_list = []
    # data = []

    # for record in df_toubiao.to_records():
    #     index, id,  *(_) = record
    #     logger.debug(record)

    #     if(type(id) is str):
    #         sub_sum.append(record.tolist())
    #         data.append([])
    #     elif(type(id) is float):
    #         sub_titles.append(record.tolist())
    #         data.append([])
    #     else:
    #         result = df_jiesuan[(df_jiesuan.id == id)].to_records()
    #         if len(result) > 0:
    #             record_toubiao = record.tolist()
    #             record_jiesuan = result[0].tolist()
    #             # logger.debug(record_toubiao)
    #             # logger.debug(record_jiesuan[5:])
    #             record_sum = record_toubiao+record_jiesuan[5:]
    #             data.append(record_sum)
    #         else:
    #             unknow_list.append(record)
    ##################### Solution 2 #####################

    sub_sum_toubiao = []
    sub_titles_toubiao = []
    unknow_list_toubiao = []
    sub_sum_jiesuan = []
    sub_titles_jiesuan = []
    unknow_list_jiesuan = []
    unmatched_list = []
    result_data = []

    for record in df_toubiao.to_records():
        index, id,  *(_) = record
        # logger.debug(record)
        if(type(id) is str):
            sub_sum_toubiao.append(record.tolist())
            # data.append([])
        elif(type(id) is float):
            sub_titles_toubiao.append(record.tolist())
            # data.append([])

    for record in df_jiesuan.to_records():
        index, id,  *(_) = record
        logger.debug(record)
        if(type(id) is str):
            sub_sum_jiesuan.append(record.tolist())
            # data.append([])
        elif(type(id) is float):
            sub_titles_jiesuan.append(record.tolist())
            # data.append([])

    print_split('分部分项数')
    if len(sub_titles_toubiao) == len(sub_titles_jiesuan):
        logger.debug("是否一致: {}, 投标: {}, 结算: {}".format(
            len(sub_titles_toubiao) == len(sub_titles_jiesuan),
            len(sub_titles_toubiao), len(sub_titles_jiesuan)))
    else:
        logger.warning("是否一致: {}, 投标: {}, 结算: {}".format(
            len(sub_titles_toubiao) == len(sub_titles_jiesuan),
            len(sub_titles_toubiao), len(sub_titles_jiesuan)))
    logger.debug(sub_titles_toubiao)
    logger.debug(sub_titles_jiesuan)

    ##################### check each section #####################
    for index in range(len(sub_titles_toubiao)):
        toubiao_head = sub_titles_toubiao[index]
        toubiao_tail = sub_sum_toubiao[index]
        count_toubiao = toubiao_tail[0]-toubiao_head[0]

        jiesuan_head = sub_titles_jiesuan[index]
        jiesuan_tail = sub_sum_jiesuan[index]
        count_jiesuan = jiesuan_tail[0]-jiesuan_head[0]

        logger.debug(
            "第{}部分 -- 投标记录数: {}, 结算记录数: {} (Ln {} -> Ln {})".format(
                NUMBER_CN[index], count_toubiao, count_jiesuan, toubiao_head[0], toubiao_tail[0]))

        df_toubiao_slice = df_toubiao.loc[toubiao_head[0]+1:toubiao_tail[0]-1]
        df_jiesuan_slice = df_jiesuan.loc[jiesuan_head[0]+1:jiesuan_tail[0]-1]
        result_data.append([])  # add blank line for sub title
        for record in df_toubiao_slice.to_records():
            index, id, project_id,  *(_) = record
            result = df_jiesuan_slice[(df_jiesuan_slice.project_id ==
                                       project_id)].to_records()
            if len(result) == 1:
                record_toubiao = record.tolist()
                record_jiesuan = result[0].tolist()
                # Todo: should add validation of record is really match here, order/name/unit_price
                record_sum = record_toubiao+record_jiesuan[5:]
                result_data.append(record_sum)
            else:
                logger.debug(record_toubiao)
                logger.debug(result)
                unmatched_list.append(
                    {"toubiao": record_toubiao, "jiesuan": result})
        result_data.append([])  # add blank line for sub sum

    if len(sub_titles_toubiao) < len(sub_titles_jiesuan):
        print_split('结算增加部分({}个)'.format(
            len(sub_titles_jiesuan)-len(sub_titles_toubiao)))
        for index in range(len(sub_titles_toubiao), len(sub_titles_jiesuan)):
            logger.debug(sub_titles_jiesuan[index])
            logger.debug(sub_sum_jiesuan[index])

            jiesuan_head = sub_titles_jiesuan[index]
            jiesuan_tail = sub_sum_jiesuan[index]
            df_jiesuan_slice = df_jiesuan.loc[jiesuan_head[0] +
                                              1:jiesuan_tail[0]-1]
            logger.debug(df_jiesuan_slice)
            result_data.append([])  # add blank line for sub title
            for record in df_jiesuan_slice.to_records():
                record_jiesuan = record.tolist()
                record_sum = record_jiesuan[0:5] + \
                    (None, None, None)+record_jiesuan[5:]
                list_record_sum = list(record_sum)
                list_record_sum[1] = None
                record_sum = tuple(list_record_sum)
                logger.debug(record_sum)
                result_data.append(record_sum)
            result_data.append([])  # add blank line for sub sum

    # end_line_toubiao = sub_sum_toubiao[-1][0]
    # if end_line_toubiao < end_line_jiesuan:
    #     for line in range(end_line_toubiao, end_line_jiesuan):
    #         logger.debug(sub_titles_jiesuan)
    print_split('结果')
    logger.debug(result_data)
    if len(unmatched_list) > 0:
        print_split('无法匹配：')
        logger.debug(unmatched_list)

    logger.debug(len(sub_sum_toubiao))
    logger.debug(sub_sum_toubiao)
    logger.debug(len(sub_sum_jiesuan))
    logger.debug(sub_sum_jiesuan)
    logger.debug(len(unknow_list_toubiao))
    if len(unknow_list_toubiao) > 0:
        logger.error(unknow_list_toubiao)
    # quit()

##################### Write Excel #####################

    # Create a Pandas dataframe from some data.
    # df = pd.DataFrame({'Data': [10, 20, 30, 20, 15, 30, 45]})
    columns = ['index', 'id', 'project_id', 'project_name', 'unit', 'project_amount',
               'unit_price', 'sum_price', 'j_project_amount',
               'j_unit_price', 'j_sum_price']
    df = pd.DataFrame(result_data, columns=columns)
    df = df[columns[1:]]
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('result2.xlsx', engine='xlsxwriter')

    df.to_excel(writer, sheet_name='Sheet1',
                startrow=4, header=False, index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    workbook.set_size(1024, 800)
    worksheet = writer.sheets['Sheet1']

    style_validate = workbook.add_format({
        'bold': True,
        'font_size': 10,
        'color': 'red',
        'align': 'center',
        'valign': 'vcenter'
    })

    def gen_Title():
        print_split('Set Title')
        style_title = workbook.add_format({
            'bold': True,
            'font_size': 22,
            'font_name': "黑体",
            'align': 'center',
            'valign': 'vcenter'
        })

        worksheet.merge_range("A1:N1", 'Title', style_title)
        worksheet.set_row(0, 45)

    def set_column_style():
        print_split('Set Column Style')
        style_left = workbook.add_format({
            'font_size': 9,
            'font_name': 'Times New Roman',
            'align':'left'
        })
        style_center = workbook.add_format({
            'font_size': 9,
            'font_name': 'Times New Roman',
            'align':'center'
        })
        style_accounting = workbook.add_format({
            'font_size': 9,
            'font_name': 'Times New Roman',
            # 'fg_color': '#D7E4BC',
            # 'border': 1,
            'num_format': '#,##0.00'
        })
        column_config = [(7.67,  style_center),
                         (15,  style_left),
                         (40,  style_left),
                         (5.83,  style_center),
                         (12, style_accounting),
                         (15.50, style_accounting),
                         (12, style_accounting),
                         (12, style_accounting),
                         (15.50, style_accounting),
                         (12, style_accounting),
                         (12, style_accounting),
                         (15.5, style_accounting),
                         (12, style_left)]
        col = 0
        for config in column_config:
            worksheet.set_column(col, col, config[0], config[1])
            col += 1

    def gen_table_header():
        print_split('Set Table Header')
        style_header = workbook.add_format({
            'border': 1,
            'bold': True,
            # 'align': 'vCenter',
            'font_size': 10,
            'text_wrap': True,
            'fg_color': '#DAEEF3',
            'align': 'center',
            'valign': 'vcenter'
        })
        style_project_name = workbook.add_format({
            'font_size': 10,
            'align': 'left'
        })
        style_unit = workbook.add_format({
            'font_size': 9,
            'align': 'right'
        })

        worksheet.merge_range(
            "A2:K2", '工程名称：联合厂房3-总装车间、总装准备车间建筑工程', style_project_name)
        worksheet.merge_range("L2:N2", '单位：人民币元', style_unit)
        worksheet.merge_range("A3:A4", '序号', style_header)
        worksheet.merge_range("B3:B4", '项目编码', style_header)
        worksheet.merge_range("C3:C4", '项目名称', style_header)
        worksheet.merge_range("D3:D4", '计量单位', style_header)
        worksheet.merge_range("E3:G3", '合同价', style_header)
        worksheet.merge_range("H3:J3", '送审价', style_header)
        worksheet.merge_range("K3:M3", '审核价', style_header)
        worksheet.merge_range("N3:N4", '备注', style_header)
        col_name = ['工程量', '综合单价', '合 价', '工程量',
                    '综合单价', '合 价', '工程量', '综合单价', '合 价']
        for index in range(len(col_name)):
            worksheet.write(3, 4+index, col_name[index], style_header)

    def gen_sub_title():
        print_split('Set Sub Title')

        total_sum_toubiao = []
        total_sum_jiesuan = []
        style_subtitle = workbook.add_format({
            # 'border': 1,
            'bold': True,
            'font_size': 9,
            'text_wrap': True,
            'fg_color': '#FDE9D9',
            'align': 'center',
            'valign': 'vcenter'
        })
        style_subtitle_accounting = style_subtitle
        style_subtitle_accounting.set_num_format('#,##0.00')
        style_subtitle_accounting.set_align('right')

        for index in range(len(sub_titles_toubiao)):
            # if index < (len(sub_titles)-1):
            row = 4+sub_titles_toubiao[index][0]
            end_row = 4+sub_sum_toubiao[index][0]
            for col in range(14):
                worksheet.write(row, col, None, style_subtitle)
            worksheet.write(row, 0, NUMBER_CN[index], style_subtitle)
            worksheet.write(
                row, 2, sub_titles_toubiao[index][3], style_subtitle)
            worksheet.write(row, 6, '=SUM(G{}:G{})'.format(
                row+2, end_row), style_subtitle_accounting)
            total_sum_toubiao.append('G{}:G{}'.format(row+2, end_row))
            worksheet.write(row, 9, '=SUM(J{}:J{})'.format(
                row+2, end_row), style_subtitle_accounting)
            total_sum_jiesuan.append('J{}:J{}'.format(row+2, end_row))

        # 结算增加部分
        for index in range(len(sub_titles_toubiao), len(sub_titles_jiesuan)):
            row = 4+sub_titles_jiesuan[index][0]
            end_row = 4+sub_sum_jiesuan[index][0]
            for col in range(14):
                worksheet.write(row, col, None, style_subtitle)
            worksheet.write(row, 0, NUMBER_CN[index], style_subtitle)
            worksheet.write(
                row, 2, sub_titles_jiesuan[index][3] and "清单漏项", style_subtitle)  # Todo: 如果有名称 用名称
            worksheet.write(row, 6, '=SUM(G{}:G{})'.format(
                row+2, end_row), style_subtitle_accounting)
            total_sum_toubiao.append('G{}:G{}'.format(row+2, end_row))
            worksheet.write(row, 9, '=SUM(J{}:J{})'.format(
                row+2, end_row), style_subtitle_accounting)
            total_sum_jiesuan.append('J{}:J{}'.format(row+2, end_row))

        return total_sum_toubiao, total_sum_jiesuan

    def gen_sub_sum():
        print_split('Set Sub Sum')
        for index in range(len(sub_sum_toubiao)-1):
            row = 4+sub_sum_toubiao[index][0]
            worksheet.merge_range("A{}:C{}".format(
                row+1, row+1), None)
            worksheet.write(row, 6, sub_sum_toubiao[index][-1], style_validate)
            worksheet.write(row, 9, sub_sum_jiesuan[index][-1], style_validate)

        for index in range(len(sub_titles_toubiao), len(sub_titles_jiesuan)):
            row = 4+sub_sum_jiesuan[index][0]
            worksheet.merge_range("A{}:C{}".format(
                row+1, row+1), None)
            worksheet.write(row, 9, sub_sum_jiesuan[index][-1], style_validate)

    def set_footer(str_total_sum_toubiao, str_total_sum_jiesuan):
        print_split('Set Footer')
        style_sum = workbook.add_format({
            'bold': False,
            'font_size': 10,
            'fg_color': '#DAEEF3',
            'align': 'center',
            'valign': 'vcenter'
        })
        row = sub_sum_jiesuan[-1][0]+4
        for col in range(14):
            worksheet.write(row, col, None, style_sum)

        worksheet.set_row(row, 30)
        worksheet.merge_range("A{}:C{}".format(
            row+1, row+1), sub_sum_jiesuan[-1][1], style_sum)
        # Todo: need change to fomula
        style_sum.set_bold(True)
        worksheet.write(row, 6, '=SUM({})'.format(
            str_total_sum_toubiao), style_sum)
        # worksheet.write(row, 6, sub_sum_toubiao[-1][-1], style_sum)
        worksheet.write(row+1, 6, sub_sum_toubiao[-1][-1], style_validate)
        # Todo: need change to fomula
        worksheet.write(row, 9, '=SUM({})'.format(
            str_total_sum_toubiao), style_sum)
        # worksheet.write(row, 9, sub_sum_jiesuan[-1][-1], style_sum)
        worksheet.write(row+1, 9, sub_sum_jiesuan[-1][-1], style_validate)

    def add_cell_border():
        style_boder = workbook.add_format({'border': 1})
        for row in range(4, sub_sum_toubiao[-1][0]+4):
            for col in range(14):
                worksheet.write(row, col, None, style_boder)

    def set_table_boder(begin_row_num, end_row_num):
        style_boder = workbook.add_format({'border': 1})
        style_strong_boder_t = workbook.add_format({'top': 2})
        style_strong_boder_r = workbook.add_format({'right': 2})
        style_strong_boder_b = workbook.add_format({'bottom': 2})
        style_strong_boder_l = workbook.add_format({'left': 2})

        # worksheet.conditional_format('A{}:N{}'.format(begin_row_num, end_row_num), {
        #                              'type': 'no_errors', 'format': style_boder})
        # Top
        logger.debug('Top - A{}:N{}'.format(begin_row_num, begin_row_num))
        worksheet.conditional_format('A{}:N{}'.format(begin_row_num, begin_row_num), {
                                     'type': 'no_errors', 'format': style_strong_boder_t})
        # Right
        logger.debug('Right - N{}:N{}'.format(begin_row_num, end_row_num))
        worksheet.conditional_format('N{}:N{}'.format(begin_row_num, end_row_num), {
                                     'type': 'no_errors', 'format': style_strong_boder_r})
        # Bottom
        logger.debug('Bottom - A{}:N{}'.format(end_row_num, end_row_num))
        worksheet.conditional_format('A{}:N{}'.format(end_row_num, end_row_num), {
                                     'type': 'no_errors', 'format': style_strong_boder_b})
        # Left
        logger.debug('Left - A{}:A{}'.format(begin_row_num, begin_row_num))
        worksheet.conditional_format('A{}:A{}'.format(begin_row_num, begin_row_num), {
                                     'type': 'no_errors', 'format': style_strong_boder_l})

    begin_row = 3
    end_row_num = sub_sum_jiesuan[-1][0]+5
    gen_Title()
    
    for i in range(1, end_row_num-1):
        worksheet.set_row(i, 19.5)
    # add_cell_border()
    set_column_style()
    gen_table_header()
    total_sum_toubiao, total_sum_jiesuan = gen_sub_title()
    gen_sub_sum()
    set_footer(','.join(total_sum_toubiao), ','.join(total_sum_jiesuan))

    set_table_boder(begin_row, end_row_num)
    # Write the column headers with the defined format.
    # for col_num, value in enumerate(df.columns.values):
    #     worksheet.write(0, col_num + 1, value, header_format)

    # logger.debug(index)
    # Convert the dataframe to an XlsxWriter Excel object.
    # df1 = pd.DataFrame(columns=[11, 12, 13, 14])
    # df1.to_excel(writer, sheet_name='Sheet1', startrow=1)

    # Close the Pandas Excel writer and output the Excel file.

    # for i in range(1, sub_sum[-1][0]+4):
    #     worksheet.set_row(i, 19.5, style_vcenter)

    writer.save()


def main():

    file_name_toubiao = 'toubiao.xlsx'
    file_name_jiesuan = "jiesuan.xlsx"
    print_split('started')
    print_split('投标')
    df_toubiao = read_excel_file(file_name_toubiao)
    print_split('结算')
    df_jiesuan = read_excel_file(file_name_jiesuan)

    print_split('处理数据')
    process_toubiao(df_toubiao, df_jiesuan)


if(__name__ == '__main__'):
    main()
