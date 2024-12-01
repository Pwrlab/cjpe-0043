import argparse
import sys
import os
import json
from pandas import read_excel, DataFrame, ExcelWriter
import refilling


class RefillingTool:
    def __init__(self):
        self.config = self.load_config()
        # 配置参数
        self.input_file = './example.xlsx'
        self.out_file = ''
        self.methods = {"refilling": False, "exp": False, "line": False, "et": False}
        self.res_format = ['rate', 'sum', 'night_sum']
        # 创建解析器
        self.parser = argparse.ArgumentParser()
        self.args = None
        # 声明数据存储的变量
        self.f = None
        self.columns_num = 0
        self.rows_num = 0
        self.columns = []
        self.time = []
        self.solar_radiation = []
        # 存放切片的夜间数据
        self.temp_data = []
        # 结果数据
        self.res_rate = {method: {attribute: [] for attribute in self.res_format} for method in self.methods.keys()}
        self.res_verbose = {method: [] for method in self.methods.keys()}
        self.res_sheet = None

    def load_config(self):
        with open('config.json', 'r') as f:
            config = json.load(f)
        return config

    def generate_parser(self):
        self.parser.add_argument('-i', '--input', type=str, help='Input Excel file.')
        self.parser.add_argument('-o', '--output', type=str, default='.', help='Output excel file.')
        self.parser.add_argument('-a', '--all', action='store_true', default=False, help='use origin refilling method.')
        self.parser.add_argument('-r', '--refilling', action='store_true', help='use origin refilling method.')
        self.parser.add_argument('-e', '--exp', action='store_true', help='use exp refilling method.')
        self.parser.add_argument('-l', '--line', action='store_true', help='use line refilling method.')
        self.parser.add_argument('-t', '--et', action='store_true', help='use extend refilling method.')
        self.parser.add_argument('-v', '--verbose', action='store_true', default=False,
                                 help='output verbose refilling information.')

    def parse_args(self):
        # 解析参数
        self.args = self.parser.parse_args()
        if len(sys.argv) == 1:
            self.parser.print_help(sys.stderr)
            sys.exit()  # 结束程序
        if self.args.input:
            self.input_file = self.args.input
        if self.args.output:
            out_path = self.args.output
            self.out_file = out_path + '/res.xlsx'
        if self.args.all:
            self.methods = {key: True for key in self.methods}
        if self.args.refilling:
            if self.methods["refilling"]:
                pass
            else:
                self.methods["refilling"] = True
        if self.args.exp:
            if self.methods["exp"]:
                pass
            else:
                self.methods["exp"] = True
        if self.args.line:
            if self.methods["line"]:
                pass
            else:
                self.methods["line"] = True
        if self.args.et:
            if self.methods["et"]:
                pass
            else:
                self.methods["et"] = True

        if self.input_file == '':
            print("Please input the file path.")
            sys.exit()

    def read_file(self):
        # 判断文件是否存在
        if not os.path.isfile(self.input_file):
            print("The file does not exist.")
            sys.exit()

        # 读取文件
        try:
            self.f = read_excel(self.input_file, sheet_name=0, header=0, skiprows=0)
        except Exception as e:
            print(e)
            sys.exit()

        self.columns_num = self.f.shape[1]
        self.rows_num = self.f.shape[0]
        self.columns = self.f.columns
        # 按照光合强度划分每天的时间段
        self.temp_data = [[] for _ in self.columns]

    def extract_data(self):
        # 按照时间段划分数据
        night_start_index = 0
        night_end_index = 0
        try:
            while night_end_index < self.rows_num:
                # 判断夜间是否结束
                if self.f['solar_radiation'][night_end_index] <= 5:
                    if night_end_index + 1 == self.rows_num:
                        # 把夜间液流数据划分出来
                        for i, tree in enumerate(self.columns):
                            arr = self.f[tree][night_start_index:night_end_index].values
                            self.temp_data[i].append(arr)
                        break
                    if self.f['solar_radiation'][night_end_index + 1] > 5:
                        # 把夜间液流数据划分出来
                        for i, tree in enumerate(self.columns):
                            arr = self.f[tree][night_start_index:night_end_index].values
                            self.temp_data[i].append(arr)
                # 判断夜间是否开始
                if self.f['solar_radiation'][night_end_index] > 5:
                    if night_end_index + 1 == self.rows_num:
                        break
                    if self.f['solar_radiation'][night_end_index + 1] < 5:
                        night_start_index = night_end_index + 1
                night_end_index += 1
        except Exception as e:
            print(e)
            sys.exit()

    def calculate_data(self):
        # 整理数据
        self.time = self.temp_data[0]
        self.solar_radiation = self.temp_data[1]
        trees = self.temp_data[2:]
        # 初始化结果表格
        for k in self.methods.keys():
            if self.methods[k]:
                for i in self.res_format:
                    self.res_rate[k][i] = [[] for _ in range(self.columns_num)]
                if self.args.verbose:
                    self.res_verbose[k] = [[] for _ in range(self.columns_num)]

        # 将每天的数据带入模型
        for i, tree in enumerate(trees):
            for night_data in tree:
                for k in self.methods.keys():
                    if self.methods[k]:
                        rate = []
                        fit = None
                        if k == 'refilling':
                            method_param = self.config.get('refilling', {'default_point': 30, 'kneed_process': False})
                            fit, rate = refilling.refilling(night_data, default_point=method_param['default_point'],
                                                            kneed_process=method_param['kneed_process'])
                        elif k == 'exp':
                            method_param = self.config.get('exp', {'default_point': 10, 'kneed_process': False})
                            fit, rate = refilling.extended_exp_refilling(night_data,
                                                                         default_point=method_param['default_point'],
                                                                         kneed_process=method_param['kneed_process'])
                        elif k == 'line':
                            method_param = self.config.get('line', {'default_point': 10, 'kneed_process': False})
                            fit, rate = refilling.extended_line_refilling(night_data,
                                                                          default_point=method_param['default_point'],
                                                                          kneed_process=method_param['kneed_process'])
                        elif k == 'et':
                            method_param = self.config.get('et', {'default_point': 10, 'kneed_process': False,
                                                                  'min_strategy': True})
                            fit, rate = refilling.extended_transpiration(night_data,
                                                                         default_point=method_param['default_point'],
                                                                         kneed_process=method_param['kneed_process'],
                                                                         min_strategy=method_param['min_strategy'])
                        # 保存结果
                        self.res_rate[k]['rate'][i + 2].append(rate[2])
                        self.res_rate[k]['sum'][i + 2].append(rate[1])
                        self.res_rate[k]['night_sum'][i + 2].append(rate[0])
                        if self.args.verbose:
                            self.res_verbose[k][i + 2].extend(fit)
                            self.res_verbose[k][i + 2].append(None)

        for k in self.methods.keys():
            if self.methods[k]:
                for t in self.time:
                    for f in self.res_format:
                        self.res_rate[k][f][0].append(t[0])
                    if self.args.verbose:
                        self.res_verbose[k][0].extend(t)
                        self.res_verbose[k][0].append(None)

    def save_data(self):
        # 保存数据
        writer = ExcelWriter(self.out_file)
        for k in self.methods.keys():
            if self.methods[k]:
                for f in self.res_format:
                    self.res_sheet = DataFrame(self.res_rate[k][f]).transpose()
                    self.res_sheet.to_excel(writer, sheet_name=k + '_' + f, index=False, header=self.columns)
                if self.args.verbose:
                    self.res_sheet = DataFrame(self.res_verbose[k]).transpose()
                    self.res_sheet.to_excel(writer, sheet_name=k + '_fitting', index=False, header=self.columns)
        # 避免空表错误
        DataFrame().to_excel(writer, sheet_name='Empty')
        writer.close()

    def run(self):
        logo()
        # 指令系统
        self.generate_parser()
        # 解析参数
        self.parse_args()
        # 读取文件
        self.read_file()
        # 提取数据
        self.extract_data()
        # 计算数据
        self.calculate_data()
        # 保存数据
        self.save_data()


def logo():
    pwr_lab_logo = r'''   ___  __    __  __    __    _      ___
  / _ \/ / /\ \ \/__\  / /   /_\    / __\
 / /_)/\ \/  \/ / \// / /   //_\\  /__\//
/ ___/  \  /\  / _  \/ /___/  _  \/ \/  \
\/       \/  \/\/ \_/\____/\_/ \_/\_____/
                                         '''
    print(pwr_lab_logo)


if __name__ == '__main__':
    # 指令系统
    app = RefillingTool()
    app.run()
    print("over")
