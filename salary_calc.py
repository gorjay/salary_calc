import numpy as np
from pandas import Series, DataFrame
import pandas as pd
import xlwings as xw
import logging
import tkinter as tk
import tkinter.filedialog


# type define
class JobSubType:
    s_id: int
    s_name: str
    s_price: float

    def __init__(self, sid: int, name: str, price: float):
        self.s_id = sid
        self.s_name = name
        self.s_price = price

    def print(self):
        print("[s_id %d] [s_name %s] [s_price %d]" %
              (self.s_id, self.s_name, self.s_price))


class JobType:
    def __init__(self, jid: int):
        self.j_id: int = jid
        self.j_dict_sub_types: dict[int:JobSubType] = {}

    def add_sub_type(self, sid: int, name: str, price: float):
        sub_type = JobSubType(sid, name, price)
        self.j_dict_sub_types[sid] = sub_type

    def _get_sub_type_by_id(self, sid: int)->JobSubType:
        return self.j_dict_sub_types[sid]

    def _get_sub_type_by_name(self, name: str)->JobSubType:
        for sub_type in self.j_dict_sub_types.values():
            if sub_type.s_name == name:
                return sub_type
        logging.error("Can't find sub type id %s\n", name)
        return None


class JobTypeBook:
    def __init__(self, sht: xw.main.Sheet):
        self.b_dict_job_types: dict[int:JobType] = {}
        list_region_start_cols = []
        # 1 split sheet into regions, get region start col
        continue_blank_line_cnt = 1
        ncols = 1
        while continue_blank_line_cnt < 3:
            if sht.range((1, ncols)).value is None:
                continue_blank_line_cnt += 1
            else:
                if continue_blank_line_cnt != 0:
                    list_region_start_cols.append(ncols)
                continue_blank_line_cnt = 0
            ncols += 1

        ncols -= 4

        # 2 get region nrows
        for start_col in list_region_start_cols:
            job_type = JobType(int(sht.range((1, start_col + 2)).value))
            raw_loop = 3
            while sht.range((raw_loop, start_col + 1)).value is not None:
                job_type.add_sub_type(sid=int(sht.range((raw_loop, start_col)).value),
                                      name=sht.range((raw_loop, start_col + 1)).value,
                                      price=sht.range((raw_loop, start_col + 2)).value)
                raw_loop += 1
            self.add_job_type(job_type)

    def _get_job_type(self, jid: int)->JobType:
        return self.b_dict_job_types[jid]

    def add_job_type(self, job_type: JobType):
        self.b_dict_job_types[job_type.j_id] = job_type

    def query_price_by_id(self, job_id: int, sub_id: int)->float:
        if job_id not in self.b_dict_job_types or sub_id not in self.b_dict_job_types[job_id].j_dict_sub_types:
            logging.error("query_price_by_id: Invalid job id %d or sub id %d\n" % (job_id, sub_id))
            return 0.0

        return self._get_job_type(job_id)._get_sub_type_by_id(sub_id).s_price   # TODO:add para check

    def query_price_by_name(self, job_id: int, sub_type_name: str)->float:
        return self._get_job_type(job_id)._get_sub_type_by_name(sub_type_name).s_price   # TODO:add para check

    def get_dict(self)->dict:
        """
        获取一个字典，key是所有可能的sub type id，值是每个job type对应的sub type id的单价，如下图
                  sub 0 \    sub 1 \    sub 2 \ ...
        job 0 \ price00 \ price 01 \ price 02 \...
        job 1 \ price10 \ price 11 \ price 12 \...
        job 2 \ price20 \ price 21 \ price 22 \..

        :return: 字典
        """
        dict_book = {}
        for job_type in self.b_dict_job_types.values():
            for sub_type in job_type.j_dict_sub_types.values():
                if sub_type in dict_book:
                    dict_book[sub_type.s_id].append(sub_type.s_price)
                else:
                    dict_book[sub_type.s_id] = [sub_type.s_price]
        return dict_book

    def get_data_frame(self)->DataFrame:
        return DataFrame(self.get_dict(), index=[job_type.j_id for job_type in self.b_dict_job_types.values()])


class Job:
    def __init__(self, job_type_id: int = 0, sub_type_id: int = 0, finish_count: int = 0):
        self.job_type_id = job_type_id
        self.sub_type_id = sub_type_id
        self.finish_count = finish_count


class Employee:
    def __init__(self, name: str, eid: int, file_path: str = None):
        self.e_name = name
        self.e_id = eid
        self.e_do_jobs: list = []
        if file_path is not None:
            self.load_jobs_from_file(file_path)

    def load_jobs_from_file(self, file_path: str = None):
        data_frame = pd.read_excel(file_path, sheet_name="员工产值明细", header=1, usecols=[0, 1, 2],
                                   dtype={"款号": np.int, "工序": np.int, "数量": np.int},
                                   comment="小计", convert_float=True, verbose=True)
        for index, row in data_frame.iterrows():
            new_job = Job(row["款号"], row["工序"], row["数量"])
            self.add_job(new_job)

    def add_job(self, job: Job):
        # TODO: 去掉(job id, sub type id)重复项
        self.e_do_jobs.append(job)


class OutputSheet:
    def __init__(self, job_type: int, xw_sheet: xw.Sheet):
        self.job_type: int = job_type
        self.xw_sheet: xw.Sheet = xw_sheet
        self.dict_sub_type_next_col: dict = {}

    def add_item(self, sub_type: int, employee_id: int, quantity: int)-> bool:
        if sub_type not in self.dict_sub_type_next_col:
            self.dict_sub_type_next_col[sub_type] = 2

        try:
            self.xw_sheet.range((sub_type + 2, self.dict_sub_type_next_col[sub_type])).value = employee_id
            self.xw_sheet.range((sub_type + 2, self.dict_sub_type_next_col[sub_type] + 1)).value = quantity
            self.dict_sub_type_next_col[sub_type] += 2
        except MemoryError:
            print("err")
            return False

        return False

    def last_write(self):
        return


class TotalOutputBook:
    def __init__(self):
        self.work_book = xw.Book()
        self.dict_output_sheets: dict = {}
        self.test_cnt = 1
        self.dict_data_frames: dict = {}

    def add_item(self, job_type: int, sub_type: int, employee_id: int, quantity: int):
        """
        Write a unit in workbook
        :param job_type: job type id
        :param sub_type: job sub type id
        :param employee_id: employee id
        :param quantity: item finish quantity
        :return: boolean True-success, False-fail
        """
        # if job_type not in self.dict_output_sheets:
        #     self.dict_output_sheets[job_type] = OutputSheet(job_type, self.work_book.sheets.add(str(job_type)))
        # return self.dict_output_sheets[job_type].add_item(sub_type, employee_id, quantity)
        self.work_book.sheets["Sheet1"].range((self.test_cnt, 1)).value = 1
        self.test_cnt += 1
        return True

    def export_file(self, file_path: str):
        for sheet in self.dict_output_sheets.values():
            sheet.last_write()
        self.work_book.save(file_path)
        self.work_book.close()


class Company:
    def __init__(self):
        self.c_dict_employee: dict[str:Employee] = {}
        self.c_job_type_book: JobTypeBook = None

    def add_employee(self, name: str, eid: int, file_path: str):
        e: Employee = Employee(name, eid, file_path)
        self.c_dict_employee[e.e_name] = e

    def calc_employee_salary_in_job_type(self, name: str, job_type_id: int)->int:
        _sum = 0.0
        for job in self.c_dict_employee[name].e_do_jobs:
            if job.job_type_id == job_type_id:
                _sum += self.c_job_type_book.query_price_by_id(job.job_type_id, job.sub_type_id) * job.finish_count # TODO:错误处理
        return _sum

    def export_employee_salary_sheet(self, file_path: str):
        """
                    job type 0 \ job type 1 \ job type 2 ...
        employeeA \ salary 0   \ salary 1   \ salary 2   ...
        employeeB \ salary 0   \ salary 1   \ salary 2   ...
        employeeC \ salary 0   \ salary 1   \ salary 2   ...

        :param file_path: file to output
        :return:
        """
        # 1 check file exist
        # TODO

        # 2 init data frame
        df = DataFrame()
        for employee in self.c_dict_employee.values():
            # each dict is a row in sheet
            _dict = {}
            for job_type in self.c_job_type_book.b_dict_job_types.values():
                _dict[job_type.j_id] = self.calc_employee_salary_in_job_type(employee.e_name, job_type.j_id) # TODO:错误处理
            # add row to sheet
            df = df.append(DataFrame(_dict, index=[employee.e_name]))

        # 3 export to excel TODO:列排序
        print(df.to_string())
        df.to_excel(file_path, sheet_name="员工工资总表")
        return

    def export_job_type_output_sheet(self, file_path: str):
        # TODO: check file_path
        # 1. Generate dict
        # dict_job_type_book structure ---
        # dict_job_type_book[job_type_id]
        #   = dict[sub_type_id]
        #       = list
        #           = (employee_id, job.finish_count)
        dict_job_type_book = {}
        for employee in self.c_dict_employee.values():
            for job in employee.e_do_jobs:
                if job.job_type_id not in dict_job_type_book:
                    dict_job_type_book[job.job_type_id] = {}
                if job.sub_type_id not in dict_job_type_book[job.job_type_id]:
                    dict_job_type_book[job.job_type_id][job.sub_type_id] = []
                dict_job_type_book[job.job_type_id][job.sub_type_id].append(employee.e_id)
                dict_job_type_book[job.job_type_id][job.sub_type_id].append(job.finish_count)

        # 2. create data frame
        # output data frame :
        # job_type: 180072
        #   \ e_id \ count \ e_id \ count ...
        # 0 \
        # 1 \
        dict_df = {}
        for (job_type, dict_sub_type) in dict_job_type_book.items():
            dict_df[job_type] = DataFrame()
            for (sub_type, job_list) in dict_sub_type.items():
                tmp_df = DataFrame(job_list, columns=[sub_type])
                dict_df[job_type] = dict_df[job_type].append(tmp_df.T)
            print(dict_df[job_type].to_string())

        # 2. Write to sheet
        # work_book = xw.Book(file_path)
        # for (job_type_id, dict_sub_types) in dict_job_type_book.items():
        #     work_book.sheets.add(name=str(job_type_id))
        #     for (sub_type_id, list_jobs) in dict_sub_types.items():
        #         col_next = 1
        #         for (employee_id, job_count) in list_jobs:
        #             work_book.sheets[str(job_type_id)].range((sub_type_id + 2, col_next)).value = employee_id
        #             work_book.sheets[str(job_type_id)].range((sub_type_id + 2, col_next + 1)).value = job_count
        #             col_next += 2
        # 3. save to file
        work_book = xw.Book()
        for (job_type, df) in dict_df.items():
            work_book.sheets.add(name=str(job_type)).range("A1").value = df
        work_book.save(path=file_path)
        work_book.close()


class Application(tk.Frame):
    label_selected_employee: tk.Label
    label_selected_price: tk.Label
    textbox_output_dir: tk.Text
    btn_select_employee: tk.Button
    btn_select_price: tk.Button
    btn_output: tk.Button
    quit: tk.Button

    def __init__(self, master=None):
        super().__init__(master)    #TODO:what??
        self.pack()
        self.create_widgets()
        self.my_company: Company = Company()
        self.test()

    def create_widgets(self):
        # init frame
        top_frame = tk.Frame(self)
        top_frame.pack(side=tk.TOP)
        frame = tk.Frame(self)
        frame.pack()
        bottom_frame = tk.Frame(self)
        bottom_frame.pack(side=tk.BOTTOM)

        # init widgets
        self.label_selected_employee = tk.Label(top_frame)
        self.label_selected_employee["text"] = "None employee file."
        self.label_selected_employee.pack()

        self.label_selected_price = tk.Label(top_frame)
        self.label_selected_price["text"] = "None price file."
        self.label_selected_price.pack()

        self.entry_output_dir = tk.Entry(top_frame)
        self.entry_output_dir["text"] = "None"
        self.entry_output_dir.pack()

        self.btn_select_employee = tk.Button(frame)
        self.btn_select_employee["text"] = "添加员工文件"
        self.btn_select_employee["command"] = self.btn_add_employee_file
        self.btn_select_employee.pack(side=tk.LEFT)

        self.btn_select_price = tk.Button(frame)
        self.btn_select_price["text"] = "选择单价文件"
        self.btn_select_price["command"] = self.btn_select_price_file
        self.btn_select_price.pack(side=tk.LEFT)

        self.btn_output = tk.Button(frame)
        self.btn_output["text"] = "生成输出文件"
        self.btn_output["command"] = self.btn_output
        self.btn_output.pack(side="bottom")

        self.quit = tk.Button(bottom_frame, text="退出", command=root.destroy)
        self.quit.pack(side=tk.BOTTOM)

    def handle_add_employee_from_file_list(self, excel_file_list: list):
        for _file in excel_file_list:
            logging.debug("select %s \n" % _file)
            workbook = xw.Book(_file)
            sheet_employee_job_assert(workbook.sheets["员工产值明细"])
            self.my_company.add_employee(name=workbook.sheets["员工产值明细"].range("B1").value,
                                         eid=workbook.sheets["员工产值明细"].range("D1").value,
                                         file_path=r'E:\\code\\app\\private\\前工序1807.xls')
            workbook.close()

    def handle_set_company_price_book(self, file_path: str):
        work_book = xw.Book(file_path)
        self.my_company.c_job_type_book = JobTypeBook(work_book.sheets["单价表"])
        work_book.close()

    def handle_output(self, output_dir: str):
        return

    def btn_add_employee_file(self):
        list_file = tkinter.filedialog.askopenfilenames()
        if len(list_file) == 0:
            logging.error("No file has been selected\n")
        else:
            self.label_selected_employee["text"] = "Employee file list: %s" % str(list_file)
            self.handle_add_employee_from_file_list(list_file)

    def btn_select_price_file(self):
        file_selected = tkinter.filedialog.askopenfilename()
        if file_selected is '':
            logging.error("No file has been selected\n")
        else:
            self.label_selected_price["text"] = "Price file: %s" % file_selected
            self.handle_set_company_price_book(file_selected)

    def btn_output(self):
        return

    def test(self):
        # self.handle_set_company_price_book(r'E:\\code\\app\\private\\前工序1807.xls')
        # self.handle_add_employee_from_file_list([r'E:\\code\\app\\private\\前工序1807.xls'])
        # company operation
        # self.my_company.export_employee_salary_sheet(r'E:\\code\\app\\private\\员工工资总表.xls')
        # self.my_company.export_job_type_output_sheet(r'E:\\code\\app\\private\\每个款号总产量.xlsx')
        return


# function
def sheet_employee_job_assert(sht: xw.main.Sheet):
    first_line = sht.range('A1:D1').value
    assert first_line[0] == "员工："
    assert first_line[2] == "工号："


if __name__ == '__main__':
    logging.debug("hello hsj\n")
    # init logger
    logging.basicConfig(filename='example.log', level=logging.DEBUG, format='%(asctime)s %(message)s')
    logging.debug('This message should go to the log file')
    logging.warning('And this, too')
    # init gui
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
    logging.debug("end hsj\n")
