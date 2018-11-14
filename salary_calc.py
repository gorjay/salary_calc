import os
import numpy as np
from pandas import Series, DataFrame
import pandas as pd
import xlwings as xw
import logging
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox


# TODO: 表格粗体


class JobSubType:
    """
    A sub job type
    Attr:
        s_id: sub type id
        s_name: sub type name
        s_price: sub type's price
    """
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
    """
    Describe a job type
    Attr:
        j_id: job id
        j_dict_sub_types: dict[s_id(int): JobSubType]
    """
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
        raise Exception("工序号‘%s’无效" % name)


class JobTypeBook:
    """
    A job type book, collect all job types
    Attr:
        b_dict_job_types: dict[j_id(int): JobType]
    """
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
        if job_id not in self.b_dict_job_types \
                or sub_id not in self.b_dict_job_types[job_id].j_dict_sub_types:
            logging.error("query_price_by_id: Invalid job id %d or sub id %d\n"
                          % (job_id, sub_id))
            raise Exception("所要求的货品ID无效（%d-%d）\n"
                            % (job_id, sub_id))

        return self._get_job_type(job_id)._get_sub_type_by_id(sub_id).s_price

    def query_price_by_name(self, job_id: int, sub_type_name: str)->float:
        return self._get_job_type(job_id)._get_sub_type_by_name(sub_type_name).s_price

    def get_dict(self)->dict:
        """
        获取一个字典，key是所有可能的sub type id，值是每个job type对应的sub type id的单价，如下图
                  sub 0 |    sub 1 |    sub 2 | ...
        job 0 | price00 | price 01 | price 02 |...
        job 1 | price10 | price 11 | price 12 |...
        job 2 | price20 | price 21 | price 22 |..

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
        return DataFrame(self.get_dict(),
                         index=[job_type.j_id for job_type in self.b_dict_job_types.values()])


class Job:
    """
    A Job which is finished by employee
    Attr:
        job_type_id: job type id
        sub_type_id: sub job type id
        finish_count: quantities of finished jobs
    """
    def __init__(self, job_type_id: int = 0, sub_type_id: int = 0, finish_count: int = 0):
        self.job_type_id = job_type_id
        self.sub_type_id = sub_type_id
        self.finish_count = finish_count


class Employee:
    """
    A Employee
    Attr:
        e_name: employee name
        e_id: employee id
        e_do_jobs: jobs employee had finished, list = [Job]
        e_do_jobs_dict: jobs employeee had finished, dict[j_id(int): dict[s_id(int): finish_count(int)]]
    """
    def __init__(self, name: str, eid: int, file_path: str = None):
        self.e_name = name
        self.e_id = eid
        self.e_do_jobs: list = []
        self.e_do_jobs_dict: dict = {}
        if file_path is not None:
            self.load_jobs_from_file(file_path)

    def load_jobs_from_file(self, file_path: str = None):
        data_frame = pd.read_excel(file_path, sheet_name="员工产值明细",
                                   header=1, usecols=[0, 1, 2],
                                   dtype={"款号": np.int, "工序": np.int, "数量": np.int},
                                   comment="小计", convert_float=True, verbose=True)
        for index, row in data_frame.iterrows():
            new_job = Job(row["款号"], row["工序"], row["数量"])
            self.add_job(new_job)

    def add_job(self, job: Job):
        if job.job_type_id not in self.e_do_jobs_dict:
            self.e_do_jobs_dict[job.job_type_id] = {}

        if job.sub_type_id in self.e_do_jobs_dict[job.job_type_id]:
            logging.warning("员工工作记录有重复项%s %d %d" %
                            (self.e_name, job.job_type_id, job.sub_type_id))
            for job_loop in self.e_do_jobs:
                if job_loop.job_type_id == job.job_type_id and job_loop.sub_type_id == job.sub_type_id:
                    job_loop.finish_count += job.finish_count
            self.e_do_jobs_dict[job.job_type_id][job.sub_type_id] += job.finish_count
        else:
            self.e_do_jobs.append(job)
            self.e_do_jobs_dict[job.job_type_id][job.sub_type_id] = job.finish_count


class Company:
    """
    A company
    Attr:
        c_dict_employee: employees owned by the company, dict[name(str): Employee]
        c_job_type_book: job type book
    """
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
                _sum += self.c_job_type_book.query_price_by_id(job.job_type_id, job.sub_type_id) * job.finish_count
        return _sum

    def export_employee_salary_sheet(self, file_path: str):
        """
                    job type 0 | job type 1 | job type 2 ...
        employeeA | salary 0   | salary 1   | salary 2   ...
        employeeB | salary 0   | salary 1   | salary 2   ...
        employeeC | salary 0   | salary 1   | salary 2   ...

        :param file_path: file to output
        :return:
        """
        # 2 init data frame
        df = DataFrame()
        for employee in self.c_dict_employee.values():
            # each dict is a row in sheet
            _dict = {}
            for job_type in self.c_job_type_book.b_dict_job_types.values():
                _dict[job_type.j_id] = self.calc_employee_salary_in_job_type(employee.e_name, job_type.j_id)
            # add row to sheet
            df = df.append(DataFrame(_dict, index=[employee.e_name]))

        # 3 export to excel
        df = df.sort_index()
        print(df.to_string())
        df.to_excel(file_path, sheet_name="员工工资总表")
        return

    def export_job_type_output_sheet(self, file_path: str):
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

        # 2. save to file
        work_book = xw.Book()
        for (job_type, df) in dict_df.items():
            work_sheet = work_book.sheets.add(name=str(job_type))
            work_sheet.range("A1").value = df
            work_sheet.range("A1").value = "工序号"
            work_sheet.range("A1").api.Font.Bold = True
            # work_cells = work_sheet.cells
            for col_loop in range(len(df.columns)):
                if col_loop % 2:
                    work_sheet.range((1,col_loop+2)).value = "数量"
                else:
                    work_sheet.range((1,col_loop+2)).value = "工号"
                work_sheet.range((1,col_loop+2)).api.Font.Bold = True

        work_book.save(path=file_path)
        work_book.close()


class Application(tk.Frame):
    btn_output: tk.Button
    btn_quit: tk.Button
    btn_select_employee: tk.Button
    btn_select_price: tk.Button
    btn_show_employees: tk.Button
    btn_show_job_types: tk.Button
    label_selected_employee: tk.Label
    label_selected_jobtypes: tk.Label
    label_status: tk.Label
    listbox_err_report: tk.Listbox

    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()
        self.my_company: Company = Company()

    def create_widgets(self):
        self.btn_output              = tk.Button(self, text="生成表格", command=self.btn_cmd_output, width=30)
        self.btn_quit                = tk.Button(self, text="退出", command=root.destroy, width=10)
        self.btn_select_employee     = tk.Button(self, text="加载员工", command=self.btn_cmd_add_employee)
        self.btn_select_price        = tk.Button(self, text="加载价格", command=self.btn_cmd_select_price)
        self.btn_show_employees      = tk.Button(self, text="...")
        self.btn_show_job_types      = tk.Button(self, text="...")
        self.label_selected_employee = tk.Label(self, text="NA", width=20)
        self.label_selected_jobtypes = tk.Label(self, text="NA", width=20)
        self.label_status            = tk.Label(self, text="NA")
        self.listbox_err_report      = tk.Listbox(self, width=40)

        self.btn_select_employee.grid(row=0, column=0)
        self.label_selected_employee.grid(row=0, column=1, columnspan=2)
        self.btn_show_employees.grid(row=0, column=3)

        self.btn_select_price.grid(row=1, column=0)
        self.label_selected_jobtypes.grid(row=1, column=1, columnspan=2)
        self.btn_show_job_types.grid(row=1, column=3)

        self.btn_output.grid(row=2, column=0, columnspan=4)

        self.label_status.grid(row=3, column=0, columnspan=2)
        self.btn_quit.grid(row=3, column=2, columnspan=2)

        self.listbox_err_report.grid(row=4, column=0, columnspan=4)

        self.log_debug("初始化成功")

    def btn_cmd_select_output_dir(self):
        """
        选择输出的目录按钮，记录在self.entry_output_dir中
        :return:
        """
        dir_selected = tk.filedialog.askdirectory()
        if dir_selected is not '':
            self.entry_output_dir.delete(0, tk.END)
            self.entry_output_dir.insert(0, string=dir_selected)
            self.log_debug("成功选择输出目录")

    def btn_cmd_add_employee(self):
        """
        添加员工按钮：添加到Company类中
        :return:
        """
        list_file = tk.filedialog.askopenfilenames()
        if len(list_file) != 0:
            try:
                self.handle_add_employee_from_file_list(list_file)
                self.label_selected_employee["text"] = str(list_file)
                self.log_debug("成功加载员工信息")
            except Exception as e:
                self.log_error(repr(e))

    def btn_cmd_select_price(self):
        """
        选择单价本按钮：添加到Company中
        :return:
        """
        file_selected = tk.filedialog.askopenfilename()
        if file_selected is not '':
            try:
                self.handle_set_company_price_book(file_selected)
                self.label_selected_price["text"] = file_selected
                self.log_debug("成功加载货品价格信息")
            except Exception as e:
                self.log_error(repr(e))

    def btn_cmd_output(self):
        """
        程序执行按钮：产生输出文件
        :return:
        """
        # 1. check if company is ready
        if not self.my_company.c_dict_employee or not self.my_company.c_job_type_book:
            tk.messagebox.showerror(title="ghSalaryCalc",message="尚未添加货品单价或员工信息")
            return

        # 2. check path valid
        output_dir = os.getcwd()
        # output_dir: str = self.entry_output_dir.get()
        # if not os.path.exists(output_dir):
        #     if tk.messagebox.askyesno(title="ghSalaryCalc",message="输出目录将设置为当前目录"):
        #         output_dir = os.getcwd()
        #     else:
        #         return

        # 3. export files
        try:
            file_path = output_dir + os.sep + "每个款号总产量.xlsx"
            if not os.path.exists(file_path) or \
                    tk.messagebox.askyesno(title="ghSalaryCalc",message="是否覆盖原文件：%s" % file_path):
                self.my_company.export_job_type_output_sheet(file_path)
            self.log_debug("成功输出货品产量信息")
        except Exception as e:
            self.log_error(repr(e))
            return

        try:
            file_path = output_dir + os.sep + "员工工资总表.xlsx"
            if not os.path.exists(file_path) or \
                    tk.messagebox.askyesno(title="ghSalaryCalc",message="是否覆盖原文件：%s" % file_path):
                self.my_company.export_employee_salary_sheet(file_path)
            self.log_debug("成功输出员工工资信息")
        except Exception as e:
            self.log_error(repr(e))
            return
        tk.messagebox.showinfo(title="ghSalaryCalc", message="执行结束")

    def handle_add_employee_from_file_list(self, excel_file_list: list):
        for _file in excel_file_list:
            logging.debug("select %s \n" % _file)
            # check file extension
            _file_ext = os.path.splitext(_file)[1]
            if _file_ext != ".xls" and _file_ext != ".xlsx":
                raise Exception("不支持的文件格式: %s" % _file_ext)

            workbook = xw.Book(_file)

            # check sheet names
            sheet_exist = False
            for sht in workbook.sheets:
                if sht.name == "员工产值明细":
                    sheet_exist = True

            if not sheet_exist:
                raise Exception("员工文件%s里不包含表格‘员工产值明细’" % _file)

            # check sheet info
            first_line = workbook.sheets["员工产值明细"].range('A1:D1').value
            if first_line[0] != "员工：" or first_line[2] != "工号：":
                workbook.close()
                raise Exception("表格%s格式不正确")

            # add employee by sheet
            # TODO: maybe need many employees in one file
            self.my_company.add_employee(name=workbook.sheets["员工产值明细"].range("B1").value,
                                         eid=workbook.sheets["员工产值明细"].range("D1").value,
                                         file_path=_file)
            workbook.close()

    def handle_set_company_price_book(self, file_path: str):
        work_book = xw.Book(file_path)
        # check sheet names
        sheet_exist = False
        for sht in work_book.sheets:
            if sht.name == "单价表":
                sheet_exist = True

        if not sheet_exist:
            raise Exception("文件%s里不包含表格‘单价表’" % file_path)

        self.my_company.c_job_type_book = JobTypeBook(work_book.sheets["单价表"])
        work_book.close()

    def handle_output(self, output_dir: str):
        return

    def log_debug(self, message: str):
        self.label_status["text"] = message

    def log_error(self, message: str):
        self.label_status["text"] = message
        self.listbox_err_report.insert(0, message)
        tk.messagebox.showerror(title="ghSalaryCalc",message=message)


if __name__ == '__main__':
    # myapp = xw.App(visible=False)
    logging.debug("hello hsj\n")
    # init logger
    logging.basicConfig(filename='ghSalaryCalc.log', level=logging.DEBUG, format='%(asctime)s %(message)s')
    # init gui
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
    logging.debug("end hsj\n")
    # myapp.kill()
