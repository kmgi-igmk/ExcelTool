class ExpenseSummary:
    def __init__(self, empid_in_file:int, emp_info:tuple, expense:int, target_month:int, filename:str, expenses:list=None, err_msg:str=None):
        self.empid_in_file = empid_in_file
        self.emp_info = emp_info
        self.expense = expense
        self.month = target_month
        self.filename = filename
        self.err_msg = err_msg

        self.expenses12 = self.__create_expenses12() if expenses is None or len(expenses) == 0 else expenses

    def get_empid_in_file(self):
        return self.empid_in_file
    
    def get_emp_info(self):
        return self.emp_info

    def get_expense_summary(self):
        return self.emp_info + tuple(self.expenses12)

    def get_filename(self):
        return self.filename

    def get_expense_in_file(self):
        return self.expense
    
    def get_err_msg(self):
        return self.err_msg
    
    def get_expenses(self):
        return self.expenses12
    
    def get_report(self):
        status = 'o' if self.err_msg is None else 'x'
        err_info = f", err_info=[{self.err_msg}]" if status == 'x' else ''
        return f"{status}, filename=[{self.filename}], id_in_file=[{self.empid_in_file}], expense=[{self.expense}]{err_info}"

    def __create_expenses12(self):
        return [self.expense if x == self.month else 0 for x in range(1,13)]
