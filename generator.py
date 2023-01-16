import pandas as pd
import numpy
import openpyxl


all_statements = []
sales_of_goods_words = ['ACH_DEBIT', 'CHECK_PAID']
bank_fee_words = ['FEE_TRANSACTION']


def format_data():
    doc = pd.read_excel('Sample Income Statements.xlsx', index_col=False)

    # print(doc)
    cols = doc.to_dict()

    details = cols['Details']
    date = cols['Posting Date']
    desc = cols['Description']
    amount = cols['Amount']
    type = cols['Type']

    for idx, detail in enumerate(details):
        obj = {}
        obj['Details'] = detail
        obj['Date'] = date[idx]
        obj['Description'] = desc[idx]
        obj['Amount'] = amount[idx]
        obj['Type'] = type[idx]

        if obj['Amount'] < 0:
            obj['Section'] = "Expense"
            if obj['Type'] in sales_of_goods_words:
                obj['Category'] = 'Sales of Goods Fees'
            elif obj['Type'] in bank_fee_words:
                obj['Category'] = 'Bank Fees'
            else:
                obj['Category'] = "Misc Fees"
        else:
            obj['Section'] = 'Income'
            obj['Category'] = 'Sales'

        all_statements.append(obj)

    print(all_statements[11])


class IncomeStatement:
    def __init__(self, all_statements) -> None:
            self.all_statements = all_statements
            self.revenue = 0
            self.total_expenses = 0
            self.net_income = 0
            self.expenses_list = []
            self.income_list = []
            self.misc_fee = 0
            self.bank_fees = 0
            self.cost_of_goods = 0
    
    def calc_expenses_revenue(self):
        for statement in self.all_statements:
            if statement['Section'] == 'Expense':
                added = int(statement['Amount'])
                self.total_expenses += added
            elif statement['Section'] == 'Income':
                added = int(statement['Amount'])
                self.revenue += added
        self.net_income = self.revenue + self.total_expenses
            
    def organize_fees(self):
        for statement in self.all_statements:
            if statement['Category'] == 'Misc Fees':
                self.misc_fee += int(statement['Amount'])
            elif statement['Category'] == 'Bank Fees':
                self.bank_fees += int(statement['Amount'])
            elif statement['Category'] == 'Sales of Goods Fees':
                self.cost_of_goods += int(statement['Amount'])

    



def main():
    company_name = input('What is the company name? ')
    format_data()
    project = IncomeStatement(all_statements)
    project.calc_expenses_revenue()
    project.organize_fees()
    # print(project.total_expenses)
    # print(project.revenue)
    # print(f'Net Income of {project.net_income}')


    data = {'Details': ['Revenue','Sales', '', 'Expenses', 'Cost of Goods Sold', 'Bank Fees', 'Misc Fees', ' ','Total Expenses',' ','Net Income: '],
    'Numbers':[' ', str(project.revenue), '', '',str(-project.cost_of_goods), str(-project.bank_fees), str(-project.misc_fee),'',str(-project.total_expenses),'', str(project.net_income)]}
    df = pd.DataFrame(data)
    print(df)
    with pd.ExcelWriter('Sample Income Statements.xlsx', mode='a', if_sheet_exists='replace') as writer:  
        df.to_excel(writer, sheet_name=company_name, index=False)



if __name__ == '__main__':
    main()
