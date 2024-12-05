'''
Set up a class to handle structured output form OpenAI models
'''

from pydantic import BaseModel

class FinancialStatementExtract(BaseModel):
    Income_statement_millions: str
    Revenue_Item_1: str
    Revenue_Item_2: str
    Revenue_Item_3: str
    Sales: str

    Cost_of_Sales: str
    Selling_General_and_Administrative: str
    Cost_Item_1: str
    Cost_Item_2: str
    Cost_Item_3: str
    Operating_Income: str

    Interest_Expense: str
    Other_Income_expense_net: str
    Provision_for_Income_Tax: str
    Earnings_from_Discontinued_Operations: str
    Net_Income: str

    Adjusted_Operating_Income: str
    Net_Income: str
    Interest_Expense: str
    Income_Taxes: str
    Depreciation_Amortization: str

    Adjustment_1: str
    Adjustment_2: str
    Adjustment_3: str
    Adjustment_4: str
    Adjusted_EBITDA: str
    Cash_Adjustments: str

    Cash_Flow_Statement: str
    Operating_Activities: str

    Net_Income: str
    Depreciation_and_Amortization: str
    Operating_Activity_1: str
    Operating_Activity_2: str
    Operating_Activity_3: str
    Operating_Activity_4: str
    Operating_Activity_5: str
    Operating_Activity_6: str
    Change_in_Working_Capital: str
    Cash_Flow_From_Operating_Activities: str

    Working_Capital: str

    Accounts_Receivable: str
    Inventories: str
    Prepaid_Expenses_and_Other_Current_Assets: str
    Accounts_Payable: str
    Accrued_Expenses: str
    Due_From_Due_to_Related_Party: str
    Income_Taxes_Payables: str
    Others: str
    Change_in_Working_Capital: str

    Capital_Expenditures: str
    Activity_1: str
    Activity_2: str
    Activity_3: str
    Cash_Flow_From_Investing_Activities: str

    Borrowings_net: str
    Shortterm_Borrowings_from_Parent_Net: str
    Related_Party_Loans_Net: str
    Debt_Issuance_Costs: str
    Net_Transfer_to_Parent: str
    Dividends: str
    Other: str
    Cash_Flow_From_Financing_Activities: str

    Foreign_Exchange_Rate_Effect_on_Cash_and_Cash_Equivalents: str
    Discontinued_Operations_Cash_Flows: str
    Discontinued_Operations_Cash_Balance: str
    Change_In_Cash_Cash_Equiv: str
    Cash_And_Cash_Equivalents_Beginning_Of_Period: str
    Cash_And_Cash_Equivalents_End_Of_Period: str

    Current_Assets: str
    Cash_and_Cash_Equivalents: str
    Accounts_Receivables: str
    Inventories: str
    Other_Current_Assets: str
    Current_Assets_of_Discontinued_Operations: str
    Total_Current_Assets: str

    Property_Plant_and_Equipment_Net: str
    Goodwill: str
    Other_Intangiblesnet: str
    Other_Assets: str
    Noncurrent_Assets_of_Discontinued_Operations: str
    Total_Assets: str

    Current_Liabilities: str
    Trade_Accounts_Payable: str
    Accrued_and_Other_Current_Liabilities: str
    Due_to_Related_Party: str
    Income_Taxes_Payable: str
    Current_Liabilities_of_Discontinued_Operations: str
    Current_Portion_of_Debt: str
    Total_Current_Liabilities: str
    
    Long_Term_Debt: str
    Deferred_Income_Taxes: str
    Other_Noncurrent_Liabilities: str
    Noncurrent_Liabilities_of_Discontinued_Operation: str
    Total_Liabilities: str

    Shareholder_Equity: str
    Total_Liabilities_And_Shareholders_Equities: str
