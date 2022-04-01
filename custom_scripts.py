'''
Module using Calc Service ScriptForge Methods to calculate the monthly subscriptions	
	
## Links of interest
- [Introduction](https://blog.documentfoundation.org/blog/2021/01/28/introducing-the-scriptforge-basic-libraries/)
- [Documentation](https://help.libreoffice.org/7.2/en-US/text/sbasic/shared/03/lib_ScriptForge.html)
- [Examples](https://github.com/rafaelhlima/LibOCon_2021_SFCalc)
- [Demonstration of examples](https://www.youtube.com/watch?v=pHlLdyJz2sE)

- [Old interface](https://www.youtube.com/watch?v=g8I8uGjaXA8)
	'''


from scriptforge import CreateScriptService
import random as rnd

from datetime import datetime

MACRO_SHEET_PREFIX="Meta."
MACRO_VAR_NAME_COLUMN="D"
MACRO_VAR_VALUE_COLUMN="E"
NUMBER_OF_MACROS=20

# Initialize macro values
COST_COLUMN = None
PAYMENT_SPAN_COLUMN = None
TARGET_COLUMN = None
TIME_COLUMN = None
FIRST_SUBSCRIPTION_ROW = None
LAST_SUBSCRIPTION_ROW = None

def debug_print(args=None):
    doc = CreateScriptService("Calc")
    bas = CreateScriptService("Basic")
    result = doc.getValue("A1:E1")
    bas.MsgBox(str(result))

def test_read_macros(args=None):
    doc = CreateScriptService("Calc")
 
    read_macros(doc)

    print_macros()

def read_macros(doc):
    # bas = CreateScriptService("Basic")

    macro_names_cells = MACRO_SHEET_PREFIX + MACRO_VAR_NAME_COLUMN + "1:" + MACRO_VAR_NAME_COLUMN + str(NUMBER_OF_MACROS)

    # bas.MsgBox(macro_names_cells)
    macro_names = doc.getValue(macro_names_cells)

    macro_values_cells = MACRO_SHEET_PREFIX + MACRO_VAR_VALUE_COLUMN + "1:" + MACRO_VAR_VALUE_COLUMN + str(NUMBER_OF_MACROS)

    macro_values = doc.getValue(macro_values_cells)

    # bas.MsgBox(str(macro_spaced))

    global COST_COLUMN
    global PAYMENT_SPAN_COLUMN
    global TARGET_COLUMN
    global TIME_COLUMN
    global FIRST_SUBSCRIPTION_ROW
    global LAST_SUBSCRIPTION_ROW

    for index,name in enumerate(macro_names, start=1):
        if name == "cost_column":
            COST_COLUMN = doc.getValue(MACRO_SHEET_PREFIX + MACRO_VAR_VALUE_COLUMN+str(index))
        elif name == "payment_span_column":
            PAYMENT_SPAN_COLUMN = doc.getValue(MACRO_SHEET_PREFIX + MACRO_VAR_VALUE_COLUMN+str(index))
        elif name == "target_column":
            TARGET_COLUMN = doc.getValue(MACRO_SHEET_PREFIX + MACRO_VAR_VALUE_COLUMN+str(index))
        elif name == "time_column":
            TIME_COLUMN = doc.getValue(MACRO_SHEET_PREFIX + MACRO_VAR_VALUE_COLUMN+str(index))
        elif name == "first_subscription_row":
            FIRST_SUBSCRIPTION_ROW = doc.getValue(MACRO_SHEET_PREFIX + MACRO_VAR_VALUE_COLUMN+str(index))
        elif name == "last_subscription_row":
            LAST_SUBSCRIPTION_ROW = doc.getValue(MACRO_SHEET_PREFIX + MACRO_VAR_VALUE_COLUMN+str(index))

def print_macros():
    bas = CreateScriptService("Basic")

    bas.MsgBox("Captured cost_column is: " + str(COST_COLUMN) \
    + "\nCaptured payment_span_column is: " + str(PAYMENT_SPAN_COLUMN) \
    + "\nCaptured target_column is: " + str(TARGET_COLUMN) \
    + "\nCaptured time_column is: " + str(TIME_COLUMN) \
    + "\nCaptured first_subscription_row is: " + str(FIRST_SUBSCRIPTION_ROW) \
    + "\nCaptured last_subscription_row is: " + str(LAST_SUBSCRIPTION_ROW))

def subscriptions_monthlies(args=None):
    doc = CreateScriptService("Calc")

    read_macros(doc)

    # print_macros()

    total_subscriptions()
    non_installment_subscriptions()
    house_subscriptions()
    car_subscriptions()
    child_edu_subscriptions()
    total_subscriptions_with_variables()

def total_subscriptions():
    CONDITIONED=True
    CONDITION_COLUMN="C"
    CONDITION_VALUE=""
    TARGET_ROW="2"

    def run_condition(doc, condition_cell):
        return basic_condition_uncheck(doc, condition_cell, ["House Needs", "Car Needs"])

    calculate_total(FIRST_SUBSCRIPTION_ROW, LAST_SUBSCRIPTION_ROW, COST_COLUMN, PAYMENT_SPAN_COLUMN, CONDITIONED, CONDITION_COLUMN, run_condition, TARGET_ROW, TARGET_COLUMN, TIME_COLUMN)

def non_installment_subscriptions():
    CONDITIONED=True
    CONDITION_COLUMN="C"
    TARGET_ROW="3"

    def run_condition(doc, condition_cell):
        return basic_condition_uncheck(doc, condition_cell, ["House Needs", "Car Needs"])

    calculate_total(FIRST_SUBSCRIPTION_ROW, LAST_SUBSCRIPTION_ROW, COST_COLUMN, PAYMENT_SPAN_COLUMN, CONDITIONED, CONDITION_COLUMN, run_condition, TARGET_ROW, TARGET_COLUMN, TIME_COLUMN)

def house_subscriptions():
    CONDITIONED=True
    CONDITION_COLUMN="C"
    TARGET_ROW="4"

    def run_condition(doc, condition_cell):
        return basic_condition_check(doc, condition_cell, ["House Utility"])

    calculate_total(FIRST_SUBSCRIPTION_ROW, LAST_SUBSCRIPTION_ROW, COST_COLUMN, PAYMENT_SPAN_COLUMN, CONDITIONED, CONDITION_COLUMN, run_condition, TARGET_ROW, TARGET_COLUMN, TIME_COLUMN)

def car_subscriptions():
    CONDITIONED=True
    CONDITION_COLUMN="C"
    TARGET_ROW="5"

    def run_condition(doc, condition_cell):
        return basic_condition_check(doc, condition_cell, ["Car Utility"])

    calculate_total(FIRST_SUBSCRIPTION_ROW, LAST_SUBSCRIPTION_ROW, COST_COLUMN, PAYMENT_SPAN_COLUMN, CONDITIONED, CONDITION_COLUMN, run_condition, TARGET_ROW, TARGET_COLUMN, TIME_COLUMN)

def child_edu_subscriptions():
    CONDITIONED=True
    CONDITION_COLUMN="C"
    TARGET_ROW="6"

    def run_condition(doc, condition_cell):
        return basic_condition_check(doc, condition_cell, ["Child Education"])

    calculate_total(FIRST_SUBSCRIPTION_ROW, LAST_SUBSCRIPTION_ROW, COST_COLUMN, PAYMENT_SPAN_COLUMN, CONDITIONED, CONDITION_COLUMN, run_condition, TARGET_ROW, TARGET_COLUMN, TIME_COLUMN)

def total_subscriptions_with_variables():
    CONDITIONED=False
    CONDITION_COLUMN="C"
    TARGET_ROW="10"

    def run_condition(doc, condition_cell):
        return True

    calculate_total(FIRST_SUBSCRIPTION_ROW, LAST_SUBSCRIPTION_ROW, COST_COLUMN, PAYMENT_SPAN_COLUMN, CONDITIONED, CONDITION_COLUMN, run_condition, TARGET_ROW, TARGET_COLUMN, TIME_COLUMN)

def basic_condition_check(doc, condition_cell, values):
    VUT = doc.getValue(condition_cell)
    for value in values:
        if (VUT != value):
            return False
    return True

def basic_condition_uncheck(doc, condition_cell, values):
    VUT = doc.getValue(condition_cell)
    for value in values:
        if (VUT == value):
            return False
    return True

def calculate_total(FIRST_CONSIDERED_SUBSCRIPTION, LAST_CONSIDERED_SUBSCRIPTION, DIVIDEND_COLUMN, DIVISOR_COLUMN, CONDITIONED, CONDITION_COLUMN, run_condition, TARGET_ROW, TARGET_COLUMN, TIME_COLUMN):
    doc = CreateScriptService("Calc")
    target_cell = doc.Offset(TARGET_COLUMN+str(TARGET_ROW), 0, 0)
    sum = 0
    for i in range(LAST_CONSIDERED_SUBSCRIPTION):
        dividend = doc.Offset(DIVIDEND_COLUMN+str(i+FIRST_CONSIDERED_SUBSCRIPTION), 0, 0)
        divisor = doc.Offset(DIVISOR_COLUMN+str(i+FIRST_CONSIDERED_SUBSCRIPTION), 0, 0)
        condition_cell = doc.Offset(CONDITION_COLUMN+str(i+FIRST_CONSIDERED_SUBSCRIPTION), 0, 0)
        if ((not CONDITIONED or run_condition(doc, condition_cell)) and is_cell_a_number(doc, dividend) and is_cell_a_nonzero_number(doc, divisor)):
            sum = sum + doc.getValue(dividend)/doc.getValue(divisor)

    doc.setValue(target_cell, round(sum, 2))

    date_cell = doc.Offset(TIME_COLUMN+str(TARGET_ROW), 0, 0)
    now = datetime.now()
    doc.setValue(date_cell, str(now))

def is_cell_a_number(doc, cell):
    val = doc.getValue(cell)
    return isinstance(val, int) or isinstance(val, float) 

def is_cell_a_nonzero_number(doc, cell):
    val = doc.getValue(cell)
    return (isinstance(val, int) or isinstance(val, float)) and (val != 0)

g_exportedScripts = (
    subscriptions_monthlies,
    debug_print,
    test_read_macros,
	)