
import win32com.client as win32

def clear_pts(ws):
    for pt in ws.PivotTables():
        pt.TableRange2.Clear()

def insert_pt_filed_set1(pt):
    field_rows={}
    field_rows['Customer code']=pt.PivotFields("Customer Code")
    field_rows['Customer']=pt.PivotFields("Customer")

    field_values={}
    field_values['total']=pt.PivotFields("Amount")

    field_filter={}
    field_filter["filter"]=pt.PivotFields("G/L Account")

    filed_column={}
    filed_column['age']=pt.PivotFields("aging")
    
    field_rows['Customer code'].Orientation=1
    field_rows['Customer code'].Position=1

    field_rows['Customer'].Orientation=1
    field_rows['Customer'].Position=2

    #insert data fileds
    field_values['total'].Orientation=4
    field_values['total'].Function=-4157#-4112
##    field_values['total'].NumerFormat="#,##0"    #for sum function- -4157
    
    #inserting filter
    field_filter["filter"].Orientation=3

    #inserting column
    filed_column['age'].Orientation=2
    
    

    
# Open the Excel file
xlapp = win32.gencache.EnsureDispatch('Excel.Application')
xlapp.Visible=True
wb = xlapp.Workbooks.Open(r"D:\BASIC_Tech\Excel_python\pivot_table_win32.xlsx")
ws = wb.Sheets('Sheet1')
ws_report=wb.Worksheets("report")
clear_pts(ws_report)
### Create a pivot table
pt_cache=wb.PivotCaches().Create(1,ws.Range("A1").CurrentRegion)
pt=pt_cache.CreatePivotTable(ws_report.Range("B3"),"My_report summary")

pt.ColumnGrand=True
pt.RowGrand=True


##pt.SubtotalLocation(1)
pt.RowAxisLayout(1)
pt.PivotFields('Customer Code').Subtotals = tuple(False for _ in range(12))
#chnage pivot table style
pt.TableStyle2="PivotStyleMedium9"

#creating report
insert_pt_filed_set1(pt)


wb.Save()
xlapp.Application.Quit()
##
