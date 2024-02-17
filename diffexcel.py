import argparse
import pandas as pd
import sys, os
#import xlsxwriter
import time
from pathlib import Path
from datetime import datetime
#from multiprocessing import Process

def excel_diff(df_OLD, df_NEW, path_OLD, path_NEW, sheet):

    print("\nStart Time :", datetime.now().strftime("%H:%M:%S"))
    start_time = time.perf_counter()
    ## Perform Diff
    #newRows = []
    report = []
    
    rows_OLD = df_OLD.index
    rows_NEW = df_NEW.index
    rowsCombined = list(set(rows_OLD).union(rows_NEW))
    cols_OLD = df_OLD.columns
    cols_NEW = df_NEW.columns
    colsCombined = list(set(cols_OLD).union(cols_NEW))
    
    ## Compare Rows
    for row in rowsCombined:
        if not(row in rows_OLD):
            report.append('diff: !{sheet}({row},XXX): new row'.format(row=row,sheet=sheet))
        elif not(row in rows_NEW):
            report.append('diff: !{sheet}({row},XXX): missing row'.format(row=row,sheet=sheet))
        else:
            for col in colsCombined:
            	if not(col in cols_OLD):
            	    report.append('diff: !{sheet}(XXX,{col}): new col ({val})'.format(col=col,sheet=sheet,val=df_NEW.loc[row,col]))
            	elif not(col in cols_NEW):
            	    report.append('diff: !{sheet}(XXX,{col}): missing col ({val})'.format(col=col,sheet=sheet,val=df_OLD.loc[row,col]))
            	else:
            	    value_OLD = df_OLD.loc[row,col]
            	    value_NEW = df_NEW.loc[row,col]
            	    if value_OLD==value_NEW:
            	        #dfDiff.loc[row,col] = df_NEW.loc[row,col]
            	        pass
            	    else:
            	        #dfDiff.loc[row,col] = ('{}→{}').format(value_OLD,value_NEW)
            	        #dfDiff.loc[last_report_row,0] = 'diff: !{sheet}({row},{col}): {value_OLD} -> {value_NEW}'.format(value_OLD=value_OLD,value_NEW=value_NEW,row=row,col=col,sheet=sheet)
            	        report.append('diff: !{sheet}({row},{col}): {value_OLD} -> {value_NEW}'.format(value_OLD=value_OLD,value_NEW=value_NEW,row=row,col=col,sheet=sheet))
    #    else:
    #        newRows.append(row)


    print("\nProcessing time:", datetime.now().strftime("%H:%M:%S"))
    return report

def main(args):

    print("\nStart Time of Program :", datetime.now().strftime("%H:%M:%S"), "\n")
    start_time = time.perf_counter()

    ## Set Variables
    indexColName = None # args.index
    path_OLD = Path(args.old_excel)
    path_NEW = Path(args.new_excel)
 
    ### Starting Multiprocessing for Environment sheets
    #procs = []
    
    report = []
    
    sheets_OLD = pd.ExcelFile(path_OLD,engine='openpyxl').sheet_names
    sheets_NEW = pd.ExcelFile(path_OLD,engine='openpyxl').sheet_names
    sheetsCombined = list(set(sheets_OLD).union(sheets_NEW))
    
    #for (env, sheet_num) in zip(["PROD", "TEST"],[0, 1]):
    #for (env, sheet_num) in zip(["PROD"],sheets):
    env = "PROD"
    for sheet_num in sheetsCombined:
        
        ## Reading Sheets from Excel files
        print("\nReading sheet `{sheet}` for {env} ...".format(env=env,sheet=sheet_num))
        df_OLD = pd.read_excel(path_OLD, sheet_name=sheet_num, index_col=indexColName,engine='openpyxl').fillna(0)
        df_NEW = pd.read_excel(path_NEW, sheet_name=sheet_num, index_col=indexColName,engine='openpyxl').fillna(0)
        
        #proc = Process(target=excel_diff, args=(df_OLD, df_NEW, path_OLD, path_NEW, sheet_num, env))
        #procs.append(proc)
        #proc.start()
        if not(sheet_num in sheets_OLD):
            report.append('diff: !{sheet}: new'.format(sheet=sheet_num))
        elif not(sheet_num in sheets_NEW):
            report.append('diff: !{sheet}: missing'.format(sheet=sheet_num))
        else:
            report = report + excel_diff(df_OLD, df_NEW, path_OLD, path_NEW, sheet_num)

    ### Complete the processes
    #for proc in procs:
    #    proc.join()
    
    print("\nSaving the outputs in new excel for Current Time: ", datetime.now().strftime("%H:%M:%S")) 
    
     
    ## Save output and format
    fname = 'diff {} vs {}.txt'.format(path_OLD.stem,path_NEW.stem)
    if os.path.isfile(fname):
        os.remove(fname)
    #writer = pd.ExcelWriter(fname, engine='xlsxwriter')
    #
    #dfDiff.to_excel(writer, sheet_name='DIFF', index=True)
    ##df_NEW.to_excel(writer, sheet_name=path_NEW.stem, index=True)
    ##df_OLD.to_excel(writer, sheet_name=path_OLD.stem, index=True)
    #
    ### Get xlsxwriter objects
    #workbook  = writer.book
    #worksheet = writer.sheets['DIFF']
    #worksheet.set_default_row(15)
    #
    ### Define formats
    #highlight_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color':'#B1B3B3'})
    #new_fmt = workbook.add_format({'bg_color': '#32CD32', 'bold':True})
    #
    #
    ### Highlight changed cells
    #worksheet.conditional_format('A1:AZ100000', {'type': 'text',
    #                                        'criteria': 'containing',
    #                                        'value':'→',
    #                                        'format': highlight_fmt})
    ### Highlight new rows
    #for row, row_data in enumerate(dfDiff.index):
    #    if row_data in newRows:
    #        worksheet.set_row(row+1, 15, new_fmt)
    #
    ### Saving Workbook
    #writer.save()
    #df = pd.DataFrame(report, columns=['Results'])
    #df.to_excel(fname, sheet_name='diff')
    f_out = open(fname, "w")
    f_out.writelines(['{s}{linebreak}'.format(s=s,linebreak="\n") for s in report])
    #for line in report:
   #     f_out.write(line)
    f_out = None
    print("\nSaved Excel sheet in {fname} .".format(fname=fname))
    print("\nEnd Time: {}".format(datetime.now().strftime("%H:%M:%S")))
    
    
    ## Calculate Total Run time
    end_time = time.perf_counter()
    total_process_time = '{:0.2f}'.format(((end_time - start_time)/60))
    print(f"\nTotal Run Time: {total_process_time} minutes")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Compare Two Excel Workbooks."
    )
    parser.add_argument(
        '-o',
        '--old-excel',
        metavar='old_file.xlsx',
        help='Old Excel File path',
        required=True
    )
    parser.add_argument(
        '-n',
        '--new-excel',
        metavar='new_file.xlsx',
        help='New Excel file path',
        required=True
    )
    #parser.add_argument(
    #    '-i',
    #    '--index',
    #    metavar='Account Number',
    #    help='Common Index Column',
    #    required=True
    #)
    args = parser.parse_args()
    main(args)