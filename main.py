import os
import pandas as pd

pd.options.display.max_columns = 25

fd = os.path.join('~', 'Dropbox', 'Haaf Transcontinental', 'Hamilton', 'Pont Neuf - Hamilton Data')
fn = '__MASTER Sales Summary and Daily Wraps Doc.xlsx'


def read_MASTER(fd_f, fn_f, sheet_f):
    df_f = pd.read_excel(io=os.path.join(fd_f, fn_f), sheet_name=sheet_f, header=2)

    t_colnames = list(df_f.columns[1:])
    t_colnames.insert(0, 'notes')
    df_f.columns = t_colnames

    df_f.Date = pd.to_datetime(df_f.Date, errors='coerce').dt.date

    df_f.drop(df_f.loc[df_f.Date.isnull()].index, inplace=True)
    df_f.drop(columns=['Total', 'CUMULATIVE TOTAL'], inplace=True)
    return df_f


def read_type1(fd_f, fn_f, sheet_f, header_row_ind_f):
    df_f = pd.read_excel(io=os.path.join(fd_f, fn_f), sheet_name=sheet_f, header=header_row_ind_f)

    t_colnames = list(df_f.columns[2:])
    t_colnames.insert(0, 'notes')
    t_colnames.insert(0, 'Date')
    df_f.columns = t_colnames

    df_f.Date = pd.to_datetime(df_f.Date, errors='coerce').dt.date

    df_f.drop(df_f.loc[df_f.Date.isnull()].index, inplace=True)
    t_dropcol_candidates = {'Total', 'CUMULATIVE TOTAL', 'Average Ticket Price'}
    t_dropcols = [x for x in df_f.columns if x.find('Unnamed') > -1]
    t_dropcols = t_dropcols + list(t_dropcol_candidates)   #### Set intersection with existing columns
    df_f.drop(columns=t_dropcols, inplace=True)

    return df_f


def read_type2(fd_f, fn_f, sheet_f):
    # get params
    df_f = pd.read_excel(io=os.path.join(fd_f, fn_f), sheet_name=sheet_f, header=None, nrows=4)
    params_f = dict(zip(df_f.iloc[:, 5], df_f.iloc[:, 7]))
    del df_f

    # get sales
    t_cols = ['Date', 'Time', 'Total Tickets Sold', 'Total Gross', 'Remaining Inventory']
    df_f = pd.read_excel(io=os.path.join(fd_f, fn_f), sheet_name=sheet_f, header=5, usecols=t_cols)

    # break into two components: total and Box Office
    t_breakpoint = int(df_f.loc[df_f.Date == 'Date'].index[0])
    df_f_all = df_f.iloc[0:t_breakpoint, ].reindex()
    df_f_all.rename(columns={'Remaining Inventory': 'Total Remaining Inventory'}, inplace=True)

    df_f_bo = df_f.iloc[t_breakpoint:df_f.shape[0], ].reindex()
    df_f_bo.columns = df_f_bo.iloc[0]
    df_f_bo.drop(index=df_f_bo.index[0], inplace=True)
    df_f_bo.rename(columns={'Remaining Inventory': 'BO Remaining Inventory'}, inplace=True)

    def clean_date_time(df_f_f):
        end_of_day_labels = ['end of day', 'midnight']  # lower case
        df_f_f.Date = pd.to_datetime(df_f_f.Date, errors='coerce').dt.date
        df_f_f.drop(df_f_f.loc[df_f_f.Date.isnull()].index, inplace=True)
        df_f_f.loc[df_f_f.Time.astype(str).str.lower().isin(end_of_day_labels), 'Time'] = '23:59:59'
        df_f_f.Time = pd.to_datetime(df_f_f.Time, format='%H:%M:%S').dt.time
        return df_f_f

    df_f_all = clean_date_time(df_f_all)
    df_f_bo = clean_date_time(df_f_bo)

    df_ff = pd.merge(df_f_all, df_f_bo, how='outer', on=['Date', 'Time'])

    if (df_f_all.shape[0] != df_ff.shape[0]) or (df_f_bo.shape[0] != df_ff.shape[0]):
        print('Warning: Daily wraps files had row expansion merging aggregate and box office data for file %s' % fn_f)

    for t_key, t_value in params_f.items():
        df_ff[t_key] = t_value

    return {fn_f: {'params': params_f, 'data': df_ff}}


# All sales
df_master = read_MASTER(fd_f=fd, fn_f='__MASTER Sales Summary and Daily Wraps Doc.xlsx', sheet_f='Daily Wrap')
df_MfMyers = read_type1(fd, fn_f='HAMILTON Ft. Myers Sales Summary.xlsx', sheet_f='Ft. Myers - Daily Wrap', header_row_ind_f=3)
df_GrandRapids = read_type1(fd, fn_f='HAMILTON Grand Rapids Sales Summary.xlsx', sheet_f='Daily Wraps', header_row_ind_f=3)
df_Indianapolis = read_type1(fd, fn_f='HAMILTON Indianapolis 2019 Sales Summary.xlsx', sheet_f='IND Daily Wraps', header_row_ind_f=4)
df_LosAngeles = read_type1(fd, fn_f='Hamilton LA Sales Summary.xlsx', sheet_f='LA - Daily Sales', header_row_ind_f=4)
df_Miami = read_type1(fd, fn_f='Hamilton Miami Sales Summary.xlsx', sheet_f='Daily Wraps - FINAL', header_row_ind_f=4)
df_Naples = read_type1(fd, fn_f='HAMILTON Naples Summary.xlsx', sheet_f='Naples- Daily Wrap', header_row_ind_f=3)
df_Nashville = read_type1(fd, fn_f='Hamilton Nashville Sales Summary.xlsx', sheet_f='Nashville Daily Wraps', header_row_ind_f=3)
df_Norfolk = read_type1(fd, fn_f='HAMILTON Norfolk Sales Summary.xlsx', sheet_f='Norfolk - Daily Wrap', header_row_ind_f=3)
df_Philadelphia = read_type1(fd, fn_f='HAMILTON Philadelphia Sales Summary.xlsx', sheet_f='Daily Wrap', header_row_ind_f=2)
df_Richmond = read_type1(fd, fn_f='HAMILTON Richmond Sales Summary.xlsx', sheet_f='Richmond - Daily Wrap', header_row_ind_f=3)
df_Toronto = read_type1(fd, fn_f='Hamilton Toronto Sales Summary.xlsx', sheet_f='Daily Wraps', header_row_ind_f=5)
df_WestPalmBeach = read_type1(fd, fn_f='Hamilton West Palm Beach Sales Summary.xlsx', sheet_f='Daily Wraps', header_row_ind_f=4)

# First day sales
fn_first_day_sales = {'Hamilton_Atlanta II OnSale_WrapReporting 12.2.19.xlsx': {'city': 'Atlanta', 'state': 'GA'},
                      'Hamilton_FtMyers_OnSale_Wraps 11.2.19 with total.xlsx': {'city': 'Fort Meyers', 'state': 'FL'},
                      'Hamilton_Nashville OnSale_WrapReporting 11.11.19.xlsx': {'city': 'Nashville', 'state': 'TN'},
                      'Indianapolis_Hamilton_OnSale_Wrap_Reporting 10.17.19 Final .xlsx': {'city': 'Indianapolis', 'state': 'IN'},
                      'MAD Hamilton OnSale Wrap Reporting- FINAL.xlsx': {'city': 'Madison', 'state': 'WI'},
                      'Milwaukee_Hamilton_OnSale_Wrap_Reporting 9.10.19.xlsx': {'city': 'Milwaukee', 'state': 'WI'},
                      'Norfolk_Hamilton_OnSale_Wrap_Reporting_9.27.19 FINAL.xlsx': {'city': 'Norfolk', 'state': 'VA'},
                      'Richmond_Hamilton_OnSale_Wrap_Reporting 9.27.19 Final.xlsx': {'city': 'Richmond', 'state': 'VA'}
                      }
t_dict_out = dict()
for t_fn, _ in fn_first_day_sales.items():
    print('File: %s' % t_fn)
    t_dict_out.update(read_type2(fd_f=os.path.join(fd, 'On Sale Wraps'), fn_f=t_fn, sheet_f='Hamilton Sales By Hour'))
del t_fn

# create one data frame of all daily wraps sales
def flatten_dict(t_key_f, t_dict_out_f):
    t_df = t_dict_out_f[t_key_f]['data'].reindex()
    t_df['city'] = fn_first_day_sales[t_key_f]['city']
    t_df['state'] = fn_first_day_sales[t_key_f]['state']
    t_df['source'] = t_key_f
    return t_df
df_first_day = flatten_dict(t_key_f=next(iter(t_dict_out)), t_dict_out_f=t_dict_out)
df_first_day.drop(df_first_day.index, inplace=True)

for t_fn, _ in fn_first_day_sales.items():
    df_first_day = pd.concat([df_first_day, flatten_dict(t_key_f=t_fn, t_dict_out_f=t_dict_out)], axis=0, ignore_index=True)

print(df_first_day.groupby('city')['Total Gross', 'BO Gross'].sum())