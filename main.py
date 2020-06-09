import os
import pandas as pd

pd.options.display.max_columns = 25

fd = os.path.join('~', 'Dropbox', 'Haaf Transcontinental', 'Hamilton', 'Pont Neuf - Hamilton Data')
fn = '__MASTER Sales Summary and Daily Wraps Doc.xlsx'

# ----- Define functions -----
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
    t_dropcol_candidates = {'CUMULATIVE TOTAL', 'Average Ticket Price'} #'Total',
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


def import_daily_wraps_data(t_fn_f, t_atts_f):
    df_f = read_type1(fd, fn_f=t_fn_f, sheet_f=t_atts_f['sheet'], header_row_ind_f=t_atts_f['header_row_ind'])
    df_f.columns = [x.lower() for x in df_f.columns]
    df_f['city'] = t_atts_f['city']
    df_f['state'] = t_atts_f['state']
    df_f['source'] = t_fn_f
    return df_f


# ---- Import hardcoded info -----
fn_daily_wraps = {'HAMILTON Ft. Myers Sales Summary.xlsx': {'sheet': 'Ft. Myers - Daily Wrap', 'header_row_ind': 3, 'city': 'Fort Myers', 'state': 'FL'},
                  'HAMILTON Grand Rapids Sales Summary.xlsx': {'sheet': 'Daily Wraps', 'header_row_ind': 3, 'city': 'Grand Rapids', 'state': 'MI'},
                  'HAMILTON Indianapolis 2019 Sales Summary.xlsx':  {'sheet': 'IND Daily Wraps', 'header_row_ind': 4, 'city': 'Indianapolis', 'state': 'IN'},
                  'Hamilton LA Sales Summary.xlsx':  {'sheet': 'LA - Daily Sales', 'header_row_ind': 4, 'city': 'Los Angeles', 'state': 'CA'},
                  'Hamilton Miami Sales Summary.xlsx':  {'sheet': 'Daily Wraps - FINAL', 'header_row_ind': 4, 'city': 'Miami', 'state': 'FL'},
                  'HAMILTON Naples Summary.xlsx':  {'sheet': 'Naples- Daily Wrap', 'header_row_ind': 3, 'city': 'Naples', 'state': 'FL'},
                  'Hamilton Nashville Sales Summary.xlsx':  {'sheet': 'Nashville Daily Wraps', 'header_row_ind': 3, 'city': 'Nashville', 'state': 'TN'},
                  'HAMILTON Norfolk Sales Summary.xlsx':  {'sheet': 'Norfolk - Daily Wrap', 'header_row_ind': 3, 'city': 'Norfolk', 'state': 'VA'},
                  'HAMILTON Philadelphia Sales Summary.xlsx':  {'sheet': 'Daily Wrap', 'header_row_ind': 2, 'city': 'Philadelphia', 'state': 'PA'},
                  'HAMILTON Richmond Sales Summary.xlsx':  {'sheet': 'Richmond - Daily Wrap', 'header_row_ind': 3, 'city': 'Richmond', 'state': 'VA'},
                  'Hamilton Toronto Sales Summary.xlsx':  {'sheet': 'Daily Wraps', 'header_row_ind': 5, 'city': 'Toronto', 'state': 'CN'},
                  'Hamilton West Palm Beach Sales Summary.xlsx':  {'sheet': 'Daily Wraps', 'header_row_ind': 4, 'city': 'West Palm Beach', 'state': 'FL'}
                  }

fn_daily_wraps_meta = {'HAMILTON Ft. Myers Sales Summary.xlsx': {'cap_per_perf': 1831, 'weekly_cap': 14648, 'num_weeks': 2, 'num_sub_weeks': None, 'num_perf': 16, 'open_date': pd.to_datetime('1/13/20', format='%m/%d/%y')},
                       'HAMILTON Grand Rapids Sales Summary.xlsx': {'cap_per_perf': 2293, 'weekly_cap': 18344, 'num_weeks': 3, 'num_sub_weeks': None, 'num_perf': 25, 'open_date': pd.to_datetime('1/27/20', format='%m/%d/%y')},
                       'HAMILTON Indianapolis 2019 Sales Summary.xlsx':  {'cap_per_perf': 2589, 'weekly_cap': None, 'num_weeks': 3, 'num_sub_weeks': 1, 'num_perf': 24, 'open_date': pd.to_datetime('12/9/19', format='%m/%d/%y')},
                       'Hamilton LA Sales Summary.xlsx':  {'cap_per_perf': 2703, 'weekly_cap': 21624, 'num_weeks': 37, 'num_sub_weeks': None, 'num_perf': None, 'open_date': pd.to_datetime('3/16/20', format='%m/%d/%y')},
                       'Hamilton Miami Sales Summary.xlsx':  {'cap_per_perf': 2232, 'weekly_cap': 17856, 'num_weeks': 4, 'num_sub_weeks': None, 'num_perf': 32, 'open_date': pd.to_datetime('2/24/20', format='%m/%d/%y')},
                       'HAMILTON Naples Summary.xlsx':  {'cap_per_perf': 1381, 'weekly_cap': 11048, 'num_weeks': 2, 'num_sub_weeks': None, 'num_perf': 16, 'open_date': pd.to_datetime('1/6/20', format='%m/%d/%y')},
                       'Hamilton Nashville Sales Summary.xlsx':  {'cap_per_perf': 2425, 'weekly_cap': 19400, 'num_weeks': 3, 'num_sub_weeks': None, 'num_perf': 24, 'open_date': pd.to_datetime('1/6/20', format='%m/%d/%y')},
                       'HAMILTON Norfolk Sales Summary.xlsx':  {'cap_per_perf': 2317, 'weekly_cap': None, 'num_weeks': 3, 'num_sub_weeks': 1, 'num_perf': 24, 'open_date': pd.to_datetime('12/16/19', format='%m/%d/%y')},
                       'HAMILTON Philadelphia Sales Summary.xlsx':  {'cap_per_perf': None, 'weekly_cap': None, 'num_weeks': None, 'num_sub_weeks': None, 'num_perf': None, 'open_date': pd.to_datetime('9/2/19', format='%m/%d/%y')},
                       'HAMILTON Richmond Sales Summary.xlsx':  {'cap_per_perf': 3399, 'weekly_cap': None, 'num_weeks': 3, 'num_sub_weeks': 1, 'num_perf': 24, 'open_date': pd.to_datetime('11/25/19', format='%m/%d/%y')},
                       'Hamilton Toronto Sales Summary.xlsx':  {'cap_per_perf': 2229, 'weekly_cap': 17408, 'num_weeks': 14, 'num_sub_weeks': 6, 'num_perf': 12, 'open_date': pd.to_datetime('2/3/20', format='%m/%d/%y')},
                       'Hamilton West Palm Beach Sales Summary.xlsx':  {'cap_per_perf': 2063, 'weekly_cap': 16504, 'num_weeks': 3, 'num_sub_weeks': None, 'num_perf': 24, 'open_date': pd.to_datetime('2/3/20', format='%m/%d/%y')}
                       }

# ---- Aggregate daily wraps data -----
df_daily_wrap = read_MASTER(fd_f=fd, fn_f='__MASTER Sales Summary and Daily Wraps Doc.xlsx', sheet_f='Daily Wrap')
df_daily_wrap['city'] = None
df_daily_wrap['state'] = None
df_daily_wrap['source'] = None
df_daily_wrap.drop(df_daily_wrap.index, inplace=True)
df_daily_wrap.columns = [x.lower() for x in df_daily_wrap.columns]
df_city_daily_wrap_dict = {'__MASTER Sales Summary and Daily Wraps Doc.xlsx': df_daily_wrap}
for t_fn, t_atts in fn_daily_wraps.items():
    t_df = import_daily_wraps_data(t_fn_f=t_fn, t_atts_f=t_atts)
    df_city_daily_wrap_dict.update({t_fn: t_df})
    df_daily_wrap = pd.concat([df_daily_wrap, t_df], axis=0, ignore_index=True)
rev_exclude_cols = ['date', 'city', 'state', 'source', 'total', 'notes', 'total tix', 'comps']
df_daily_wrap['total_from_cat'] = df_daily_wrap[[x for x in df_daily_wrap.columns if not any([x == y for y in rev_exclude_cols])]].sum(axis=1)
df_daily_wrap['total_diff'] = df_daily_wrap['total_from_cat'] - df_daily_wrap['total']

# field summary by city
df_city_field_summary = pd.DataFrame(index=list(df_daily_wrap.columns.unique()), columns=list(df_daily_wrap.city.unique()))
for r in df_city_field_summary.index:
    df_city_field_summary.loc[r, df_daily_wrap.loc[df_daily_wrap[r].notnull(), 'city'].unique()] = 1
df_city_field_summary.to_csv(os.path.join('intermediate_data', 'fields_by_city.csv'), index=True)

# city meta information
t_rows = list()
for t_key, t_value in fn_daily_wraps_meta.items():
    t_value.update({'city': fn_daily_wraps[t_key]['city'], 'state': fn_daily_wraps[t_key]['state']})
    t_rows.append(t_value)
df_city_stats = pd.DataFrame.from_dict(t_rows, orient='columns')
del t_rows, t_value, t_key

# create summary output
df_review = df_city_stats.reindex()
# min/max dates
df_review = df_review.merge(pd.DataFrame(df_daily_wrap.groupby('city')['date'].min().reset_index()).rename(columns={'date': 'first_sales_date'}), how='left', on='city')
df_review = df_review.merge(pd.DataFrame(df_daily_wrap.groupby('city')['date'].max().reset_index()).rename(columns={'date': 'last_sales_date'}), how='left', on='city')
df_review['first_sales_date'] = pd.to_datetime(df_review['first_sales_date'], format='%Y-%m-%d')
df_review['last_sales_date'] = pd.to_datetime(df_review['last_sales_date'], format='%Y-%m-%d')
# number of weeks out
df_review['first_rel_sales_date_weeks'] = round((df_review['first_sales_date'] - df_review['open_date']).dt.days/7)
df_review['last_rel_sales_date_weeks'] = round((df_review['last_sales_date'] - df_review['open_date']).dt.days/7)
# capacity per perf
# weekly capacity
# weeks
# performances
# gross potential
# total $
df_review = df_review.merge(pd.DataFrame(df_daily_wrap[['city', 'total_from_cat', 'total_diff']].groupby('city').sum().reset_index()), how='left', on='city')
# total tickets
df_review = df_review.merge(pd.DataFrame(df_daily_wrap[['city', 'total tix']].groupby('city').sum().reset_index()), how='left', on='city')
# average ticket price
df_review['avg_ticket_price'] = round(df_review['total_from_cat'] / df_review['total tix'], 2)

# ----- First day sales -----
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

# create one data frame of all first day sales
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

print('Summary of first day sales:')
print(df_first_day.groupby(['city'])['Total Gross', 'BO Gross'].sum())

