# def read_FtMyers(fd_f, fn_f, sheet_f):
#     df_f = pd.read_excel(io=os.path.join(fd_f, fn_f), sheet_name=sheet_f, header=3)
#
#     t_colnames = list(df_f.columns[2:])
#     t_colnames.insert(0, 'notes')
#     t_colnames.insert(0, 'Date')
#     df_f.columns = t_colnames
#
#     df_f.Date = pd.to_datetime(df_f.Date, errors='coerce').dt.date
#
#     df_f.drop(df_f.loc[df_f.Date.isnull()].index, inplace=True)
#     df_f.drop(columns=['Total', 'CUMULATIVE TOTAL', 'Average Ticket Price'], inplace=True)
#
#     return df_f
#
#
# def read_GrandRapids(fd_f, fn_f, sheet_f):
#     df_f = pd.read_excel(io=os.path.join(fd_f, fn_f), sheet_name=sheet_f, header=3)
#
#     t_colnames = list(df_f.columns[2:])
#     t_colnames.insert(0, 'notes')
#     t_colnames.insert(0, 'Date')
#     df_f.columns = t_colnames
#
#     df_f.Date = pd.to_datetime(df_f.Date, errors='coerce').dt.date
#
#     df_f.drop(df_f.loc[df_f.Date.isnull()].index, inplace=True)
#     df_f.drop(columns=['Total', 'CUMULATIVE TOTAL', 'Average Ticket Price'], inplace=True)
#
#     return df_f



# All sales
# df_master = read_MASTER(fd_f=fd, fn_f='__MASTER Sales Summary and Daily Wraps Doc.xlsx', sheet_f='Daily Wrap')
# df_MfMyers = read_type1(fd, fn_f='HAMILTON Ft. Myers Sales Summary.xlsx', sheet_f='Ft. Myers - Daily Wrap', header_row_ind_f=3)
# df_GrandRapids = read_type1(fd, fn_f='HAMILTON Grand Rapids Sales Summary.xlsx', sheet_f='Daily Wraps', header_row_ind_f=3)
# df_Indianapolis = read_type1(fd, fn_f='HAMILTON Indianapolis 2019 Sales Summary.xlsx', sheet_f='IND Daily Wraps', header_row_ind_f=4)
# df_LosAngeles = read_type1(fd, fn_f='Hamilton LA Sales Summary.xlsx', sheet_f='LA - Daily Sales', header_row_ind_f=4)
# df_Miami = read_type1(fd, fn_f='Hamilton Miami Sales Summary.xlsx', sheet_f='Daily Wraps - FINAL', header_row_ind_f=4)
# df_Naples = read_type1(fd, fn_f='HAMILTON Naples Summary.xlsx', sheet_f='Naples- Daily Wrap', header_row_ind_f=3)
# df_Nashville = read_type1(fd, fn_f='Hamilton Nashville Sales Summary.xlsx', sheet_f='Nashville Daily Wraps', header_row_ind_f=3)
# df_Norfolk = read_type1(fd, fn_f='HAMILTON Norfolk Sales Summary.xlsx', sheet_f='Norfolk - Daily Wrap', header_row_ind_f=3)
# df_Philadelphia = read_type1(fd, fn_f='HAMILTON Philadelphia Sales Summary.xlsx', sheet_f='Daily Wrap', header_row_ind_f=2)
# df_Richmond = read_type1(fd, fn_f='HAMILTON Richmond Sales Summary.xlsx', sheet_f='Richmond - Daily Wrap', header_row_ind_f=3)
# df_Toronto = read_type1(fd, fn_f='Hamilton Toronto Sales Summary.xlsx', sheet_f='Daily Wraps', header_row_ind_f=5)
# df_WestPalmBeach = read_type1(fd, fn_f='Hamilton West Palm Beach Sales Summary.xlsx', sheet_f='Daily Wraps', header_row_ind_f=4)
