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
