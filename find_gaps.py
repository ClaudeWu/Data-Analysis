# load modules
import pandas as pd
import numpy as np
import time
import openpyxl

# log start time
start_time = time.time()

# load input data set (Python pickle file)
df = pd.read_pickle(r'/Users/piyo/Desktop/px.xz') # replace <path> with proper file path

# USER CODE
# set time format
df['dt'] = pd.to_datetime(df['dt'], format='%Y-%m-%d')

# create two columns to store start & end date of each stock
# store dt in start & store each second date in end with group by bbgid
df = df.assign(start = df.dt, end = df.groupby('bbgid')['dt'].shift(-1))

# set dt as index
df.set_index('dt', inplace=True)

# group by bbgid with resample date in unit 1 day & insert missing date
# reset index
df = df.groupby('bbgid', as_index=False).resample(rule='1D').ffill().reset_index()

# create length column and insert gaps
df['length'] = np.where((df['dt'] != df['start']), (df['end']-df['start']-pd.to_timedelta('1 day')).dt.days, np.nan)

# drop NaN value
df = df.dropna()

# convert length to int
df['length'] = df.length.astype(int)

# find max value and group by bbgid
df = df.loc[df.groupby(['bbgid'])['length'].idxmax()]

# drop dt & level_0 columns
df = df.drop(columns = ['dt', 'level_0'])
df = df.drop_duplicates()

# reorder attributes
df = df[['start', 'end', 'length', 'bbgid']]

# sorting
df = df.sort_values(by='length', ascending=False)

# reset index
df = df.reset_index()

# remove index column
df = df.drop(columns = ['index'])

# df['Total'] = df.groupby(['bbgid'])['length'].transform('sum')
# create Total column & insert sum of length of each bbgid
# df = df.assign(Total = df.groupby(['bbgid'])['length'].transform('sum'))

# convert Total's  data to int
# df = df.astype({'Total': 'int32'})

# reorder date with bbgid
# df2 = df.groupby('bbgid').agg({'start': ['min'], 'end': ['max'], 'Total': ['max']})

# rename columns
# df2.columns=['start', 'end', 'length']

# reset index
# df2 = df.reset_index()

# relocate column bbgid
# df2 = df2[['start', 'end', 'length', 'bbgid']]

# sort length columns with ascending
# df2 = df2.sort_values(by='length', ascending=False)

# reset index
# df2 = df2.reset_index()

# remove index column
# df2 = df2.drop(columns = ['index'])

# export result to Excel
df.iloc[0:1000].to_excel(r'/Users/piyo/Desktop/px_stats.xlsx') # replace <path> with proper file path

# show execution time
print("--- %s seconds ---" % (time.time() - start_time))
