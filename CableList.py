import pandas as pd
from datetime import datetime

with open("config_cableist", 'r', encoding='utf8') as f:
    conf = f.readlines()
#more to do:
#1. auto duplicate XB terminals with :WH connection points
# VVVVV 3. write a log file to record what has been done to a the specific file

fpath = conf[0].lstrip('filename:').rstrip('\n')
df = pd.read_excel(fpath)

temp = df.copy(deep=True) 

############# 1. duplicate data, get source and target swapped for the second part of data  VVVVVVVVVVV
temp[['From','From where','To','To Where']] = df[['To','To Where','From','From where']]
df = pd.concat([df,temp])

df.sort_values(['From where','To Where'], ascending=True,inplace=True)

df.reset_index(inplace=True)

filt = df['To Where'] == df['From where']

#    df2.drop(df2[filt1].index, inplace=True)

df_internal = df.loc[filt]
df_internal.drop(df_internal[df_internal['Cable Name'].duplicated()].index,inplace=True)
temp_internal = df_internal['From where'].value_counts()._info_axis.sort_values()

df_external = df.loc[~filt]
temp_external = df_external['From where'].value_counts()._info_axis.sort_values()
    

with open((fpath.rstrip('.xls') + ".txt"), 'w', encoding='utf8') as f:
    for i in temp_internal:
        f.write("=====" + i + " internal cables=====\n")
        df_string = df_internal[df_internal['From where'] ==i]['Cable Name'].to_string(header = False, index = False).replace(' ','').replace("'",'')         #filter all the labels inside a location
        f.write(df_string + "\n")         #write to the file
        f.write(df_string + "\n")         #print one more time for internal cables 
    for i in temp_external:
        f.write("=====Cables out from "+ i +"=====\n")
        df_string = df_external[df_external['From where'] ==i]['Cable Name'].to_string(header = False, index = False).replace(' ','').replace("'",'')         #filter all the labels inside a location
        f.write(df_string + "\n")         #write to the file


print("Over!, time:",datetime.now().strftime("%H:%M:%S"))
