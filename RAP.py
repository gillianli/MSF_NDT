import pandas as pd
#定义函数
def list_devices(row):
    if row.StartDevice not in device_list:
        device_list.append(row.StartDevice)
        location_list.append(row.StartDeviceLocation)
    if row.EndDevice not in device_list:
        device_list.append(row.EndDevice)
        location_list.append(row.EndDeviceLocation)

def uhigh(row):
    uhigh_dic = {'01C0': 'U42', '01M0': 'U46', '02M0': 'U45', '11T1': 'U17', '12T1': 'U16', '01RM1':'U35',\
                 '13T1': 'U15', '14T1': 'U14', '15T1': 'U13', '16T1': 'U12', '17T1': 'U11', '18T1': 'U10', 'T0': 'U26','P1': 'BV37'}

    if row.Name.split('-')[-1] in uhigh_dic.keys():
        suggestion_uhigh = uhigh_dic[row.Name.split('-')[-1]]
    else:
        suggestion_uhigh = uhigh_dic[row.Name.split('_')[-1][-2:]]
    uhigh = input(f'Please provide the u-high of {row.Name} in {row.Rack}: (Suggestion is {suggestion_uhigh})')
    if uhigh == '':
        row.Rack = row.Rack.split('-')[1] + '.' + suggestion_uhigh
    else:
        row.Rack = row.Rack.split('-')[1] + '.' + uhigh
def enter_uhigh(row):
    uhigh = input(f'Please provide the U-high of {row.Name} in Rack {row.Rack}:(exp: U18 or BV37)')
    row.Rack = row.Rack.split('-')[1] + '.' + uhigh

def order_column(df):
    order = ['StartDevice','StartPort','StartDeviceLocation','RU_Start','RU_End','Length','EndDevice','EndPort',\
             'EndDeviceLocation','LinkType','Speed']
    return df[order]

def order_ndt(df):
    order = ['StartDevice','StartPort','StartDeviceLocation','RU_Start','RU_End','EndDevice','EndPort',\
             'EndDeviceLocation','LinkType','Speed']
    return df[order]


#正文开始，读取文件
filename = input('Please enter the NDT file`s name (NO Extensions):')
ndt = pd.read_excel(f'{filename}.xlsx')
if ndt.columns[0] == "#Fields:StartDevice":
    ndt.rename(columns={'#Fields:StartDevice': 'StartDevice'}, inplace=True)

# 初始数据
device_list = []
location_list = []
ndt.Speed = ndt.Speed.fillna(0).astype(int)
ndt.StartPort = ndt.StartPort.str.replace('Ethernet', 'Eth')
ndt.StartPort = ndt.StartPort.str.replace('Management', 'mgmt')
ndt.EndPort = ndt.EndPort.str.replace('Ethernet', 'Eth')
ndt.EndPort = ndt.EndPort.str.replace('Management', 'mgmt')

# 构建Device表
# ndt.apply(list_devices, axis=1)
# device = pd.DataFrame({'Name': device_list, 'Rack': location_list})
# device.apply(enter_uhigh, axis=1)
ndt.apply(list_devices, axis=1)
device = pd.DataFrame({'Name': device_list, 'Rack': location_list})
print('Press Enter if suggestion is correct or provide the u-high manually. (Example: U26 or BV37)')
device.apply(uhigh, axis=1)

# 构建标签列
ndt = ndt.merge(device, how='left', left_on='StartDevice', right_on='Name')
ndt = ndt.merge(device, how='left', left_on='EndDevice', right_on='Name')
ndt['RU_Start'] = ndt['Rack_x'] + '.' + ndt['StartPort']
ndt['RU_End'] = ndt['Rack_y'] + '.' + ndt['EndPort']
ndt.drop(columns=['Rack_x','Rack_y','Name_x','Name_y'], inplace=True)

#构建AOC表
print('*'*20)
aoc = ndt[(ndt.LinkType == 'Data') & (ndt.Speed != 1000)].reset_index(drop=True)#Pandas用 & | ~表示与，或，非。
d1 = aoc.drop_duplicates(subset='StartDeviceLocation', keep='last').StartDeviceLocation
l1 = [0]
for i in d1.index:
    l1.append(i)
    length = input(f'Please enter the AOC cable length between {aoc.StartDeviceLocation.at[i]} and\
 {aoc.EndDeviceLocation.at[i]} (Numbers Only):')
    aoc.loc[l1[0]:l1[1],'Length'] = 'AOC ' + str(int(aoc.Speed.at[i]/1000)) + 'G ' + length + 'M'
    l1.pop(0)
    l1[0] += 1
aoc_count = aoc.Length.value_counts().to_frame()

# #构建Copper表
copper = ndt[(ndt.Speed.isin([9600,115200,1000,0])) & (ndt.StartPort != 'power0')].sort_values(by=['StartDeviceLocation'],ascending=True)\
    .reset_index(drop=True)
d1 = copper.drop_duplicates(subset='StartDeviceLocation', keep='last').StartDeviceLocation
l1 = [0]
for i in d1.index:
    l1.append(i)
    if copper.StartDeviceLocation.at[i] == copper.EndDeviceLocation.at[i]:
        l1.pop(0)
        l1[0] += 1
    else:
        length = input(f'Please enter the Copper cable length between {copper.StartDeviceLocation.at[i]} and\
 {copper.EndDeviceLocation.at[i]} (Numbers Only):')
        copper.loc[l1[0]:l1[1], 'Length'] = 'CAT6 ' + length + 'FT'
        l1.pop(0)
        l1[0] += 1
copper_count = copper.Length.value_counts().to_frame()

#构建Power表
power = ndt[(ndt.Speed == 0) & (ndt.StartPort == 'power0')].reset_index(drop=True)

#构建LableMaster
lablemaster = pd.concat([aoc[['RU_Start','RU_End','Length']],copper[copper.Length.notnull()]\
    [['RU_Start','RU_End','Length']]], ignore_index=True)

#构建CableCount
cable_count = pd.concat([aoc_count,copper_count])


#写入excel
with pd.ExcelWriter(f'Done_{filename}.xlsx') as writer:
    order_ndt(ndt).to_excel(writer, sheet_name='NDT', index=False)
    device.to_excel(writer, sheet_name='Device', index=False)
    order_column(aoc).to_excel(writer, sheet_name='AOC', index=False)
    order_column(copper).to_excel(writer, sheet_name='Copper', index=False)
    lablemaster.to_excel(writer, sheet_name='LableMaster', index=False)
    cable_count.to_excel(writer, sheet_name='CableCount')




