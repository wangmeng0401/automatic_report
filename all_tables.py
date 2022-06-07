import pandas as pd

df = pd.read_excel('.\data_5\月新增注册.xlsx', sheet_name='表1')
df_2 = pd.read_excel('.\data_5\月新增注册.xlsx', sheet_name='表2')
df_3 = pd.read_excel('.\data_5\投融资.xlsx', sheet_name='表4')
df_4 = pd.read_excel('.\data_5\投融资.xlsx', sheet_name='表5')
df_5 = pd.read_excel('.\data_5\创新发展.xlsx', sheet_name='表6')
df_6 = pd.read_excel('.\data_5\政策.xlsx')

for i in df['功能区'].unique():
    condition = (df['功能区'] == i)
    df[condition][['公司名称', '企业法人', '注册资本（万元）', '经营范围']].to_excel(
        '.\data_5\当月新增-表格\新增注册-' + i + '.xlsx', sheet_name=i, index=False)

for i in df_2['功能区'].unique():
    condition = (df_2['功能区'] == i)
    df_2[condition][['企业名称', '法人名称', '注册资本（万元）', '经营范围']].to_excel(
        '.\data_5\当月新增-表格\新增注销-' + i + '.xlsx', sheet_name=i, index=False)

for i in df_3['功能区'].unique():
    condition = (df_3['功能区'] == i)
    df_3[condition][['日期', '企业', '法人名称', '被投资企业', '投资金额（万元）', '经营范围']].to_excel(
        '.\data_5\投融资-表格\对外投资-' + i + '.xlsx', sheet_name=i, index=False)

for i in df_4['功能区'].unique():
    condition = (df_4['功能区'] == i)
    df_4[condition][['日期', '企业名称', '法人名称', '投资人', '融资金额（万元）', '融资轮次']].to_excel(
        '.\data_5\投融资-表格\融资情况-' + i + '.xlsx', sheet_name=i, index=False)

for i in df_5['功能区'].unique():
    condition = (df_5['功能区'] == i)
    df_5[condition][['专利名称', '申请机构/个人', '所属行业', '申请时间']].to_excel(
        '.\data_5\创新发展-表格\创新发展-' + i + '.xlsx', sheet_name=i, index=False)

for i in ['城南地区', '回天地区', '怀柔科学城', '商务中心区', '临空经济区', '金融街', '通州高端商务服务区', '北京经济技术开发区', '中关村科技园区海淀园', '新首钢高端产业综合服务区',
          '奥林匹克中心区', '丽泽金融商务区']:
    condition = (df_6['area'] == i)
    df_6[condition][['政策标题', '政策内容', '发布时间']].to_excel(
        '.\data_5\政策-表格\政策信息-' + i + '.xlsx', sheet_name=i, index=False)
