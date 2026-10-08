# %%
import requests
import pandas as pd
import re
from lxml import html
from progressbar import ProgressBar, SimpleProgress, Percentage, Bar
import time
from connections import engine
# %%
tbl = pd.read_csv('药品名规格清洗清单.csv')
tbl = tbl.rename(columns={'药品ID': 'medicine_id',
                          '药品名': 'full_name',
                          '商品名': 'alias',
                        #   '匹配状态': 'match_status',
            '药品类别': 'category',
            # '药品子类': 'sub_category',
            '规格': 'specification'})
tbl
# %%
sql = '''select * from sanyi_care.c_medicine_info_v2'''
with engine.connect() as con:
    med = pd.read_sql(sql, con)
med = med.drop('pharma_dose', axis=1)
effect_note_dict = med.set_index('id')[['effect', 'note', 'specification', 'manufacturer']].to_dict()
effect_note_dict
# %%
med_new = med.query('id in @tbl["medicine_id"]')
med_new = med_new.join(tbl.set_index('medicine_id')[['匹配状态', '清洗备注', '来源药品ID']],
                       on='id', how='left')
med_new['effect_new'] = (med_new['来源药品ID'].str.split('|')
                         .map(lambda x: set([effect_note_dict['effect'][int(i)] for i in x])))
med_new['note_new'] = (med_new['来源药品ID'].str.split('|')
                         .map(lambda x: set([effect_note_dict['note'][int(i)] for i in x])))
med_new['specification_new'] = (med_new['来源药品ID'].str.split('|')
                         .map(lambda x: set([effect_note_dict['specification'][int(i)] for i in x])))
# %%
med_new.to_excel('药品名规格入库清单.xlsx', index=False)
# %%
med_new = pd.read_excel('药品名规格入库清单.xlsx')
med_new
# %%
med_new['manufacturer'] = med_new['id'].map(lambda x: effect_note_dict['manufacturer'][x])
med_new
# %%
med_new.to_excel('药品名规格入库清单.xlsx', index=False)
# %%
df = pd.read_excel('国家人口健康科学数据中心药品库.xlsx')
# %%
med_new_chk = med_new.join(df.set_index('产品名称')['detail_link'], on='full_name',
             how='left')
print(med_new_chk.info())
med_new_link = med_new_chk[['id', 'detail_link']].dropna()
med_new_link = med_new_link.reset_index(drop=True)
med_new_link
# %%
# 提取当前网页的药品详情，复用上面已经获取的响应
def get_medicine_info(rsp):

    rsp.raise_for_status()
    rsp.encoding = 'utf-8'
    detail_page = html.fromstring(rsp.text)
    detail_sections = detail_page.xpath(
        '//div[contains(concat(" ", normalize-space(@class), " "), " congczt ")]'
    )
    if not detail_sections:
        raise ValueError('未找到药品详情区域，请检查网页内容或页面结构是否变化。')

    # 保留空字段；只在字段名之后拆分，避免截断规格和日期中的冒号
    detail_dict = {}
    for paragraph in detail_sections[0].xpath('.//p'):
        paragraph_text = ''.join(paragraph.itertext()).strip()
        matched = re.match(r'^[〖【](.*?)[〗】]\s*[:：]\s*(.*)$', paragraph_text, re.S)
        if matched:
            field, value = matched.groups()
            detail_dict[field.strip()] = value.strip()

    if not detail_dict:
        raise ValueError('药品详情区域中未识别到字段，请检查字段格式是否变化。')

    detail_df = pd.DataFrame([detail_dict])
    detail_df['detail_link'] = txt.url
    return detail_df

# %%
med_chk_list = []
pbar = ProgressBar(
        maxval=med_new_chk['detail_link'].nunique(),
        widgets=['已完成 ', SimpleProgress(), ' 行 ', Percentage(), ' ', Bar()],
    ).start()
# 启动后设置刷新间隔，每完成一页刷新一次
pbar.update_interval = 1
pbar.next_update = 1
for i in pbar(med_new_chk['detail_link'].unique().tolist()):
    txt = requests.get(i)
    med_i = get_medicine_info(txt)
    med_chk_list.append(med_i)
    if len(med_chk_list) % 5 == 0:
        time.sleep(3)
# %%
med_info = pd.concat(med_chk_list)
med_info = med_info.reset_index(drop = True)
med_info
# %%
med_info['规格'] = med_info['规格'].fillna(med_info['规格（中文）'])
med_info[['detail_link', '规格']].isna().sum(axis = 1).mean()
# %%
med_new_chk = med_new_chk.join(med_info.set_index('detail_link')[['规格']],
                               on='detail_link', how='left')
med_new_chk.info()
# %%
(med_new_chk.query('alias != alias and detail_link == detail_link')
 [['id', 'alias', 'full_name', 'specification', '规格']])
# %%
med_new.loc[med_new['id']==653, 'alias'] = 'y'
med_new.loc[med_new['id']==66, 'alias'] = 'y'
med_new.loc[med_new['id']==188, 'alias'] = 'y'
med_new.loc[med_new['id']==203, 'alias'] = 'y'
med_new.loc[med_new['id']==263, 'alias'] = 'y'
med_new.loc[med_new['id']==370, 'alias'] = 'n'
med_new.to_excel('药品名规格入库清单.xlsx', index=False)
med_new = pd.read_excel('药品名规格入库清单.xlsx')
med_new_chk = med_new.join(df.set_index('产品名称')['detail_link'], on='full_name',
             how='left')
med_new_chk = med_new_chk.join(med_info.set_index('detail_link')[['规格']],
                               on='detail_link', how='left')
# %%
