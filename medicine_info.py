# %%
import pandas as pd
from connections import engine
# %%
sql = '''select mi.id, mi.full_name, mi.alias, mt.name, mt.sub_category,
mi.specification
from sanyi_care.c_medicine_info_v2 as mi
left join sanyi_care.c_medicine_types as mt
on mi.category_id = mt.id
'''
with engine.connect() as con:
    med = pd.read_sql(sql, con)
med
# %%
med.rename({'full_name': '药品名', 'alias': '商品名',
            'name': '药品类别', 'sub_category': '药品子类',
            'specification': '规格'}, axis=1).to_csv('原药品库.csv', index=False)
# %%
