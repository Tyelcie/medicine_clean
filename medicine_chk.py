# %%
import pandas as pd
from connections import engine
# %%
med = pd.read_excel('药品名规格入库清单.xlsx')
med
# %%
med_to_find = med.query('alias != alias')[['id', '来源药品ID']]
med_to_find['id'] = med_to_find['id'].astype(int)
med_to_find['source_id'] = med_to_find['来源药品ID'].str.split('|')
med_to_find
# %%
med_to_find = med_to_find.explode('source_id')
med_to_find
# %%
sql = f'''select medicine_id, count(1) as cnts
from sanyi_care.c_prescription_tag as pre
where medicine_id in {tuple(med_to_find['source_id'].unique())}
group by medicine_id'''
with engine.connect() as con:
    pres = pd.read_sql(sql, con)
pres
# %%
pres['medicine_id'] = pres['medicine_id'].astype(int).astype(str)
med_to_find = med_to_find.join(pres.set_index('medicine_id'),
                               on='source_id', how='left')
med_to_find
# %%
med_to_find = med_to_find.groupby(['id'])['cnts'].sum().to_frame('pres_cnts')
med_to_find
# %%
med = med.join(med_to_find, on='id', how='left')
med['id']= med['id'].astype(int).astype(str)
# %%
med.to_excel('药品名规格入库清单.xlsx', index=False)
# %%
