import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


# load data
file='data_sheet.xlsx'
data= pd.ExcelFile(file)
print(data.sheet_names)

# create data frame from data
df_col = data.parse('RD0557_ARC6_404_R3C4@lam940nm')
df_diffuser = data.parse('Diffuser_(RD0479_Jul20_OB6@940')
df_fanout = data.parse('Fanout_(RD0556_ARC6_402)@940nm')
df_1M = data.parse('MetaLens_(RD0526)_1M_A(8)')



# fanout graphs
fanout_nonuniformity = df_fanout['Non_uniformity (%)']
fanout_eff = df_fanout['Efficiency (%)']

fig, ax = plt.subplots(1,2)
ax[0].boxplot(fanout_nonuniformity)
ax[1].boxplot(fanout_eff)


#diffuser 0_oder
diff_zero_order = df_diffuser['0_order efficiency(%)']
fig, ax = plt.subplots()
ax.boxplot(diff_zero_order)
ax.set_xlabel('diffuser')
ax.set_ylabel('0_order %')
ax.set_title('Statistical measurements')




# diffuser collimator
dif_col_EFL = df_col['EFL_dif.col (um)'].dropna()
dif_col_eff = df_col['collimation eff_dif.col (%)'].dropna()
fan_col_EFL = df_col['EFL_fan.col (um)'].dropna()
fan_col_eff = df_col['collimation eff_dif.col (%)'].dropna()


#plot the column of interest
fig, ax = plt.subplots(2, 1, sharex=True)
ax[0].boxplot([dif_col_EFL, fan_col_EFL])
ax[1].boxplot([dif_col_eff, fan_col_eff])

ax[1].set_xticklabels(['diffuser', 'fanout'])
ax[0].set_ylabel('EFL ($\mu$m)')
ax[1].set_ylabel('Efficiency (%)')


plt.show()