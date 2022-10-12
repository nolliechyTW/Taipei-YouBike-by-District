import requests
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import xlwings as xw

res = requests.get ("https://tcgbusfs.blob.core.windows.net/dotapp/youbike/v2/youbike_immediate.json")
json = res.json()
print (json)
# observe data are stored in Dictionary; every elements are composed of key:value

# save data in list instead
YouBikeData = list(json)

# use pandas to create the table
df = pd.DataFrame(YouBikeData,columns=['sarea','sna','sbi','bemp','tot'])
df.columns = ["District","Stop Name","Avalible Amount to Rent","Avalible Amount to Return","Total Amount"]

# try to use groupby but fail; check df info to see why
df.info()

# need to transform object to int, see: https://stackoverflow.com/questions/39173813/pandas-convert-dtype-object-to-int
df["Avalible Amount to Rent"] = pd.to_numeric(df["Avalible Amount to Rent"])
df["Avalible Amount to Return"] = pd.to_numeric(df["Avalible Amount to Return"])
df["Total Amount"] = pd.to_numeric(df["Total Amount"])

# check df info again 
df.info()

# add up the number of YouBikes in each district
df2 = df.groupby("District").sum()

# calculate the YouBike ratio of each district to the first decimal place
df2["Total Amount Ratio%"] = (df2 ["Total Amount"] *100 / df2 ["Total Amount"].sum()).round(1)
df2["Avalible Amount to Rent Ratio%"] = (df2 ["Avalible Amount to Rent"] *100 / df2 ["Total Amount"].sum()).round(1)
df2 

# order by Avalible Amount to Rent Ratio%, descending
df3 = df2.sort_values("Avalible Amount to Rent Ratio%", ascending=False)
df3

# draw the chart using matlotlib and FontProperties
# choose a style
plt.style.use("ggplot")

# choose a font
plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]

# choose a size of the chart
plt.rcParams["figure.figsize"] = (20,6)

plt.plot(df3.index , df3["Total Amount Ratio%"],label="Total Amount Ratio%",linewidth=3)
plt.plot(df3.index, df3["Avalible Amount to Rent"],label="Avalible Amount to Rent Ratio%",linewidth=3)
plt.ylabel("Ratio%",fontsize=15)
plt.xlabel("District",fontsize=15)
plt.title("The proportion of Youbike in each district of Taipei City",fontsize=15)

plt.legend(fontsize=15)
plt.show()

# save the chart in Excel
# create a new Excel workbook
wb = xw.Book()

# choose a worksheet
sheet = wb.sheets[0]

# write the DataFrame to a range starting at A1 (upper left corner)
sheet.range("A1").value = df3

# set a name to the Excel workbook
sheet.name = "The proportion of Youbike in each district of Taipei City"

# save the file
wb.save("The proportion of Youbike in each district of Taipei City_xw.xlsx")
