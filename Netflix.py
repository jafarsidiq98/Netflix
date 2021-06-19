import pandas as pd

data = pd.read_csv("netflix_titles.csv")
data['date_added'] = data['date_added'].astype("datetime64")
data['type'] = data['type'].astype("category")

country = data['country'].str.split(", ", expand=True).melt()
country = pd.DataFrame(country['value'].value_counts()).reset_index().rename(columns = {'index': 'country'})

country['country'] = country['country'].str.replace('^.*Germany.*$','Germany').str.replace(
    '^.*Soviet Union.*$','Russia').str.replace(",", "")

data_country = country.groupby('country').value.sum().reset_index()

data_director = data['director'].str.split(", ", expand = True).melt()
data_director = pd.DataFrame(data_director['value'].value_counts()).reset_index().rename(columns = {'index': 'director'})

data_listed_in = data['listed_in'].str.split(", ", expand = True).melt()
data_listed_in = pd.DataFrame(data_listed_in['value'].value_counts()).reset_index().rename(columns = {'index': 'listed_in'})

data_cast = data['cast'].str.split(", ", expand = True).melt()
data_cast = pd.DataFrame(data_cast['value'].value_counts()).reset_index().rename(columns = {'index': 'cast'})

writer = pd.ExcelWriter("Netflix.xlsx", engine = "xlsxwriter")
data.to_excel(writer, sheet_name = 'data_netflix',index = False)
data_country.to_excel(writer, sheet_name = 'data_netflix_country',index = False)
data_director.to_excel(writer, sheet_name = 'data_netflix_director',index = False)
data_listed_in.to_excel(writer, sheet_name = 'data_netflix_listed_in',index = False)
data_cast.to_excel(writer, sheet_name = 'data_netflix_cast',index = False)
writer.save()
