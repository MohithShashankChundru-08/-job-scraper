import pandas as pd
data = [
    {'Title': 'Book1 ', 'Price': '20', 'Rating': 4},
    {'Title': 'Book1 ', 'Price': '20', 'Rating': 4},
    {'Title': None, 'Price': '30', 'Rating': 5},
    {'Title': 'Book3', 'Price': None, 'Rating': 3}
]
df = pd.DataFrame(data)
df = df.drop_duplicates()
df['Name'] = df['Name'].fillna('Unknown book')
df['Price'] = df['Price'].astype(float)
df['Price'] = df['Price'].fillna(int(df['Price'].mean()))
print(df)