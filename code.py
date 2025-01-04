import pandas as pd
from fuzzywuzzy import fuzz, process

# Load data into DataFrames
df_1527 = pd.read_excel(r'blacklistedCustomer.xlsx', sheet_name = 'Sheet1')
df_1_2_million = pd.read_excel(r'dataset.xlsx', sheet_name = 'Sheet4')



# Preprocess names for approximate matching
df_1527['Name Matching'] = df_1527['Name_NRC'].str.lower().str.replace(' ', '')
df_1_2_million['Name Matching'] = df_1_2_million['NameAndNRC'].str.lower().str.replace(' ', '')
#df_1527['Name Matching'] = df_1527['NAMEDOB']
#df_1_2_million['Name Matching'] = df_1_2_million['Name+DOB']

# Initialize an empty DataFrame to store the matches
matches = pd.DataFrame(columns=['Name_NRC','NameAndNRC', 'Score'])

# Perform approximate matching using fuzzywuzzy
for idx, row in df_1527.iterrows():
    name = row['Name Matching']
    match, score, _ = process.extractOne(name, df_1_2_million['Name Matching'])
    if score >= 80:  #adjust the threshold (0-100) for matching accuracy
        matched_row = df_1_2_million.loc[df_1_2_million['Name Matching'] == match]
        matches = matches._append({           
            'Name_1572': row['Name_NRC'],
            'CIF' : matched_row['CIF'].iloc[0],
            'Name_1_2_million': matched_row['NameAndNRC'].iloc[0],
            'Score': score
        }, ignore_index=True)

# Save the matches to a new CSV file
matches.to_csv('NRC_matched_data_sheet4.csv', index=False)
