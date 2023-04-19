import pandas as pd
import streamlit as st 

#%%  SELECT FOLDER 
st.markdown("""
            ### Select folder
            """)
            
folder = st.selectbox("Choose your folder", ["04_Afgeronde projecten","COTU","Documentatie","Normen"])

# Path naar TXT
# Update: voeg papier formaat toe aan de dictionary + doorzoek een map waar alle txt files in staan.
# Inladen txt file in dictionary:
    
# Geef het pad op waar het txt-bestand zich bevindt
txt_path = r'Y:\Algemeen\Software\Python\Python middag\Apps\Zoekmachine project referentie\03 txt\{}.txt'.format(folder)

# Lees het txt-bestand in als een pandas DataFrame
df = pd.read_csv(txt_path, sep='\t')

# Maak een lege dictionary om de gegevens op te slaan
pdf_data = {}

# Loop door elke rij in de DataFrame en voeg de gegevens toe aan de dictionary
for index, row in df.iterrows():
    pdf_data[row['filename']] = {'text': str(row['text']), 'path': row['path'], 'paper_size': row['paper_size']}


#%% SEARCH CRITERIA
st.markdown("""
            ### Input search criteria
            """)
search_terms = st.text_input("Input search criteria, separated by commas", key="Input2")
search_terms_list = search_terms.split(',')

output_path = r'Y:\Algemeen\Software\Python\Python middag\Apps\Zoekmachine project referentie\04 Excel\List_{}.xlsx'.format(''.join(search_terms))
 # pas het pad en de bestandsnaam aan
 
#%% SEARCH 

# Create a dictionary to store results for each search term
all_results = {}

# Loop over each search term
for search_term in search_terms_list:
    results = {} # een lege dictionary om de resultaten bij te houden

    # loop over elk item in de pdf_data dictionary
    for filename, data in pdf_data.items():
        # zoek naar de zoekterm in de 'text'-waarde van de huidige dictionary
        if search_term in data['text']:
            count = data['text'].count(search_term) # tel het aantal keren dat de zoekterm voorkomt
            if filename in results:
                results[filename][search_term] = count
            else:
                results[filename] = {search_term: count, 'path': data['path'], 'filename': filename}

    # Voeg het aantal resultaten toe aan de dictionary met alle resultaten
    all_results[search_term] = {
        'count': len(results),
        'results': results
    }

#%% WRITE TO EXCEL

# Write all the results to the Excel file
writer = pd.ExcelWriter(output_path, engine='xlsxwriter')


# Combine all the search results into a single dataframe
combined_results = pd.DataFrame(columns=['filename','paper_size'] + search_terms_list + ['path'])
for search_term, data in all_results.items():
    sorted_results = data['results']
    for filename, result in sorted_results.items():
        row = {'filename': filename, 'path': result['path'], 'paper_size': pdf_data[filename]['paper_size']} 
        for term in search_terms_list:
            if term in result:
                row[term] = result[term]
            else:
                row[term] = 0
        combined_results = combined_results.append(row, ignore_index=True)

# Define a function to format a path as a hyperlink
def format_path_as_hyperlink(path):
    return '=HYPERLINK("{}")'.format(path)

#%% GROUP

# Group the results by filename and sum the counts
grouped_results = combined_results.groupby('filename').sum()

# Add a 'path' column to the grouped results
grouped_results['path'] = combined_results.groupby('filename')['path'].apply(lambda x: format_path_as_hyperlink(x.iloc[0]))

# Reset the index of the grouped results
grouped_results = grouped_results.reset_index()

# Replace any NaN values in the grouped results with 0
grouped_results = grouped_results.fillna(0)

# Sort the grouped results by total counts of search results
grouped_results['Total Counts'] = grouped_results[search_terms_list].sum(axis=1)
grouped_results = grouped_results.sort_values('Total Counts', ascending=False)

# Remove the Total Counts column
grouped_results = grouped_results.drop(columns=['Total Counts'])

# Write the grouped results to the Excel file
grouped_results.to_excel(writer, sheet_name='Grouped', index=False)


# Sort the combined results by filename
combined_results = combined_results.sort_values('filename')

# Sort the combined results by total counts of search results
combined_results = combined_results.astype({'filename': 'str'})
combined_results['Total Counts'] = combined_results[search_terms_list].sum(axis=1)
combined_results = combined_results.sort_values('Total Counts', ascending=False)

#%% OUTPUT

# Print the size of the sorted_results dictionary
st.write("Number of results:", len(sorted_results))

writer.save()

st.write(output_path)
