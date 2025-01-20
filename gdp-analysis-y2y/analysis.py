import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import statsmodels.api as sm

FIRST_ROW_INDEX = 6
FIRST_COL_INDEX = 3

COL_NAME_ROW_INDEX = 6
COL_NAME_COL_INDEX = 1

POP_DATA_ROW_END = 95

gdp_data_path = 'data/gdp-data-1929-to-2023.xlsx'
population_data_path = 'data/us-population-data-1929-to-2023.xlsx'

gdp_data_save_path = 'data/gdp-nominal-data-clean.xlsx'
gdp_real_data_save_path = 'data/gdp-real-data-clean.xlsx'
population_data_save_path = 'data/population-data-clean.xlsx'
merged_data_save_path = 'data/all-data-clean.xlsx'
growth_data_save_path = 'data/growth-rate-data-clean.xlsx'

# Load the Excel file data
gdp_data_xls = pd.ExcelFile(gdp_data_path)
population_data_xls = pd.ExcelFile(population_data_path)

# Load the sheets into pandas dataframes
df_gdp = pd.read_excel(gdp_data_xls, sheet_name='T10105-A')
df_real_gdp = pd.read_excel(gdp_data_xls, sheet_name='T10106-A')
df_pop = pd.read_excel(population_data_path)


### PROCESS THE DATA ###

# Perform data processing and cleaning
def process_gdp_data(df):
    df.iloc[COL_NAME_ROW_INDEX, COL_NAME_COL_INDEX] = 'Year'  # Set the first cell to 'Year'
    df_clean = df.iloc[FIRST_ROW_INDEX:, FIRST_COL_INDEX:]  # Skip the first 8 rows and the first 3 columns without data
    df_clean = df_clean.transpose()
    df_clean.columns = df.iloc[COL_NAME_ROW_INDEX:, COL_NAME_COL_INDEX]  # Set the column of series
    df_clean.columns = df_clean.columns.map(lambda x : str(x).strip())
    df_clean['Year'] = df_clean['Year'].astype(int)  # Convert the 'Year' column from float to integer
    return df_clean
def process_pop_data(df):
    df_clean = df.iloc[0:POP_DATA_ROW_END, :]
    df_clean.loc[:, 'Population'] = df_clean['Population'].apply(lambda pop_str : float(str(pop_str).replace("million", "").strip()))
    return df_clean
df_gdp = process_gdp_data(df_gdp) # Note that values are in millions of usd
df_real_gdp = process_gdp_data(df_real_gdp) # Note that values are in millions of usd
df_pop = process_pop_data(df_pop) # Note that population in is millions of people

# Save the transformed data to new Excel files
df_gdp.to_excel(gdp_data_save_path, index=False)
df_real_gdp.to_excel(gdp_real_data_save_path, index=False)
df_pop.to_excel(population_data_save_path, index=False)

# Merge the dataframes based on the 'Year' column
df = pd.merge(df_gdp, df_real_gdp, on='Year', suffixes=(' nominal', ' real'))
df = pd.merge(df_pop, df, on='Year')

# Save the merged data to a new Excel file
df.to_excel(merged_data_save_path, index=False)


### PLOT THE DATA ### 

# Create a PDF to save the plots
with PdfPages('economic_indicator_graphs.pdf') as pdf:
    # Graph 1: Real and Nominal GDP/capita from 1929 to 2023
    df['GDP per capita nominal'] = df['Gross domestic product nominal'] / df['Population']
    df['GDP per capita real'] = df['Gross domestic product real'] /  df['Population']
    plt.figure(figsize=(10, 6))
    plt.plot(df['Year'], df['GDP per capita nominal'], label='Nominal GDP per Capita', color='blue')
    plt.plot(df['Year'], df['GDP per capita real'], label='Real GDP per Capita', color='green')
    plt.title('Graph 1. Real and Nominal GDP per Capita (1929 to 2023)')
    plt.xlabel('Year')
    plt.ylabel('GDP per Capita (USD)')
    plt.legend()
    plt.grid(True)
    pdf.savefig()

    # Graph 2. Real and Nominal Personal Consumption Expenditures/capita (1929 to 2023) 
    df['PCE per capita nominal'] = df['Personal consumption expenditures nominal'] / df['Population']
    df['PCE per capita real'] = df['Personal consumption expenditures real'] / df['Population']
    plt.figure(figsize=(10, 6))
    plt.plot(df['Year'], df['PCE per capita nominal'], label='Nominal PCE per Capita', color='blue')
    plt.plot(df['Year'], df['PCE per capita real'], label='Real PCE per Capita', color='green')
    plt.title('Graph 2. Real and Nominal Personal Consumption Expenditures per Capita (1929 to 2023)')
    plt.xlabel('Year')
    plt.ylabel('PCE per Capita (USD)')
    plt.legend()
    plt.grid(True)
    pdf.savefig()

    # Graph 3. Real and Nominal Gross Private Domestic Investment/capita (1929 to 2023) 
    df['GPD investment per capita nominal'] = df['Gross private domestic investment nominal'] / df['Population']
    df['GPD investment per capita real'] = df['Gross private domestic investment real'] / df['Population']
    plt.figure(figsize=(10, 6))
    plt.plot(df['Year'], df['GPD investment per capita nominal'], label='Nominal GPD Investment per Capita', color='blue')
    plt.plot(df['Year'], df['GPD investment per capita real'], label='Real GPD Investment per Capita', color='green')
    plt.title('Graph 3. Real and Nominal Gross Private Domestic Investment per Capita (1929 to 2023)')
    plt.xlabel('Year')
    plt.ylabel('GPD Investment per Capita (USD)')
    plt.legend()
    plt.grid(True)
    pdf.savefig()

    # Graph 4. Real and Nominal (Exports – Imports)/capita (1929 to 2023)
    df['net exports per capita nominal'] = (df['Exports nominal'] - df['Imports nominal'])/ df['Population']
    df['net exports per capita real'] = (df['Exports real'] - df['Imports real'])/ df['Population']
    plt.figure(figsize=(10, 6))
    plt.plot(df['Year'], df['net exports per capita nominal'], label='Nominal Net Exports Investment per Capita', color='blue')
    plt.plot(df['Year'], df['net exports per capita real'], label='Real Net Exports per Capita', color='green')
    plt.title('Graph 4. Real and Nominal Gross (Exports - Imports) per Capita (1929 to 2023)')
    plt.xlabel('Year')
    plt.ylabel('Exports - Imports per Capita (USD)')
    plt.legend()
    plt.grid(True)
    pdf.savefig()

    # Graph 5. Real and Nominal Government Spending/capita (1929 to 2023) 
    df['govt spending per capita nominal'] = df['Government consumption expenditures and gross investment nominal'] / df['Population']
    df['govt spending per capita real'] = df['Government consumption expenditures and gross investment real'] / df['Population']
    plt.figure(figsize=(10, 6))
    plt.plot(df['Year'], df['govt spending per capita nominal'], label='Nominal Government Spending per Capita', color='blue')
    plt.plot(df['Year'], df['govt spending per capita real'], label='Real Government Spending per Capita', color='green')
    plt.title('Graph 5. Real and Nominal Government Spending per Capita (1929 to 2023)')
    plt.xlabel('Year')
    plt.ylabel('Government Spending per Capita (USD)')
    plt.legend()
    plt.grid(True)
    pdf.savefig()


### CALCULATE GROWTH RATES ###

# Create a new DataFrame for just the growth rates initialized with only the 'Year' column
growth_df = pd.DataFrame(df['Year']) 

# Calculate the growth rates of GDP per capita and components
growth_df['GDP per capita nominal growth'] = df['GDP per capita nominal'].pct_change() * 100
growth_df['GDP per capita real growth'] = df['GDP per capita real'].pct_change() * 100

growth_df['PCE per capita nominal growth'] = df['PCE per capita nominal'].pct_change() * 100
growth_df['PCE per capita real growth'] = df['PCE per capita real'].pct_change() * 100

growth_df['GPD investment per capita nominal growth'] = df['GPD investment per capita nominal'].pct_change() * 100
growth_df['GPD investment per capita real growth'] = df['GPD investment per capita real'].pct_change() * 100

growth_df['net exports per capita nominal growth'] = df['net exports per capita nominal'].pct_change() * 100
growth_df['net exports per capita real growth'] = df['net exports per capita real'].pct_change() * 100

growth_df['govt spending per capita nominal growth'] = df['govt spending per capita nominal'].pct_change() * 100
growth_df['govt spending per capita real growth'] = df['govt spending per capita real'].pct_change() * 100

growth_df = growth_df.dropna() # Drop the first row after computing the growth rates as it will have NaN values

# Save the merged data to a new Excel file
growth_df.to_excel(growth_data_save_path, index=False)


### PERFORM REGRESSION ###

def regress_over_period(start_year, end_year, df):
    # Filter the DataFrame for the given period
    period_df = df[(df['Year'] >= start_year) & (df['Year'] <= end_year)]

    X_nominal = period_df[['PCE per capita nominal growth',
        'GPD investment per capita nominal growth',
        'net exports per capita nominal growth',
        'govt spending per capita nominal growth']]
    X_nominal = sm.add_constant(X_nominal) # Add a constant to the independent variables
    Y_nominal = period_df['GDP per capita nominal growth']

    X_real = period_df[['PCE per capita real growth',
        'GPD investment per capita real growth',
        'net exports per capita real growth',
        'govt spending per capita real growth']]
    X_real = sm.add_constant(X_real) # Add a constant to the independent variables
    Y_real = period_df['GDP per capita real growth']

    # Perform OLS regression on nominal
    model_nominal = sm.OLS(Y_nominal, X_nominal).fit()
    print(model_nominal.summary())
    coefficients_nominal = {
        'Component': ['Constant nominal', 'PCE per capita nominal growth',
            'GPD investment per capita nominal growth',
            'net exports per capita nominal growth',
            'govt spending per capita nominal growth'],
        'Coefficient': model_nominal.params[0:].values
    }

    # Convert the nominal dictionary into a DataFrame for better display
    table = pd.DataFrame(coefficients_nominal)

    # Perform OLS regression on real
    model_real = sm.OLS(Y_real, X_real).fit()
    print(model_real.summary())
    coefficients_real = {
        'Component': ['Constant real', 'PCE per capita real growth',
            'GPD investment per capita real growth',
            'net exports per capita real growth',
            'govt spending per capita real growth'],
        'Coefficient': model_real.params[0:].values
    }

    # Add real dictionary to the table for better display
    table = pd.concat([table, pd.DataFrame(coefficients_real)], ignore_index=True)
    
    # Display the coefficients table
    print(f"Regression Coefficients for the period {start_year} to {end_year}:")
    print(table)

# Perform regression for the entire period
regress_over_period(1929, 2023, growth_df)

# Perform regression over WWII period
regress_over_period(1940, 1947, growth_df)

# Perform regression over the Economic Crisis of 2008 period
regress_over_period(2004, 2012, growth_df)

# Perform regression for the COVID-19 Pandemic period
regress_over_period(2019, 2023, growth_df)