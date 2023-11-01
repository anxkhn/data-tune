import streamlit as st
import pandas as pd
import io

# Streamlit UI
st.title("DataTune 🎚️")

# Function to download the cleaned DataFrame
def download(cleaned_file, file_name_with_extension):
    st.download_button(
        label="Download Cleaned File",
        data=cleaned_file,
        key="cleaned_file",
        file_name=file_name_with_extension,
    )

# Function to replace null values in the DataFrame
def replace(df):
    custom_value = st.text_input("Enter a value to replace nulls:", "NULL")
    df = df.fillna(custom_value)
    return df

# Function to remove duplicate rows in the DataFrame
def remove_duplicates(df):
    return df.drop_duplicates(keep='first')

# Function to remove rows with missing values
def remove_missing_values(df):
    return df.dropna()

# Function to convert selected columns to lowercase
def convert_to_lowercase(df, selected_columns):
    for column in selected_columns:
        if column in df.columns:
            if df[column].dtype == 'object':
                df[column] = df[column].str.lower()
            else:
                st.error(f"Column '{column}' is not of string data type. Select a string column to convert to lowercase.")
                return df
    return df

# Function to delete selected columns from the DataFrame
def delete_columns(df, columns_to_delete):
    df = df.drop(columns=columns_to_delete, errors='ignore')
    return df

# Function to sort a given column in ascending or descending order
def sort_column(df, column_name, ascending):
    if column_name in df.columns:
        # Convert all values in the column to strings and then sort
        df[column_name] = df[column_name].astype(str)
        df = df.sort_values(by=column_name, ascending=ascending)
    return df

# Function to capitalize the first letter of elements in the selected columns (alphabet characters only)
def capitalize_column_values(df, selected_columns):
    for selected_column in selected_columns:
        if selected_column in df.columns:
            if df[selected_column].dtype == 'object':
                df[selected_column] = df[selected_column].apply(lambda x: x.capitalize() if isinstance(x, str) and x and x[0].isalpha() else x)
            else:
                st.error(f"Column '{selected_column}' is not of string data type. Select a string column to capitalize.")
    return df

# Initialize a dictionary to store dataframes for each sheet
sheets_dataframes = {}

uploaded_file = st.file_uploader("Upload an Excel or CSV file", type=["xls", "xlsx", "csv"])

if uploaded_file is not None:
    # Read the uploaded file into a DataFrame
    try:
        if uploaded_file.name.endswith(('xls', 'xlsx')):
            # Load Excel file using openpyxl
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names

            for sheet_name in sheet_names:
                sheets_dataframes[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
        else:
            df = pd.read_csv(uploaded_file)
            sheets_dataframes["Sheet 1"] = df  # Store the CSV data in a sheet named "Sheet 1"
    except Exception as e:
        st.error(f"Error: {e}")
    
    st.success("File uploaded successfully!")

    # Select a sheet to work with
    selected_sheet = st.selectbox("Select a sheet to clean:", list(sheets_dataframes.keys()))

    # List of radio buttons for cleaning options with multiselect
    clean_options = st.multiselect("Select cleaning options:", ["Replace", "Remove Duplicate", "Remove Missing Values", "Convert to Lowercase", "Delete Columns", "Sort Column", "Capitalize Columns"])

    if "Replace" in clean_options:
        sheets_dataframes[selected_sheet] = replace(sheets_dataframes[selected_sheet])
        st.write("Replaced null values with custom value.")
    if "Remove Duplicate" in clean_options:
        sheets_dataframes[selected_sheet] = remove_duplicates(sheets_dataframes[selected_sheet])
        st.write("Removed duplicate rows.")
    if "Remove Missing Values" in clean_options:
        sheets_dataframes[selected_sheet] = remove_missing_values(sheets_dataframes[selected_sheet])
        st.write("Removed rows with missing values.")
    if "Convert to Lowercase" in clean_options:
        selected_columns = st.multiselect("Select columns to convert to lowercase:", sheets_dataframes[selected_sheet].columns)
        sheets_dataframes[selected_sheet] = convert_to_lowercase(sheets_dataframes[selected_sheet], selected_columns)
        st.write(f"Converted selected columns to lowercase: {', '.join(selected_columns)}")
    if "Delete Columns" in clean_options:
        columns_to_delete = st.multiselect("Select columns to delete:", sheets_dataframes[selected_sheet].columns)
        sheets_dataframes[selected_sheet] = delete_columns(sheets_dataframes[selected_sheet], columns_to_delete)
        st.write(f"Deleted selected columns: {', '.join(columns_to_delete)}")
    if "Sort Column" in clean_options:
        column_to_sort = st.selectbox("Select a column to sort:", sheets_dataframes[selected_sheet].columns)
        sort_order = st.radio("Select sorting order:", ["Ascending", "Descending"])
        ascending = sort_order == "Ascending"
        sheets_dataframes[selected_sheet] = sort_column(sheets_dataframes[selected_sheet], column_to_sort, ascending)
        st.write(f"Sorted column '{column_to_sort}' in {sort_order.lower()} order.")
    if "Capitalize Columns" in clean_options:
        selected_columns_to_capitalize = st.multiselect("Select columns to capitalize:", sheets_dataframes[selected_sheet].columns)
        sheets_dataframes[selected_sheet] = capitalize_column_values(sheets_dataframes[selected_sheet], selected_columns_to_capitalize)
        st.write(f"Capitalized selected columns: {', '.join(selected_columns_to_capitalize)}")

    # Allow users to download the cleaned file in the same format as the input file
    cleaned_file = io.BytesIO()
    
    if uploaded_file.name.endswith(('xls', 'xlsx')):
        with pd.ExcelWriter(cleaned_file, engine='openpyxl') as writer:
            for sheet_name, sheet_df in sheets_dataframes.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        file_extension = "xlsx"
    else:
        sheets_dataframes[selected_sheet].to_csv(cleaned_file, index=False)
        file_extension = "csv"
    
    # Get the name of the uploaded file without the extension
    input_file_name = uploaded_file.name.rsplit('.', 1)[0]

    download(cleaned_file, f"{input_file_name}_cleaned.{file_extension}")

    # If uploaded file is XLS or XLSX, add a button to download as CSV
    if uploaded_file.name.endswith(('xls', 'xlsx')):
        cleaned_xls_file = io.BytesIO()
        with pd.ExcelWriter(cleaned_xls_file, engine='openpyxl') as writer:
            for sheet_name, sheet_df in sheets_dataframes.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        csv_file = io.BytesIO()
        sheets_dataframes[selected_sheet].to_csv(csv_file, index=False)
        st.download_button(
            label="Download Cleaned File (CSV)",
            data=csv_file,
            key="cleaned_file_csv",
            file_name=f"{input_file_name}_cleaned.csv",
        )

    # Display the cleaned DataFrame
    st.dataframe(sheets_dataframes[selected_sheet])

    # Display a summary of applied cleaning operations
    st.sidebar.markdown("### Summary of Cleaning Operations")
    if "Replace" in clean_options:
        st.sidebar.write("- Replaced null values with custom value.")
    if "Remove Duplicate" in clean_options:
        st.sidebar.write("- Removed duplicate rows.")
    if "Remove Missing Values" in clean_options:
        st.sidebar.write("- Removed rows with missing values.")
    if "Convert to Lowercase" in clean_options:
        st.sidebar.write(f"- Converted selected columns to lowercase: {', '.join(selected_columns)}")
    if "Delete Columns" in clean_options:
        st.sidebar.write(f"- Deleted selected columns: {', '.join(columns_to_delete)}")
    if "Sort Column" in clean_options:
        st.sidebar.write(f"- Sorted column '{column_to_sort}' in {sort_order.lower()} order.")
    if "Capitalize Columns" in clean_options:
        st.sidebar.write(f"- Capitalized selected columns: {', '.join(selected_columns_to_capitalize)}")


# Add a note explaining the features and aim of the project in the sidebar
st.sidebar.markdown("### What Does Data Tune Do? 🧹")
st.sidebar.markdown("""
Data Tune is a powerful ML Ops tool that provides an intuitive interface for data preparation and cleaning of tabular data files such as CSV, XLS, and XLSX. With Data Tune, you can seamlessly perform a wide range of data tuning operations, including:

- ✨ **Null Value Replacement:** Easily replace missing values with custom data.
- 🔄 **Duplicate Row Removal:** Identify and remove duplicate rows from your dataset.
- 🗑️ **Missing Value Elimination:** Quickly eliminate rows with missing data.
- 🔡 **Text Conversion to Lowercase:** Transform text data to lowercase, enhancing consistency.
- ❌ **Column Deletion:** Select and remove specific columns from your dataset.
- 🔀 **Column Sorting:** Sort columns in ascending or descending order to enhance data organization.
- 📝 **Column Capitalization:** Capitalize the first letters of text data in chosen columns.

Data Tune streamlines the data preparation process for machine learning by providing clean and well-structured data for model training. You can download your refined data in the same format as your input file or in CSV format.
""")

st.sidebar.markdown("🚀 Created by [Anas Khan](https://github.com/anxkhn)")

