import streamlit as st
import pandas as pd
import io

def save_to_excel(df):
    output = io.BytesIO()  # Create a BytesIO buffer
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)  # Go to the beginning of the BytesIO buffer
    return output

# Step 2: Create a file uploader
uploaded_file = st.file_uploader("Choose a input file", type="xlsx")

# Step 3: Check if a file has been uploaded
if uploaded_file is not None:
    # Read the CSV file
    df = pd.read_excel(uploaded_file)
    # manage-df
    df = df.drop(columns=['Email','Name','Last modified time'])
    [row, column] = df.shape
    removed_column = df[['ID','Start time','Completion time','Comment','Name - First Name']]
    df2 = df.drop(columns=['ID','Start time','Completion time','Comment','Name - First Name'])
    # Calculation
    count_in_range = ((df2 >= 2) & (df2 <= 4)).sum()
    percent = count_in_range/row*100
    status = percent.apply(lambda x: 'Pass' if x > 85 else 'Not pass')
    # Insert pre-remove column
    df2.insert(0, 'ID', removed_column['ID'])
    df2.insert(1, 'Start time', removed_column['Start time'])
    df2.insert(2, 'Completion time', removed_column['Completion time'])
    df2.insert(3, 'Name - First Name', removed_column['Name - First Name'])
    df2.insert(column-1, 'Comment', removed_column['Comment'])
    # insert output column
    df2.loc[len(df2)]= percent
    df2.loc[len(df2)]= status

    # Display the DataFrame
    st.write("Uploaded DataFrame:")
    st.dataframe(df2)

    #save
    # Create a download button
    st.title("Download output File")

    excel_file = save_to_excel(df2)  # Get the BytesIO buffer with the Excel file

    st.download_button(
      label="Download Excel output file",
      data=excel_file,
      file_name='output.xlsx',
      mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  )
