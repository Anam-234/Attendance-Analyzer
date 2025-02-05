import os
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# Configure Upload Folders
UPLOAD_FOLDER = './uploads'
STATIC_FOLDER = './static'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(STATIC_FOLDER, exist_ok=True)

# Function to check allowed file extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Function to format seconds into hh:mm:ss
def seconds_to_hms(seconds):
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    seconds = int(seconds % 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"

# Function to categorize work time
def categorize_time(hours):
    if pd.isnull(hours):
        return "Invalid Time"
    if hours < 4.5:
        return "Half Day"
    elif 4.5 <= hours <= 8.5:
        return "Regularization"
    elif 8.5 <= hours <= 9.0:
        return "Full Day"
    else:
        return "Overtime"

# File Processing Function
def process_file(file):
    excel_data = pd.ExcelFile(file)
    sheet_name = 'First In And Last Out'
    data = excel_data.parse(sheet_name)

    if data.empty:
        st.error("No data found in the Excel sheet.")
        return None, None, None
    
    cleaned_data = data.rename(columns={
        'First In And Last Out': 'Personnel ID',
        'Unnamed: 1': 'First Name',
        'Unnamed: 2': 'Last Name',
        'Unnamed: 3': 'First in Reader Name',
        'Unnamed: 4': 'First In Time',
        'Unnamed: 5': 'Last Out Reader Name',
        'Unnamed: 6': 'Last Out Time',
        'Unnamed: 7': 'Department Name'
    }).iloc[1:]

    cleaned_data['First In Time'] = pd.to_datetime(cleaned_data['First In Time'], errors='coerce')
    cleaned_data['Last Out Time'] = pd.to_datetime(cleaned_data['Last Out Time'], errors='coerce')

    cleaned_data['Total Time (seconds)'] = (
        (cleaned_data['Last Out Time'] - cleaned_data['First In Time']).dt.total_seconds()
    ).fillna(0)

    cleaned_data['Time Done'] = cleaned_data['Total Time (seconds)'].apply(seconds_to_hms)
    cleaned_data['Name'] = cleaned_data['First Name'] + " " + cleaned_data['Last Name']
    cleaned_data['Date'] = cleaned_data['First In Time'].dt.date
    cleaned_data['Work Mode'] = cleaned_data['Total Time (seconds)'].apply(lambda x: categorize_time(x / 3600))

    output_data = cleaned_data[['Personnel ID', 'Name', 'Date', 'Time Done', 'Work Mode']]

    # Generate Pie Chart for Work Categories
    work_category_counts = cleaned_data['Work Mode'].value_counts()
    fig, ax = plt.subplots()
    ax.pie(work_category_counts, labels=work_category_counts.index, autopct='%1.1f%%')
    plt.title('Work Categories')
    chart_path = os.path.join(STATIC_FOLDER, 'pie_chart.png')
    plt.savefig(chart_path)
    plt.close()

    return output_data, None, chart_path

# Streamlit App
st.markdown(
    """
    <style>
    .reportview-container {
        background-color: #eaf4e5;  /* Light green */
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Add company logo
st.image("static/company_logo.png", width=200)

st.title("Mini Done - Calculate Working Hours of Interns")

# File Upload
st.subheader("Upload Excel File")
uploaded_file = st.file_uploader("Choose an Excel file", type=['xls', 'xlsx'])

if uploaded_file and allowed_file(uploaded_file.name):
    # Process the uploaded file
    cleaned_data, output_path, chart_path = process_file(uploaded_file)
    
    # Display results
    st.subheader("Results")
    st.dataframe(cleaned_data)
    st.image(chart_path)

    # Save output to Excel
    output_file_path = os.path.join(UPLOAD_FOLDER, 'processed_output.xlsx')
    cleaned_data.to_excel(output_file_path, index=False)
    st.success(f"Processed data saved to {output_file_path}")

    # Download link for the processed Excel file
    st.download_button(
        label="Download Processed Data",
        data=open(output_file_path, 'rb').read(),
        file_name='processed_output.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
