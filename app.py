import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Function to upload file
def upload_file():
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])
    if uploaded_file is not None:
        return uploaded_file
    return None

# Function to process the uploaded file
def process_file(file):
    excel_data = pd.ExcelFile(file)
    sheet_name = 'First In And Last Out'
    data = excel_data.parse(sheet_name)

    cleaned_data = data.rename(columns={
        'First In And Last Out': 'Personnel ID',
        'Unnamed: 1': 'First Name',
        'Unnamed: 2': 'Last Name',
        'Unnamed: 3': 'First in Reader Name',
        'Unnamed: 4': 'First In Time',
        'Unnamed: 5': 'Last Out Reader Name',
        'Unnamed: 6': 'Last Out Time',
        'Unnamed: 7': 'Department Name'
    }).iloc[1:]  # Remove the first row as it's an invalid header row

    # Convert time columns to datetime
    cleaned_data['First In Time'] = pd.to_datetime(cleaned_data['First In Time'], errors='coerce')
    cleaned_data['Last Out Time'] = pd.to_datetime(cleaned_data['Last Out Time'], errors='coerce')

    # Calculate total hours worked in seconds
    cleaned_data['Total Time (seconds)'] = (
        (cleaned_data['Last Out Time'] - cleaned_data['First In Time'])
        .dt.total_seconds()
    ).fillna(0)

    # Function to convert seconds into hh:mm:ss
    def seconds_to_hms(seconds):
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = int(seconds % 60)
        return f"{hours:02}:{minutes:02}:{seconds:02}"

    # Apply conversion function
    cleaned_data['Time Done'] = cleaned_data['Total Time (seconds)'].apply(seconds_to_hms)

    # Categorize work time into work categories
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

    cleaned_data['Work Category'] = cleaned_data['Total Time (seconds)'].apply(lambda x: categorize_time(x / 3600))

    # Add 'Work Mode' column for display purposes
    cleaned_data['Work Mode'] = cleaned_data['Work Category']
    cleaned_data['Name'] = cleaned_data['First Name'] + " " + cleaned_data['Last Name']
    cleaned_data['Date'] = cleaned_data['First In Time'].dt.date

    output_data = cleaned_data[['Personnel ID', 'Name', 'Date', 'Time Done', 'Work Mode', 'Total Time (seconds)', 'Work Category']]

    return cleaned_data, output_data


# Function to visualize data with different chart types
def visualize_data(cleaned_data):
    # Define the categories and filter the data
    categories = ['Full Day', 'Regularization', 'Overtime', 'Half Day']
    work_category_counts = cleaned_data['Work Category'].value_counts()[categories]

    # Let user choose chart type
    chart_type = st.selectbox('Select Chart Type', ['Pie Chart', 'Bar Chart', 'Line Chart'])

    # Pie Chart
    if chart_type == 'Pie Chart':
        colors = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99']
        fig, ax = plt.subplots()
        work_category_counts.plot.pie(autopct='%1.1f%%', startangle=90, ax=ax, colors=colors, labels=categories)
        ax.set_ylabel('')
        st.pyplot(fig)

    # Bar Chart
    elif chart_type == 'Bar Chart':
        fig, ax = plt.subplots()
        work_category_counts.plot.bar(ax=ax, color=['#ff9999', '#66b3ff', '#99ff99', '#ffcc99'])
        ax.set_ylabel('Count')
        ax.set_title('Work Categories Bar Chart')
        st.pyplot(fig)

    

# Main application
def main():
    st.title("Intern Performance Analysis")
    
    # Upload file
    file = upload_file()
    if file:
        cleaned_data, output_data = process_file(file)
        st.markdown("""
        <style>
        [data-testid="stAppViewContainer"]{background-color: #2E4269;}
        </style>""",
                    unsafe_allow_html=True)
        # Display data preview
        st.subheader("Data Preview")
        st.dataframe(cleaned_data[['Personnel ID', 'Name', 'Date', 'Time Done', 'Work Category']])

        # Download processed data
        st.subheader("Download Processed Data")
        csv = output_data.to_csv(index=False).encode('utf-8')
        st.download_button("Download Processed Data", csv, "processed_data.csv", "text/csv")
        
        # Display work categories chart
        st.subheader("Work Categories Chart")
        visualize_data(cleaned_data)

if __name__ == "__main__":
    main()
