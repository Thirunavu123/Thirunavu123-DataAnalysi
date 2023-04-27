# App to upload the file using streamlit using python
import streamlit as st
import pandas as pd
import pdfkit
from st_aggrid import AgGrid
import matplotlib.pyplot as plt
from pandas.plotting import table
import numpy as np

# Import the required packages
from st_aggrid import AgGrid
import matplotlib.pyplot as plt
from pyecharts import options as opts
from pyecharts.charts import Bar
from pyecharts.faker import Faker
import plotly.express as px
# pip install kaleido
import plotly.io as pio
import kaleido

from PIL import Image
import plotly.tools as tls
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

import pdfkit
from fpdf import FPDF
from st_aggrid import AgGrid
import matplotlib.pyplot as plt
from pyecharts import options as opts
from pyecharts.charts import Bar
from pyecharts.faker import Faker
import io
import plotly.io as pio
import plotly.express as px
import plotly.graph_objs as go

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO
import matplotlib.pyplot as plt

from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import Paragraph
import tempfile
import os


######################################
# Create the pdf file for the graph
title = '20000 Leagues Under the Seas'

class PDF(FPDF):
    def header(self):
        # Logo
        self.image('TGF - Logo.png', 10, 8, 33)
        # Arial bold 15
        self.set_font('Arial', 'B', 15)
        # Move to the right
        self.cell(80)
        # Title
        self.cell(30, 10, 'Budget Analysis', 0, 0, 'C')
        # Line break
        self.ln(20)

    def footer(self):
        # Position at 1.5 cm from bottom
        self.set_y(-15)
        # Arial italic 8
        self.set_font('Arial', 'I', 8)
        # Text color in gray
        self.set_text_color(128)
        # Page number
        self.cell(0, 10, 'Page ' + str(self.page_no()), 0, 0, 'C')

    def chapter_title(self, num, label):
        # Arial 12
        self.set_font('Arial', '', 12)
        # Background color
        self.set_fill_color(200, 220, 255)
        # Title
        self.cell(0, 6, 'Graph %d : %s' % (num, label), 0, 1, 'L', 1)
        # Line break
        self.ln(4)

    #def chapter_body(self, name):
    def chapter_body(self, txt):
        # Read text file
        #with open(name, 'rb') as fh:
        #    txt = fh.read().decode('latin-1')
        # Times 12
        self.set_font('Times', '', 12)
        # Output justified text
        self.multi_cell(0, 5, txt)
        # Line break
        self.ln()
        # Mention in italics
        self.set_font('', 'I')
        self.cell(0, 5, '(end of excerpt)')
    
    def bullet(self):
        self.cell(5, 5, '-', 0, 0)

    def add_graph(self, imagefile = None, description='', description_data = '', table_data=None, columnwidth = []):
        # Draw a graph
        self.image(imagefile, x=10, y=self.get_y(), w=100)
        # Move down by 50 units
        self.ln(80)
        # Print the description
        self.set_font('', 'I')
        self.multi_cell(0, 5, description)
        self.ln()
        
        # Add bullet points to the description
        for item in description.split('.'):
            if item.strip():
                self.bullet()
                self.cell(0, 5, item, 0, 1)
        self.ln()

        if table_data is not None:
            # Set font for table
            self.set_font('Times', '', 10)
            # Set column widths
            col_widths = columnwidth
            # Calculate row height
            row_height = self.font_size + 2
            # Determine starting position of table
            x_pos = 50
            y_pos = self.get_y() + 20

            # Print column headers
            for i, header in enumerate(table_data.columns):
                # Check if the length of the header is greater than the column width
                if len(str(header)) > col_widths[i]:
                    # If it is, wrap the header by setting the max_line_height to a value greater than 0
                    self.cell(col_widths[i], row_height, str(header), border=1, max_line_height=10)
                else:
                    # If not, print the header as is
                    self.cell(col_widths[i], row_height, str(header), border=1)
            self.ln()
            # Print table rows
            for row in table_data.values:
                for i, value in enumerate(row):
                    self.cell(col_widths[i], row_height, str(value), border=1)
                self.ln()
            # Move to next line
            self.ln()
            # Move down by 50 units
            #self.ln(40)
            # Print the description
            self.set_font('', 'I')
            self.multi_cell(0, 7, description_data)
            self.ln()

    def print_chapter(self, num, title, name, add_graph=False, graph_description=''):
        self.add_page()
        self.chapter_title(num, title)
        self.chapter_body(name)
        if add_graph:
            self.add_graph(graph_description)




######################################
# Define grid options with scrollbars
grid_options = {
    'enableSorting': True,
    'enableFilter': True,
    'enableColResize': True,
    'enableRangeSelection': True,
    'pagination': 'true',
    'paginationPageSize': 10,
    'height': '200px',
    'width': '100%'
}

######

#############################
# Function to load the uploaded excel files to the session
def excelload(file, sheetname, rowstoskip):
    df = pd.read_excel(file, engine='openpyxl', sheet_name = sheetname, skiprows = rowstoskip)
    return(df)

# Function to export the excel file to destination
def excelexport(df, sheetname, targetfile, targetpath):
    df.to_excel(targetpath +"/" +  targetfile, sheet_name = sheetname)
    
# Function to export the excel to the Azure Storage

# Aggregate function extract create the Aggregate dataframe
def aggregatefunc(df, groupnycols, aggcols, aggtype):
    df_agg = df.groupby(groupnycols)[aggcols].aggregate(aggtype).reset_index()
    return(df_agg)

# Get the file names from the uploaded files
def get_file_options(uploaded_file):
    if uploaded_files:
        file_dict = {}
        for uploaded_file in uploaded_files:
            file_dict[uploaded_file.name] = uploaded_file.getvalue()
        file_names = list(file_dict.keys())
        return(file_names)
        #st.write("Uploaded file names:")
        #st.write(file_names)

############################
# Page related program
# Logo on the Page
image = Image.open('D:\\Radiare\\Projects\\TGF\\Thowfeek\\Streamlit\\Output\\TGF - Logo.png')
st.image(image)

# Title of the Application
st.title("Finance - Budget Analysis Tool - Draft")

# Logo on the Page(sidebar)
image = Image.open('D:\\Radiare\\Projects\\TGF\\Thowfeek\\Streamlit\\Output\\TGF - Logo.png')
st.sidebar.image(image, width = 100, use_column_width=False)

# Upload the files to application
uploaded_files = st.sidebar.file_uploader("Upload your Excel file here", type=["xlsx", "xls"],accept_multiple_files=True)

# Create a dictionary to store the dataframes for each uploaded file
dfs = {}

# Define the expected columns
expected_columns = ['Module','Intervention','Cost Input','Y1 Unit Cost (Payment Currency)',
                    'Y1 Unit Cost (Grant Currency)', 'Y1 Total Quantity', 'Y1 Total Cash Outflow']

# Iterate over the uploaded files and read them as dataframes
for file in uploaded_files:
    try:
        df = pd.read_excel(file, engine='openpyxl', sheet_name='Detailed Budget - Non-HP', skiprows=4)
        grant_name = pd.read_excel(file, engine='openpyxl', sheet_name='Budget Summary', usecols='E', skiprows=7-1, nrows=1, header=None, names=["Value"]).iloc[0]["Value"]
        #dfs[file.name] = df
        st.write("Grant Name: " + grant_name)
        # Check if all expected columns are present in the dataframe
        if set(expected_columns).issubset(df.columns):
            #st.write('All expected columns are present')
            #st.write(f"All expected columns are present in uploded file: {selected_file}"
            #st.write(df.head())  # Display the first few rows of the dataframe
            dfs[file.name] = df
        else:
            st.warning(f"{file.name} columns are not present in the uploaded file " + (set(expected_columns) - set(df.columns)))
    except:
        st.warning(f"{file.name} is not a valid Excel file.")

# Process the data and create the required files
try:
    if list(dfs.keys()):
        # Display the selected file and selected subset of data
        st.write(file.name)
        #st.write(list(dfs.keys()))
        #st.write(dfs[file.name])
        AgGrid(dfs[file.name].head(10),  grid_options=grid_options)
        # The three report types
        groupby_cols = ['Key Module and Intervention','key Implementers','Budget Phasing Trend']
        data = dfs[file.name].copy()
        #st.write(data)
        # create an empty dictionary to store the filtered DataFrames
        agg_dfs = {}
        for grpagg in groupby_cols:
            if grpagg == groupby_cols[0]:
                agg_dfs[grpagg] = aggregatefunc(data, groupnycols = ['Module', 'Intervention'], aggcols = ['Y1-4 Total Cash Outflow'], aggtype = 'sum')
                st.write("Module level aggregated data")
                #AgGrid(agg_dfs[grpagg])
                AgGrid(agg_dfs[grpagg], grid_options=grid_options) 
                # Add color to the bargraph Module
                color_map = {'Module': 'orange', 'Y1-4 Total Cash Outflow': 'orange'}
                color_map = {'Module': 'orange', 'Y1-4 Total Cash Outflow': 'blue', 'Y5-10 Total Cash Outflow': 'green'}


                fig_chart = px.bar(agg_dfs[grpagg], x='Module', y=['Y1-4 Total Cash Outflow','Y1-4 Total Cash Outflow'], title="Module Level Analysis", color_discrete_map=color_map)
                #fig_chart.update_layout(barmode='group')
                st.write(fig_chart)
                #####
                # For fpdf to save the graph temperorily
                # convert the graph to an image in memory
                image_data = pio.to_image(fig_chart, format='png')
                # create a temporary file and write the image data to it
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
                    f.write(image_data)
                    image_file = f.name
                
                
                
            elif grpagg == groupby_cols[1]:
                agg_dfs[grpagg] = aggregatefunc(data, groupnycols = ['Implementer'], aggcols = ['Y1-4 Total Cash Outflow'], aggtype = 'sum')
                #st.write(agg_dfs[grpagg])
                st.write("Implementer level aggregated data")
                AgGrid(agg_dfs[grpagg])
                # Add color to the bargraph Module
                color_map1 = {'Implementer': 'orange', 'Y1-4 Total Cash Outflow': 'orange'}
                color_map1 = {'Implementer': 'orange', 'Y1-4 Total Cash Outflow': 'blue', 'Y5-10 Total Cash Outflow': 'green'}


                fig_chart1 = px.bar(agg_dfs[grpagg], x='Implementer', y=['Y1-4 Total Cash Outflow','Y1-4 Total Cash Outflow'], title="Implementer Level Analysis", color_discrete_map=color_map)
                #fig_chart.update_layout(barmode='group')
                st.write(fig_chart1)
                #####
                # For fpdf to save the graph temperorily
                # convert the graph to an image in memory
                image_data1 = pio.to_image(fig_chart1, format='png')
                # create a temporary file and write the image data to it
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
                    f.write(image_data1)
                    image_file1 = f.name
                
            elif grpagg == groupby_cols[2]:
                #Data preparation for graph
                yearbudcols = ['Y1 Total Cash Outflow','Y2 Total Cash Outflow', 'Y3 Total Cash Outflow']
                intgrap_data = data[yearbudcols]
                # create new dataframe with the sum of each column in df
                intgrap_data = pd.DataFrame(intgrap_data.sum())
                # reset index and rename columns
                agg_dfs[grpagg] = intgrap_data.reset_index().rename(columns={'index': 'Year', 0: 'Values'})
                #st.write(agg_dfs[grpagg])
                st.write("Year wise aggregated data")
                AgGrid(agg_dfs[grpagg])
                
                # Define a custom color sequence
                color_sequence = ['#FF5733', '#33FFB5', '#B533FF']

                # Create the line chart
                fig_chart4 = px.line(agg_dfs[grpagg], x='Year', y='Values',  color_discrete_sequence=color_sequence)

                # Update the chart layout
                fig_chart4.update_layout(title='Year wise Total Cash Outflow', xaxis_title='Year', yaxis_title='Values')

                # Display the chart
                st.plotly_chart(fig_chart4)
                
                #####
                # For fpdf to save the graph temperorily
                # convert the graph to an image in memory
                image_data2 = pio.to_image(fig_chart4, format='png')
                # create a temporary file and write the image data to it
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
                    f.write(image_data2)
                    image_file2 = f.name
                
        
        ################################
        # create an empty list to store max values for the coumn width
        max_values_list = []
        
        df = agg_dfs[groupby_cols[0]].copy()
        df = df[['Module','Y1-4 Total Cash Outflow']]
        df['Y1-4 Total Cash Outflow'] = df['Y1-4 Total Cash Outflow'].round(2)
        # iterate over columns
        for col in df.columns:
            # create list of column name and values
            col_list = df[col].tolist()
            col_list.insert(0, col)
            # get max length of values
            max_length = len(str(max([x for x in col_list if isinstance(x, str)], key=len)))
            # append column name and max value to max_values_list
            max_values_list.append( max_length + 10)
            #st.write(max_values_list)
        max_values_list = [100,40]
        
        max_values_list1 = []
        df1 = agg_dfs[groupby_cols[1]].copy()
        df1['Y1-4 Total Cash Outflow'] = df1['Y1-4 Total Cash Outflow'].round(2)
        # iterate over columns
        for col in df1.columns:
            # create list of column name and values
            col_list = df1[col].tolist()
            col_list.insert(0, col)
            # get max length of values
            max_length = len(str(max([x for x in col_list if isinstance(x, str)], key=len)))
            # append column name and max value to max_values_list
            max_values_list1.append( max_length + 40)
        
        max_values_list2 = []
        df2 = agg_dfs[groupby_cols[2]].copy()
        df2['Values'] = df2['Values'].round(2)
        # iterate over columns
        for col in df2.columns:
            # create list of column name and values
            col_list = df2[col].tolist()
            col_list.insert(0, col)
            # get max length of values
            max_length = len(str(max([x for x in col_list if isinstance(x, str)], key=len)))
            # append column name and max value to max_values_list
            max_values_list2.append( max_length + 40)
                    
        pdf = PDF()
        pdf.set_title(title)
        pdf.set_author('Jules Verne')
        pdf.print_chapter(1, 'Module Level Cash Outflow for Granth: ' + grant_name, 'A bar graph is a graphical representation of information. It uses bars that extend to different heights to depict value. Bar graphs can be created with vertical bars, horizontal bars, grouped bars (multiple bars that compare values in a category), or stacked bars (bars containing multiple types of information)', add_graph=False, graph_description='A bar graph is a graphical representation of information. It uses bars that extend to different heights to depict value. Bar graphs can be created with vertical bars, horizontal bars, grouped bars (multiple bars that compare values in a category), or stacked bars (bars containing multiple types of information)')
        # Add graph with dataframe
        pdf.add_graph(imagefile = image_file, description='Explanation about the graph. Can customised to create the bullet points about the data. Can also create the some notes here.', description_data = 'Note: Description of the data table. Can also write paragraphs or bullet points here. Mostly about the data table or notes about the grant etc...', table_data=df, columnwidth = max_values_list)
        #pdf.write_table( dataframe=df)
        pdf.print_chapter(2, 'Implementer Level Cash Outflow for Granth: ' + grant_name, 'A bar graph is a graphical representation of information. It uses bars that extend to different heights to depict value. Bar graphs can be created with vertical bars, horizontal bars, grouped bars (multiple bars that compare values in a category), or stacked bars (bars containing multiple types of information).', add_graph=False, graph_description='Description of the second graph The method used to print the paragraphs is multi_cell. Each time a line reaches the right extremity of the cell or a carriage return character is met.')
        # Add graph with dataframe
        pdf.add_graph(imagefile = image_file1, description= 'Explanation about the graph. Can customised to create the bullet points about the data. Can also create the some notes here.', description_data = 'Note: Description of the data table. Can also write paragraphs or bullet points here. Mostly about the data table or notes about the grant etc...', table_data=df1, columnwidth = max_values_list1)
        pdf.print_chapter(3, 'Year wise Cash Outflow for Grant : ' + grant_name, 'A line graph also known as a line plot or a line chart is a graph that uses lines to connect individual data points. A line graph displays quantitative values over a specified time interval. In finance, line graphs are commonly used to depict the historical price action of an asset or security', add_graph=False, graph_description='Description of the second graph The method used to print the paragraphs is multi_cell. Each time a line reaches the right extremity of the cell or a carriage return character is met.')
        # Add graph with dataframe
        pdf.add_graph(imagefile = image_file2, description= 'Explanation about the graph. Can customised to create the bullet points about the data. Can also create the some notes here.', description_data = 'Note: Description of the data table. Can also write paragraphs or bullet points here. Mostly about the data table or notes about the grant etc...', table_data=df2, columnwidth = max_values_list2)

        pdf.output("D:\\Radiare\\Projects\\TGF\\Thowfeek\\Streamlit\\Output\\Finance_" + file.name + ".pdf", 'F')
        
        # Save the PDF to a file
        #with open("D:\\Radiare\\Projects\\TGF\\Thowfeek\\Streamlit\\Output\\Finance_" + file.name + ".pdf", 'wb') as pdf_file:
        #    pdf_file.write(pdf_bytes.getvalue())
        #st.download_button("Download PDF", pdf)
        with open("D:\\Radiare\\Projects\\TGF\\Thowfeek\\Streamlit\\Output\\Finance_" + file.name + ".pdf", "rb") as pdf_file_download_latest:
            PDFbyte_latest = pdf_file_download_latest.read()
        
        if st.download_button(label = "Export_Report",
                            data=PDFbyte_latest,
                            file_name = "Finance_" + file.name + ".pdf", #file_options[0].split(".")[-2] + "_" + intermediate_file[1] + ".pdf",
                            mime='application/octet-stream'):
            #remove the temporary file
            os.unlink(image_file)
            os.unlink(image_file1)
            os.unlink(image_file2)
            st.stop()  # Stop the app execution after the file is downloaded
        
        # remove the temporary file
        #os.unlink(image_file)
        #os.unlink(image_file1)
        #os.unlink(image_file2)
        
        #################################
except Exception as error:
    st.write('An error occurred while processing the Excel file:', error)

st.stop()  # Stop the app execution after the file is downloaded

