import plotly.express as px
import plotly
import pandas as pd
import xlrd
from openpyxl import Workbook
import plotly.figure_factory as ff


# Read Dataframe from Excel file
# df = Workbook('task.xlsx')

df = pd.read_excel('task.xlsx')
# df = xlrd.open_workbook('task.xlsx')

# Assign Columns to variables
tasks = df['Task']
start = df['Start']
finish = df['Finish']
complete = df['Complete in %']

print(complete)

# Create Gantt Chart
fig = px.timeline(df, x_start=start, x_end=finish, y=tasks, color=complete)

""" fig = px.timeline(df, x_start=start, x_end=finish, y=tasks, color=complete, title='Task Overview', 
color_continuous_scale = [(0, "red"), (0.5, "yellow"), (1, "green")]) """

# Update Change Layout - Optional
fig.update_yaxes(autorange='reversed')
fig.update_layout(
    title_font_size=42,
    font_size=18,
    title_font_family='Arial'
)

#  Interactive Gyntt
fig = ff.create_gantt(df)

#  Save Graph and export to HTML
plotly.offline.plot(fig, filename='Task_Overview_Gantt.html')