import pandas as pd
import os

from create_minutes import *
from create_roster import create_roster

def read(excel_file):
    df = pd.read_excel(excel_file, header=1)
    active_df = df[df['Status'] == 'Active'][['Last Name', 'First Name', 'Current Office']]
    advisor_df = df[df['Current Office'].isin(advisors)][['Last Name', 'First Name', 'Current Office']]

    return active_df, advisor_df

def write(active_df, advisor_df):
    docx_output_dir = 'Minutes'
    xlsx_output_dir = 'Rosters'
    
    os.makedirs(docx_output_dir, exist_ok=True)
    os.makedirs(xlsx_output_dir, exist_ok=True)

    create_roster(xlsx_output_dir, active_df, advisor_df)

    create_bylaws_minutes(docx_output_dir, active_df)
    create_chapter_minutes(docx_output_dir, active_df, advisor_df)
    create_events_minutes(docx_output_dir, active_df)
    create_exec_minutes(docx_output_dir, active_df)
    create_finance_minutes(docx_output_dir, active_df)
    create_house_minutes(docx_output_dir, active_df, advisor_df)
    create_IOC_minutes(docx_output_dir, active_df)

def main():
    excel_file = 'data/Spring Roster 2025.xlsx'
    active_df, advisor_df = read(excel_file)
    write(active_df, advisor_df)

if __name__ == "__main__":
    main()
