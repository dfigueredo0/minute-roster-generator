import os

def create_roster(xlsx_output_dir, active_df, advisor_df):
    active_df.to_excel(os.path.join(xlsx_output_dir, 'Officer Roster and Minutes Rosters.xlsx'), index=False)