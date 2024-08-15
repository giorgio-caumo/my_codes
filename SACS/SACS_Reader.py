import pandas as pd
from SACS_output_v3 import SACSReader

# THISDIR = Path(__file__).parent
excel_file = r'SACS output_TEST.xlsx'

database_file = [r'Main pile\sacsdb.post']


selected_joints = None
selected_members = ['SD01-TOG', 'SD02-TOG', 'SD03-TOG', 'SD04-TOG', 'SD05-TOG', 'SD06-TOG', 'SD07-TOG', 'SD08-TOG',
                    'SD01-D01', 'SD02-D02', 'SD03-D03', 'SD04-D04', 'SD05-D05', 'SD06-D06', 'SD07-D07', 'SD08-D08'
                            ]
index = ["Operational"]

post_process_data = []

ii = 0
for i_database in database_file:
    sacs_output = SACSReader(index=index[ii], database_file=i_database, selected_joints=selected_joints,
                             selected_members=selected_members, excel_file_path=excel_file)

    sacs_output.get_data()
    sacs_output.plot_forces(data='Spring Joint Forces', selection='Z')
    # sacs_output.export_to_excel(selection='Fixed Joint Reactions')
    # sacs_output.export_to_excel(selection='Spring Joint Forces')
    # sacs_output.export_to_excel(selection='Member End Forces')
    # sacs_output.export_to_excel(selection='Member UC')

    sacs_output.export_to_pdf_styled()
    # Post processing
    # df_1 = sacs_output.post_process_fixed_reactions(selection='Max', direction='X', selected_joints=None)
    # df_2 = sacs_output.post_process_member_forces(selection='Max', direction='Z',
    #                                               selected_members=['SD01-TOG', 'SD02-TOG', 'SD03-TOG', 'SD04-TOG',
    #                                                                 'SD05-TOG', 'SD06-TOG', 'SD07-TOG', 'SD08-TOG'])
    # df_3 = sacs_output.post_process_fixed_reactions(selection='Min', direction='Y', selected_joints=None)
    #
    # # Append each DataFrame to the post_process_data list
    # post_process_data.append(df_1)
    # post_process_data.append(df_2)
    # post_process_data.append(df_3)
    #
    # # Concatenate all DataFrames in the post_process_data list
    # summary_df = pd.concat(post_process_data, ignore_index=True)
    # # Save the concatenated DataFrame to Excel
    # sheet_name = f"Summary - {index[ii]}"
    # sacs_output.save_results_to_excel(sheet_name = sheet_name, df = summary_df)
    # Export the PDF

    # Clear the post_process_data list for the next iteration
    post_process_data.clear()

    ii+=1

