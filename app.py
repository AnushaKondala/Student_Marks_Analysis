import matplotlib
matplotlib.use('Agg')  # Set the backend before importing pyplot
import matplotlib.pyplot as plt
from flask import Flask, render_template, request
import pandas as pd
import textwrap

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    # Check if file was uploaded
    if 'file' not in request.files:
        return 'No file uploaded'

    file = request.files['file']

    # Read the Excel file
    excel_file = pd.ExcelFile(file)

    # Dictionary to store the results for each sheet
    results = {}

    # Iterate over each sheet in the Excel file
    for sheet_name in excel_file.sheet_names:
        # Read the sheet into a DataFrame
        df = excel_file.parse(sheet_name)
        df.drop(['S.No', 'ROLLNO', 'STUDENT NAME', 'Credits'], inplace=True, axis=1)
        overalldict = {}
        column_dict = {}
        # newdict={}
        backlogs = 0
        f_class = 0
        s_class = 0
        distinction = 0
        p_class = 0

        # Calculate overall statistics
        for j in range(1, df.shape[0]):
            if df.iloc[j][df.shape[1] - 2] > 0.0:
                backlogs += 1

            sgpa = df.iloc[j][df.shape[1] - 1]

            if 7.0 < sgpa < 8.0:
                f_class += 1
            elif 6.0 < sgpa < 7.0:
                s_class += 1
            elif 8.0 < sgpa < 10.0:
                distinction += 1
            else:
                p_class += 1

        total_students = df.shape[0] - 1
        passed_students = total_students - backlogs
        pass_percentage = (passed_students / total_students) * 100
        rounded_pass_percentage = round(pass_percentage, 2)

        overalldict[sheet_name] = [total_students, passed_students, backlogs, distinction, f_class, s_class,
                                   p_class, rounded_pass_percentage]

        # Calculate column-wise statistics
        ind = -1
        for i in range(0, len(df.columns) - 2, 3):
            ind += 3
            pass_count = 0
            fail_count = 0
            absentees = 0
            column_name = df.columns[i]

            for j in range(1, df.shape[0]):
                value = df.iloc[j][ind]

                if value == "AB(F)":
                    absentees += 1

                if isinstance(value, float):
                    value = str(value)

                value_only = ''.join(x for x in value if x.isdigit())

                if value_only and int(value_only) > 40:
                    pass_count += 1
                else:
                    fail_count += 1

            pass_percentage = (pass_count / (df.shape[0] - 1)) * 100
            rounded_pass_percentage = round(pass_percentage, 2)

            column_dict[column_name] = [pass_count, fail_count, absentees, rounded_pass_percentage]

        # Store the results for the current sheet in the results dictionary
        results[sheet_name] = {'overall': overalldict, 'column-wise': column_dict}

        # Generate bar plot for current sheet's column-wise data
        df_columnwise = pd.DataFrame.from_dict(column_dict,
                                               orient='index', columns=['Pass Count', 'Fail Count', 'Absentees',
                                                                         'Pass Percentage'])
        df_columnwise.sort_values(by='Pass Percentage', ascending=False, inplace=True)

        # Plot the bar graph for column-wise data
        plt.figure(figsize=(12, 8))
        plt.bar(df_columnwise.index, df_columnwise['Pass Percentage'])
        plt.xlabel('Subject')
        plt.ylabel('Pass Percentage')
        plt.title(f'Pass Percentage by Subject - {sheet_name}')
        plt.xticks(rotation=45, ha='right')
        plt.gca().set_xticklabels(textwrap.fill(label, 10) for label in df_columnwise.index)
        plt.gca().tick_params(axis='x', labelsize=12)
        plt.tight_layout()

        # Add numbers to the bars
        for i, v in enumerate(df_columnwise['Pass Percentage']):
            plt.text(i, v, str(v) + "%", ha='center', va='bottom', fontweight='bold')

        # Save the plot to a file
        plot_filename = f'static/{sheet_name}_bar_plot.png'
        plt.savefig(plot_filename)
        plt.close()

        # Subject-wise fail count bar plot
        plt.figure(figsize=(12, 8))
        plt.bar(df_columnwise.index, df_columnwise['Fail Count'])
        plt.xlabel('Subject')
        plt.ylabel('Fail Count')
        plt.title(f'Fail Count by Subject - {sheet_name}')
        plt.xticks(rotation=45, ha='right')
        plt.gca().set_xticklabels(textwrap.fill(label, 10) for label in df_columnwise.index)
        plt.gca().tick_params(axis='x', labelsize=12)
        plt.tight_layout()

        # Add numbers to the bars
        for i, v in enumerate(df_columnwise['Fail Count']):
            plt.text(i, v, str(v) + "%", ha='center', va='bottom', fontweight='bold')

        # Save the plot to a file
        plot_file_fail = f'static/{sheet_name}_bar_plot_fail.png'
        plt.savefig(plot_file_fail)
        plt.close()

    # Generate bar plot for overall data
    df_overall = pd.DataFrame.from_dict(
        {sheet_name: results[sheet_name]['overall'][sheet_name] for sheet_name in results.keys()},
        orient='index', columns=['Total Students', 'Passed Students', 'Backlogs', 'Distinction',
                                 'First Class', 'Second Class', 'Pass Class', 'Pass Percentage'])

    # Convert the index values to strings
    df_overall.index = df_overall.index.map(str)


    plt.figure(figsize=(12, 8))
    plt.bar(df_overall.index, df_overall['Pass Percentage'])
    plt.xlabel('Sheet Name')
    plt.ylabel('Pass Percentage')
    plt.title('Overall Pass Percentage by Sheet')
    plt.xticks(rotation=45, ha='right')
    plt.gca().set_xticklabels(textwrap.fill(label, 10) for label in df_overall.index)
    plt.gca().tick_params(axis='x', labelsize=12)
    plt.tight_layout()

    # Add numbers to the bars
    for i, v in enumerate(df_overall['Pass Percentage']):
        plt.text(i, v, str(v) + "%", ha='center', va='bottom', fontweight='bold')

    # Save the overall plot to a file
    overall_plot_filename = 'static/overall_bar_plot.png'
    plt.savefig(overall_plot_filename)
    plt.close()

    new_df = pd.DataFrame.from_dict(
    {sheet_name: results[sheet_name]['overall'][sheet_name][:3] + [sheet_name] for sheet_name in results.keys()},
    orient='index', columns=['Total Students', 'Passed Students', 'Backlogs', 'Sheet Name'])
    # Generate bar plot for new_df
    plt.figure(figsize=(12, 8))
    plt.bar(new_df['Sheet Name'], new_df['Total Students'], label='Total Students', color=	'#FF8A8A')
    plt.bar(new_df['Sheet Name'], new_df['Passed Students'], label='Passed Students', color='#a020f0')
    plt.bar(new_df['Sheet Name'], new_df['Backlogs'], label='Backlogs', color='#93e9be')
    plt.xlabel('Sheet Name')
    plt.ylabel('Count')
    plt.title('Student Statistics')
    plt.xticks(rotation=45, ha='right')
    plt.gca().set_xticklabels(textwrap.fill(label, 10) for label in new_df.index)
    plt.gca().tick_params(axis='x', labelsize=12)
    plt.legend()
    plt.tight_layout()
    new_df_plot_filename = 'static/new_df_bar_plot.png'
    plt.savefig(new_df_plot_filename)
    plt.close()

    # Convert the overall results to a DataFrame
    ''' new_df['Pass Percentage'] = df_overall['Pass Percentage']
    new_df = new_df[['Sheet Name', 'Total Students', 'Passed Students', 'Backlogs', 'Pass Percentage']]'''
    dfs_columnwise = {}
    for sheet_name, sheet_results in results.items():
        columnwise_data = sheet_results['column-wise']
        df_columnwise = pd.DataFrame.from_dict(columnwise_data,
                                               orient='index', columns=['Pass Count', 'Fail Count', 'Absentees',
                                                                         'Pass Percentage'])
        dfs_columnwise[sheet_name] = df_columnwise
    # Render the results page with the plot filenames
    return render_template('result.html', df_overall=df_overall, dfs_columnwise=dfs_columnwise,
                           overall_plot_filename=overall_plot_filename, new_df=new_df,
                           new_df_plot_filename=new_df_plot_filename)

if __name__ == '__main__':
    app.run(debug=True)