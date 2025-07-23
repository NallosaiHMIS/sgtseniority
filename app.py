from flask import Flask, request, send_file, render_template
import pandas as pd

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/process', methods=['POST'])
def process_file():
    file = request.files['excel_file']
    df = pd.read_excel(file)

    # Clean columns
    df.columns = df.columns.str.strip().str.replace('\n', '', regex=False).str.replace('\r', '', regex=False)

    # Convert dates (with dayfirst=True)
    df['District Transfer'] = pd.to_datetime(df['District Transfer'], errors='coerce', dayfirst=True)
    df['Date of regularazation'] = pd.to_datetime(df['Date of regularazation'], errors='coerce', dayfirst=True)
    df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce', dayfirst=True)

    # Standardize and clean other columns
    df['Mode'] = df['Mode'].astype(str).str.strip().str.upper()
    df['TRB Year clean'] = df['TRB Year clean'].astype(str).str.strip()
    df['TRB Rank'] = pd.to_numeric(df['TRB Rank'], errors='coerce')

    # Separate TRB and Non-TRB
    df_trb = df[(df['Mode'] == 'TRB') & (df['TRB Rank'] > 0)].copy()
    df_non_trb = df[df['Mode'] != 'TRB'].copy()

    # TRB logic: same batch date + rank based seniority
    df_trb = df_trb.sort_values(by=['TRB Year clean', 'TRB Rank', 'DOB'])
    df_trb['Effective Joining Date'] = df_trb.groupby('TRB Year clean')['Date of regularazation'].transform('min')

    # Non-TRB: use District Transfer or Date of regularazation
    df_non_trb['Effective Joining Date'] = df_non_trb['District Transfer'].combine_first(df_non_trb['Date of regularazation'])

    # Combine and sort
    df_combined = pd.concat([df_trb, df_non_trb], ignore_index=True)
    df_sorted = df_combined.sort_values(by=['Effective Joining Date', 'DOB', 'TRB Rank'])
    df_sorted['Seniority Rank'] = range(1, len(df_sorted) + 1)

    # Format dates as dd-mm-yyyy
    for col in ['DOB', 'Date of regularazation', 'District Transfer', 'Effective Joining Date']:
        df_sorted[col] = df_sorted[col].apply(lambda x: x.strftime('%d-%m-%Y') if pd.notnull(x) else '')

    output_path = 'seniority_ranked_result.xlsx'
    df_sorted.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
