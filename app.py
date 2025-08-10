import os
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, session
import pandas as pd
import uuid 

app = Flask(__name__)
app.secret_key = 'a-very-simple-secret-key'

DOWNLOAD_FOLDER = 'temp_downloads'

if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part in the request.')
            return redirect(request.url)

        file = request.files['file']
        sheet_choice = request.form.get('sheet_choice')

        if file.filename == '':
            flash('No file selected.')
            return redirect(request.url)

        if file and file.filename.endswith('.xlsx'):
            try:
                df_cleaned = None
                headers = []

                if sheet_choice == 'Sheet1':
                    df = pd.read_excel(file, sheet_name='Sheet1', header=None)
                    if df.shape[1] >= 10:
                        df_data = df.iloc[:, 2:10]
                    else:
                        flash('Sheet1 does not have enough columns.')
                        return redirect(request.url)
                    headers = ['Name', 'Email', 'ID', 'Assigned To', 'Status', 'Date', 'Source', 'Type']
                    df_data.columns = headers
                    df_cleaned = df_data.dropna(subset=['Name'])
                    df_cleaned = df_cleaned[df_cleaned['Name'] != 'Preview']

                elif sheet_choice == 'Sheet2':
                    df = pd.read_excel(file, sheet_name='Sheet2', header=None)
                    headers = ['Client Name', 'Email', 'Phone No', 'Owned By', 'Status', 'Created At', 'Source', 'Contact Type']
                    final_data_list = []
                    block_data_list = df.iloc[0:1535, 0].dropna().tolist()
                    for i in range(0, len(block_data_list), 9):
                        if i + 8 < len(block_data_list):
                            chunk = block_data_list[i : i + 9]
                            record = {'Client Name': chunk[0], 'Email': chunk[2], 'Phone No': chunk[3], 'Owned By': chunk[4], 'Status': chunk[5], 'Created At': chunk[6], 'Source': chunk[7], 'Contact Type': chunk[8]}
                            final_data_list.append(record)
                    structured_df = df.iloc[1535:1668].copy()
                    if structured_df.shape[1] >= 10:
                        structured_df_data = structured_df.iloc[:, 2:10]
                        structured_df_data.columns = headers
                        structured_df_cleaned = structured_df_data.dropna(subset=['Client Name'])
                        final_data_list.extend(structured_df_cleaned.to_dict(orient='records'))
                    df_cleaned = pd.DataFrame(final_data_list)
                if df_cleaned is not None and not df_cleaned.empty:
                    
                    unique_filename = f"Converted.xlsx"
                    filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], unique_filename)
               
                    df_cleaned.to_excel(filepath, index=False)
                    
                    session['download_filename'] = unique_filename
                    
                    data_for_html = df_cleaned.to_dict(orient='records')
                    return render_template('index.html', data=data_for_html, headers=headers)
                else:
                    flash('No data was processed. Please check the file and sheet.')
                    return redirect(request.url)

            except Exception as e:
                flash(f"An error occurred: {e}")
                return redirect(request.url)
        else:
            flash('Please upload a valid .xlsx file.')
            return redirect(request.url)

    session.pop('download_filename', None) 
    return render_template('index.html', data=None)

@app.route('/download')
def download_file():
    filename = session.get('download_filename', None)

    if filename is None:
        flash("No file available to download. Please process a file first.")
        return redirect(url_for('index'))

    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
