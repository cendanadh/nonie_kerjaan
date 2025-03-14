from flask import Flask, render_template, request, send_file
import pandas as pd
import io

app = Flask(__name__)

app.config['ALLOWED_EXTENSIONS'] = {'csv'}

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file1 = request.files.get('file1')
        file2 = request.files.get('file2')
        user_input = request.form.get('wareid')
        user_input_list = user_input.split(',')
        user_input_list_stripped  = [string.strip() for string in user_input_list]

        if file1 and file2:
            file_contents1 = file1.read()
            file_contents2 = file2.read()
            df1 = pd.read_csv(io.BytesIO(file_contents1))
            df2 = pd.read_csv(io.BytesIO(file_contents2))
            df_concate = pd.concat([df1, df2], axis=0)
            df_concate = df_concate[~df_concate['itemname'].str.contains('TIDAK|DIPAKAI|DI PAKAI|DPAKAI|DPKAI|PAKAI|DPIPAKAI|DIPAKE|AFALAN|SAMPLE|KODE BRNG|KODE JGN', na=False, case=False)]
            

            if user_input:

                #Grounding
                grounding = df_concate[df_concate['itemname'].str.contains('ground', na=False, case=False)]
                grounding = grounding[grounding['wareid'].isin(user_input_list_stripped)]
                grounding['merk'] = grounding['itemname'].str.split().str[-1]
                grounding['ukuran'] = grounding['itemname'].str.split().str[1]
                grounding = grounding[['no', 'wareid', 'active', 'itemid', 'itemname', 'ukuran', 'merk', 'convert', 'purid', 'prd', 'endsls', 'endstk', 'endvle', 'opmawl','memo']]

                #Strep
                strep = df_concate[df_concate['itemname'].str.contains('strep tbg|strep tmbg', na=False, case=False)]
                strep = strep[strep['wareid'].isin(user_input_list_stripped)]
                strep['merk'] = strep['itemname'].str.replace(r'.*4M', '', regex=True).str.strip()
                strep['ukuran'] = strep['itemname'].str[10:].str[:-4].str.strip()
                strep = strep[['no', 'wareid', 'active', 'itemid', 'itemname', 'ukuran', 'merk', 'convert', 'purid', 'prd', 'endsls', 'endstk', 'endvle', 'opmawl','memo']]

                #Kawat
                kawat = df_concate[df_concate['itemname'].str.contains('kawat tbg|kawat tmbg', na=False, case=False)]
                kawat = kawat[kawat['wareid'].isin(user_input_list_stripped)]
                kawat['merk'] = kawat['itemname'].str.replace(r'.*MM', '', regex=True).str.strip()
                kawat['ukuran'] = kawat['itemname'].str[10:].str[:-4].str.strip()
                kawat = kawat[['no', 'wareid', 'active', 'itemid', 'itemname', 'ukuran', 'merk', 'convert', 'purid', 'prd', 'endsls', 'endstk', 'endvle', 'opmawl','memo']]

            else:
                #Grounding
                grounding = df_concate[df_concate['itemname'].str.contains('ground', na=False, case=False)]
                grounding['merk'] = grounding['itemname'].str.split().str[-1]
                grounding['ukuran'] = grounding['itemname'].str.split().str[1]
                grounding = grounding[['no', 'wareid', 'active', 'itemid', 'itemname', 'ukuran', 'merk', 'convert', 'purid', 'prd', 'endsls', 'endstk', 'endvle', 'opmawl','memo']]

                #Strep
                strep = df_concate[df_concate['itemname'].str.contains('strep tbg|strep tmbg', na=False, case=False)]
                strep['merk'] = strep['itemname'].str.replace(r'.*4M', '', regex=True).str.strip()
                strep['ukuran'] = strep['itemname'].str[10:].str[:-4].str.strip()
                strep = strep[['no', 'wareid', 'active', 'itemid', 'itemname', 'ukuran', 'merk', 'convert', 'purid', 'prd', 'endsls', 'endstk', 'endvle', 'opmawl','memo']]

                #Kawat
                kawat = df_concate[df_concate['itemname'].str.contains('kawat tbg|kawat tmbg', na=False, case=False)]
                kawat['merk'] = kawat['itemname'].str.replace(r'.*MM', '', regex=True).str.strip()
                kawat['ukuran'] = kawat['itemname'].str[10:].str[:-4].str.strip()
                kawat = kawat[['no', 'wareid', 'active', 'itemid', 'itemname', 'ukuran', 'merk', 'convert', 'purid', 'prd', 'endsls', 'endstk', 'endvle', 'opmawl','memo']]

            output = io.BytesIO()
            
            with pd.ExcelWriter(output) as writer:
                df_concate.to_excel(writer, sheet_name='all_data_cleaned', index=False)
                grounding.to_excel(writer, sheet_name='grounding', index=False)
                strep.to_excel(writer, sheet_name='strep', index=False)
                kawat.to_excel(writer, sheet_name='kawat', index=False)


            output.seek(0)







            return send_file(output, as_attachment=True, download_name="output.xlsx", mimetype="xlsx")
        
        return 'Invalid file format. Please upload a CSV file.'

    return render_template('index.html', table_html=None)

if __name__ == '__main__':
    app.run(debug=True)
