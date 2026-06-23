from flask import Flask, request, send_file, after_this_request
from flask_cors import CORS
from main import process_file
from cleaner import (
    clean_sales_file,
    clean_purchase_file,
    clean_jv_file,
    clean_ho_exp_file,
    clean_sundry_creditors_file,
    save_cleaned_workbook
)
import os

app = Flask(__name__)
# Enable CORS for all routes, exposing Content-Disposition so frontend can read it if needed
CORS(app, expose_headers=["Content-Disposition"])

@app.route('/clean', methods=['POST'])
def clean_data():
    if 'file' not in request.files:
        return "No file uploaded", 400
    
    file = request.files['file']
    profile = request.form.get('profile', 'debtor')
    
    # Map React profile options to Python profile keys
    if profile in ('newstock', 'sonepat', 'stock'):
        profile = 'stock'
    elif profile == 'debtors':
        profile = 'debtor'
    
    file_path = f"temp_{file.filename}"
    file.save(file_path)
    
    cleaned_file_path = None
    try:
        if profile in ('sales', 'purchase', 'jv', 'ho_exp', 'creditors'):
            sheet_name_map = {
                'sales': 'Cleaned_Sales',
                'purchase': 'Cleaned_Purchase',
                'jv': 'Cleaned_JV',
                'ho_exp': 'Cleaned_HO_Exp',
                'creditors': 'Cleaned_Creditors'
            }
            cleaned_sheet_name = sheet_name_map[profile]
            cleaned_file_path = f"temp_cleaned_{file.filename}"
            
            if profile == 'sales':
                raw_df, cleaned_df, changes = clean_sales_file(file_path)
            elif profile == 'purchase':
                raw_df, cleaned_df, changes = clean_purchase_file(file_path)
            elif profile == 'jv':
                raw_df, cleaned_df, changes = clean_jv_file(file_path)
            elif profile == 'ho_exp':
                raw_df, cleaned_df, changes = clean_ho_exp_file(file_path)
            elif profile == 'creditors':
                raw_df, cleaned_df, changes = clean_sundry_creditors_file(file_path)
                
            save_cleaned_workbook(raw_df, cleaned_df, changes, cleaned_file_path, cleaned_sheet_name)
        else:
            # Main.py logic (stock / debtors)
            cleaned_file_path = process_file(file_path, profile_key=profile)
        
        # Register a callback to delete the temp files after request finishes
        @after_this_request
        def remove_file(response):
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                if cleaned_file_path and os.path.exists(cleaned_file_path):
                    os.remove(cleaned_file_path)
            except Exception as e:
                app.logger.error(f"Error removing temp files: {e}")
            return response
        
        return send_file(cleaned_file_path, as_attachment=True, download_name=os.path.basename(cleaned_file_path))
    except Exception as e:
        # Clean up files in case of error
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
            if cleaned_file_path and os.path.exists(cleaned_file_path):
                os.remove(cleaned_file_path)
        except:
            pass
        return f"Processing Error: {str(e)}", 500

if __name__ == '__main__':
    app.run(port=5000, debug=True)