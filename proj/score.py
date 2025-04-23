import os
import pandas as pd
import requests
from flask import Flask, render_template, request, session, redirect, url_for, send_file

app = Flask(__name__)
app.secret_key = 'sk_2b9e918f_4b47_4a84_830d_f04e52c29fa4'  # Randomly generated dummy secret key

UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

GEMINI_API_KEY = 'AIzaSyCbpnOZc9kkesTyJGlpyJcQZKZ3G-U_PMo'

# 📊 Read and convert Excel content
def extract_file_content_for_prompt(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip().str.lower()  # Clean column names
        full_data = df.to_markdown(index=False)  # Full data to be used for the answer
        return full_data, df  # Return full data and dataframe
    except Exception as e:
        return f"Error reading Excel: {e}", None

# 🤖 Generate AI answer using Gemini
# 🤖 Generate AI answer using Gemini
def generate_answer(query, file_path=None):
    full_data, df = "", None

    if file_path:
        full_data, df = extract_file_content_for_prompt(file_path)

    # Apply operations based on the user input query
    df, operations_result = apply_operations(df, query)

    # Update the prompt with the operation details
    prompt = f"""
You are an intelligent Excel assistant with full control to understand and manipulate spreadsheet data based on natural language instructions from the user.

Here is the full Excel data:

{full_data}

The user may speak in natural language, casually or with vague instructions. You must:
- Interpret the user's intent clearly, even if they don't use exact column names
- Create new columns if the user asks or implies one
- Delete columns if the user wants them removed
- Rename columns based on meaning
- Modify values in columns based on instructions (like "add 500", "set to 0", "increase Bonus", etc.)
- Perform math operations like sum, average, multiplication if requested

Here is the user's request:
\"\"\" {query} \"\"\" 

Be creative, smart, and accurate. and also recognise the past prompt enter when perform new operation also add previous operation 
"""

    # Send the prompt to Gemini API for generating content
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GEMINI_API_KEY}"
    headers = {"Content-Type": "application/json"}
    data = {
        "contents": [{
            "parts": [{"text": prompt}]
        }]
    }

    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 200:
        result = response.json()
        try:
            # Store the updated dataframe in the session after applying AI's response
            session['updated_df'] = df.to_dict()  # Store as dictionary to persist the changes
            return result['candidates'][0]['content']['parts'][0]['text'], df
        except (KeyError, IndexError):
            return "Error: Unexpected response structure from Gemini API.", df
    else:
        return f"Error: {response.status_code} - {response.text}", df


# Handle adding and modifying columns
def apply_operations(df, user_input):
    try:
        operations_result = []  # To hold the details of performed operations

        # Add a new column (if user requests it)
        if "add column" in user_input.lower():
            column_name = user_input.split('add column')[-1].strip()
            if column_name not in df.columns:
                df[column_name] = None  # Adding a new column with default None values
                operations_result.append(f"Column '{column_name}' added.")

        # Apply calculation to a column (sum, avg, etc.)
        elif "apply calculation" in user_input.lower():
            parts = user_input.split('apply calculation')[-1].strip().split('to')
            if len(parts) == 2:
                calculation = parts[0].strip().lower()
                column_name = parts[1].strip().lower()

                if column_name in df.columns:
                    if calculation == "sum":
                        df[column_name + '_sum'] = df[column_name].sum()
                        operations_result.append(f"Applied sum to column '{column_name}'.")
                    elif calculation == "average":
                        df[column_name + '_average'] = df[column_name].mean()
                        operations_result.append(f"Applied average to column '{column_name}'.")
                    else:
                        operations_result.append(f"Unsupported calculation: {calculation}.")
                else:
                    operations_result.append(f"Column '{column_name}' does not exist.")
            else:
                operations_result.append("Invalid calculation format. Please use 'apply calculation [sum/average] to [column_name]'.")
        
        # Delete a column (if user requests it)
        elif "delete column" in user_input.lower():
            column_name = user_input.split('delete column')[-1].strip()
            if column_name in df.columns:
                df.drop(columns=[column_name], inplace=True)
                operations_result.append(f"Column '{column_name}' deleted.")
            else:
                operations_result.append(f"Column '{column_name}' does not exist.")

        # Rename a column (if user requests it)
        elif "rename column" in user_input.lower():
            parts = user_input.split('rename column')[-1].strip().split('to')
            if len(parts) == 2:
                old_name = parts[0].strip().lower()
                new_name = parts[1].strip().lower()
                if old_name in df.columns:
                    df.rename(columns={old_name: new_name}, inplace=True)
                    operations_result.append(f"Column '{old_name}' renamed to '{new_name}'.")
                else:
                    operations_result.append(f"Column '{old_name}' does not exist.")
            else:
                operations_result.append("Invalid rename format. Please use 'rename column [old_name] to [new_name]'.")
        else:
            operations_result.append("Invalid operation or format. Please specify a valid operation.")

        return df, "\n".join(operations_result)

    except Exception as e:
        return df, f"Error: {e}"

@app.route("/", methods=["GET", "POST"])
def index():
    answer = ""
    user_input = ""
    file_name = ""
    df = None  # To hold the dataframe

    # Handle file upload
    if request.method == "POST":
        user_input = request.form.get("question", "")
        uploaded_file = request.files.get("file")
        
        # Handle file upload
        if uploaded_file:
            file_name = uploaded_file.filename
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
            uploaded_file.save(file_path)
            session['uploaded_file'] = file_path  # Store the file path in session

        # Use file if already uploaded
        file_path = session.get('uploaded_file')
        if file_path and os.path.exists(file_path):
            full_data, df = extract_file_content_for_prompt(file_path)
            if df is not None:
                # Generate response and apply operations
                answer, df = generate_answer(user_input, file_path)
                # Save the modified dataframe back to a file
                df.to_excel(file_path, index=False)
                # Convert the updated dataframe to HTML for rendering
                df_html = df.to_html(classes='table table-bordered table-striped table-sm', index=False)

                # Save updated dataframe to session (for future retrieval)
                session['updated_df'] = df.to_dict()  # Store as dictionary to persist the changes
            else:
                answer = "Error reading the Excel file."
        else:
            answer = "Please upload an Excel file first."

    # If dataframe exists in session, use it
    if 'updated_df' in session:
        df = pd.DataFrame(session['updated_df'])  # Load the dataframe from session

    # Pass the answer and table HTML to the template
    if df is not None and not df.empty:
        df_html = df.to_html(classes='table table-bordered table-striped table-sm', index=False)
        return render_template("index.html", answer=answer, question=user_input, file_name=file_name, df_html=df_html)
    else:
        return render_template("index.html", answer=answer, question=user_input, file_name=file_name)

@app.route("/download/<filename>")
def download(filename):
    try:
        return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)
    except FileNotFoundError:
        return "File not found", 404

if __name__ == "__main__":
    app.run(debug=True)