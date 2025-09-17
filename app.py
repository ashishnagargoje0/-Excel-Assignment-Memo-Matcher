#!/usr/bin/env python3
"""
Excel Assignment-Memo Line Matcher
A Flask web application for processing Excel files and matching Assignment values with Memo Line entries.
"""

import os
import logging
import pandas as pd
import numpy as np
from flask import Flask, request, render_template_string, send_file, jsonify, flash, redirect, url_for
from werkzeug.utils import secure_filename
from datetime import datetime
import traceback
import re
from typing import List, Tuple, Dict, Any
import json

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Flask app configuration
app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create necessary directories
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Allowed extensions
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def clean_and_normalize_text(text):
    """Clean and normalize text for better matching."""
    if pd.isna(text):
        return ""
    
    # Convert to string
    text_str = str(text).strip()
    
    # Remove extra whitespace
    text_str = re.sub(r'\s+', ' ', text_str)
    
    return text_str

def find_assignment_matches(df: pd.DataFrame) -> List[Tuple[int, int, str, str]]:
    """
    Find matches between Assignment column and Memo Line column.
    Returns list of tuples: (assignment_row_idx, memo_row_idx, assignment_value, memo_value)
    """
    matches = []
    
    # Clean and prepare data
    df_clean = df.copy()
    df_clean['Assignment_clean'] = df_clean['Assignment'].apply(clean_and_normalize_text)
    df_clean['Memo Line_clean'] = df_clean['Memo Line'].apply(clean_and_normalize_text)
    
    logger.info(f"Starting match process with {len(df_clean)} rows")
    
    # Iterate through each assignment
    for assign_idx, assignment_row in df_clean.iterrows():
        assignment_val = assignment_row['Assignment_clean']
        
        if not assignment_val or assignment_val == "":
            continue
            
        logger.debug(f"Checking assignment: '{assignment_val}' (row {assign_idx})")
        
        # Check against all memo lines (excluding the same row)
        for memo_idx, memo_row in df_clean.iterrows():
            if assign_idx == memo_idx:
                continue
                
            memo_val = memo_row['Memo Line_clean']
            
            if not memo_val or memo_val == "":
                continue
            
            # Check if assignment appears in memo line
            if assignment_val.lower() in memo_val.lower():
                matches.append((
                    assign_idx, 
                    memo_idx, 
                    df.iloc[assign_idx]['Assignment'],
                    df.iloc[memo_idx]['Memo Line']
                ))
                logger.info(f"Match found: '{assignment_val}' in '{memo_val}' (rows {assign_idx}-{memo_idx})")
    
    logger.info(f"Total matches found: {len(matches)}")
    return matches

def create_filtered_output(df: pd.DataFrame, matches: List[Tuple[int, int, str, str]]) -> pd.DataFrame:
    """Create filtered output DataFrame with matched pairs."""
    if not matches:
        logger.warning("No matches found, returning empty DataFrame")
        return pd.DataFrame()
    
    filtered_rows = []
    processed_pairs = set()
    
    for assign_idx, memo_idx, assign_val, memo_val in matches:
        pair_key = (min(assign_idx, memo_idx), max(assign_idx, memo_idx))
        
        if pair_key in processed_pairs:
            continue
            
        processed_pairs.add(pair_key)
        
        # Add assignment row first, then memo row
        filtered_rows.append(df.iloc[assign_idx].copy())
        filtered_rows.append(df.iloc[memo_idx].copy())
        
        logger.debug(f"Added pair: Assignment row {assign_idx} and Memo row {memo_idx}")
    
    if filtered_rows:
        result_df = pd.DataFrame(filtered_rows)
        result_df.reset_index(drop=True, inplace=True)
        logger.info(f"Created filtered output with {len(result_df)} rows ({len(filtered_rows)//2} pairs)")
        return result_df
    else:
        return pd.DataFrame()

def generate_insights(matches: List[Tuple[int, int, str, str]], df_info: Dict[str, Any]) -> str:
    """Generate insights about the matching process using simple analytics."""
    try:
        if not matches:
            return "No matches were found between Assignment and Memo Line columns."
        
        # Basic statistics
        unique_assignments = len(set(match[2] for match in matches))
        unique_memos = len(set(match[3] for match in matches))
        total_matches = len(matches)
        
        # Match patterns
        assignment_frequency = {}
        for _, _, assign_val, _ in matches:
            assignment_frequency[assign_val] = assignment_frequency.get(assign_val, 0) + 1
        
        most_frequent = max(assignment_frequency.items(), key=lambda x: x[1]) if assignment_frequency else ("N/A", 0)
        
        insights = f"""
📊 MATCHING ANALYSIS SUMMARY

🔍 Match Statistics:
• Total matches found: {total_matches}
• Unique assignments matched: {unique_assignments}
• Unique memo lines involved: {unique_memos}
• Processing efficiency: {(total_matches / df_info['total_rows'] * 100):.1f}% of rows involved

🎯 Pattern Analysis:
• Most frequently matched assignment: "{most_frequent[0]}" ({most_frequent[1]} times)
• Average matches per assignment: {(total_matches / unique_assignments):.1f}

📋 Data Quality:
• Total input rows: {df_info['total_rows']}
• Output rows (paired): {total_matches * 2}
• Data completeness: Good (all matches are complete pairs)

💡 Recommendations:
• Consider reviewing high-frequency matches for data accuracy
• Filtered output ready for download with {total_matches} matched pairs
• All matches have been logged for audit trail
        """
        
        logger.info("Generated insights summary")
        return insights.strip()
        
    except Exception as e:
        logger.error(f"Error generating insights: {str(e)}")
        return f"Insights generation encountered an error: {str(e)}"

def validate_excel_structure(df: pd.DataFrame) -> Tuple[bool, str]:
    """Validate that the Excel file has the required columns."""
    required_columns = ['Assignment', 'Memo Line']
    
    if df.empty:
        return False, "The uploaded Excel file is empty."
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        available_cols = ", ".join(df.columns.tolist())
        return False, f"Missing required columns: {', '.join(missing_columns)}. Available columns: {available_cols}"
    
    return True, "File structure is valid."

@app.route('/')
def index():
    """Main page with upload form."""
    return render_template_string("""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Assignment-Memo Matcher</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
            line-height: 1.6;
        }
        .container {
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .header {
            text-align: center;
            margin-bottom: 30px;
        }
        .header h1 {
            color: #333;
            margin-bottom: 10px;
        }
        .header p {
            color: #666;
            font-size: 16px;
        }
        .upload-section {
            border: 2px dashed #ddd;
            padding: 40px;
            text-align: center;
            border-radius: 10px;
            margin-bottom: 20px;
            background-color: #fafafa;
        }
        .upload-section:hover {
            border-color: #4CAF50;
            background-color: #f0f8f0;
        }
        input[type="file"] {
            margin: 20px 0;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 5px;
            background: white;
        }
        button {
            background-color: #4CAF50;
            color: white;
            padding: 12px 30px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
            margin: 10px;
        }
        button:hover {
            background-color: #45a049;
        }
        button:disabled {
            background-color: #cccccc;
            cursor: not-allowed;
        }
        .requirements {
            background-color: #e7f4fd;
            padding: 20px;
            border-radius: 5px;
            border-left: 4px solid #2196F3;
        }
        .requirements h3 {
            color: #1976D2;
            margin-top: 0;
        }
        .requirements ul {
            margin: 10px 0;
            padding-left: 20px;
        }
        .alert {
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
        }
        .alert-success {
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }
        .alert-error {
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }
        .loading {
            display: none;
            text-align: center;
            padding: 20px;
        }
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #3498db;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 2s linear infinite;
            margin: 0 auto;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Excel Assignment-Memo Matcher</h1>
            <p>Upload your Excel file to find matches between Assignment and Memo Line columns</p>
        </div>
        
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ 'error' if category == 'error' else 'success' }}">
                        {{ message|safe }}
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}
        
        <form action="/upload" method="post" enctype="multipart/form-data" id="uploadForm">
            <div class="upload-section">
                <h3>📁 Select Excel File</h3>
                <p>Choose a .xlsx file containing 'Assignment' and 'Memo Line' columns</p>
                <input type="file" name="file" accept=".xlsx,.xls" required id="fileInput">
                <br>
                <button type="submit" id="uploadBtn">🚀 Process File</button>
            </div>
        </form>
        
        <div class="loading" id="loading">
            <div class="spinner"></div>
            <p>Processing your file... Please wait.</p>
        </div>
        
        <div class="requirements">
            <h3>📋 Requirements</h3>
            <ul>
                <li><strong>File Format:</strong> Excel (.xlsx or .xls)</li>
                <li><strong>Required Columns:</strong> 'Assignment' and 'Memo Line'</li>
                <li><strong>Max File Size:</strong> 16MB</li>
                <li><strong>Processing:</strong> Each Assignment value will be checked against all Memo Line entries</li>
                <li><strong>Output:</strong> Filtered Excel file with matched pairs + insights summary</li>
            </ul>
        </div>
    </div>
    
    <script>
        document.getElementById('uploadForm').addEventListener('submit', function() {
            document.getElementById('loading').style.display = 'block';
            document.getElementById('uploadBtn').disabled = true;
        });
        
        document.getElementById('fileInput').addEventListener('change', function() {
            const file = this.files[0];
            if (file) {
                document.getElementById('uploadBtn').textContent = `🚀 Process "${file.name}"`;
            }
        });
    </script>
</body>
</html>
    """)

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing."""
    try:
        logger.info("=== NEW FILE UPLOAD STARTED ===")
        
        # Check if file was uploaded
        if 'file' not in request.files:
            flash('No file selected', 'error')
            return redirect(url_for('index'))
        
        file = request.files['file']
        if file.filename == '':
            flash('No file selected', 'error')
            return redirect(url_for('index'))
        
        if not allowed_file(file.filename):
            flash('Invalid file type. Please upload an Excel file (.xlsx or .xls)', 'error')
            return redirect(url_for('index'))
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_filename = f"{timestamp}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(filepath)
        logger.info(f"File uploaded: {unique_filename}")
        
        # Read Excel file
        try:
            df = pd.read_excel(filepath)
            logger.info(f"Excel file read successfully. Shape: {df.shape}")
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            flash(f'Error reading Excel file: {str(e)}', 'error')
            return redirect(url_for('index'))
        
        # Validate file structure
        is_valid, validation_message = validate_excel_structure(df)
        if not is_valid:
            logger.error(f"File validation failed: {validation_message}")
            flash(validation_message, 'error')
            return redirect(url_for('index'))
        
        logger.info("File validation passed")
        
        # Process the data
        matches = find_assignment_matches(df)
        
        if not matches:
            flash('No matches found between Assignment and Memo Line columns.', 'error')
            logger.warning("No matches found in the uploaded file")
            return redirect(url_for('index'))
        
        # Create filtered output
        filtered_df = create_filtered_output(df, matches)
        
        # Save output file
        output_filename = f"filtered_output_{timestamp}.xlsx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, sheet_name='Matched_Pairs', index=False)
            
            # Create summary sheet
            summary_data = {
                'Metric': [
                    'Total Input Rows',
                    'Total Matches Found',
                    'Output Rows (Pairs)',
                    'Processing Date',
                    'Original File'
                ],
                'Value': [
                    len(df),
                    len(matches),
                    len(filtered_df),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    filename
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        logger.info(f"Output file saved: {output_filename}")
        
        # Generate insights
        df_info = {
            'total_rows': len(df),
            'columns': df.columns.tolist(),
            'filename': filename
        }
        insights = generate_insights(matches, df_info)
        
        # Clean up uploaded file
        try:
            os.remove(filepath)
            logger.info("Temporary upload file cleaned up")
        except Exception as e:
            logger.warning(f"Could not clean up upload file: {str(e)}")
        
        # Success response with insights
        success_message = f"""
        ✅ <strong>Processing Complete!</strong><br>
        🔍 Found {len(matches)} matches<br>
        📄 Created {len(filtered_df)} rows of paired data<br>
        📥 <a href="/download/{output_filename}" style="color: #4CAF50; text-decoration: none;"><strong>Download Results</strong></a>
        """
        flash(success_message, 'success')
        
        return render_template_string("""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Processing Results</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 900px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
            line-height: 1.6;
        }
        .container {
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .success-header {
            text-align: center;
            color: #4CAF50;
            margin-bottom: 30px;
        }
        .insights {
            background-color: #f8f9fa;
            border-left: 4px solid #4CAF50;
            padding: 20px;
            margin: 20px 0;
            border-radius: 5px;
            font-family: 'Courier New', monospace;
            white-space: pre-line;
        }
        .download-section {
            text-align: center;
            margin: 30px 0;
            padding: 20px;
            background-color: #e7f4fd;
            border-radius: 10px;
        }
        .btn {
            display: inline-block;
            padding: 12px 30px;
            margin: 10px;
            text-decoration: none;
            border-radius: 5px;
            font-weight: bold;
            transition: all 0.3s;
        }
        .btn-primary {
            background-color: #4CAF50;
            color: white;
        }
        .btn-primary:hover {
            background-color: #45a049;
            transform: translateY(-2px);
        }
        .btn-secondary {
            background-color: #6c757d;
            color: white;
        }
        .btn-secondary:hover {
            background-color: #5a6268;
        }
        .alert {
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
        }
        .alert-success {
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="success-header">
            <h1>🎉 Processing Complete!</h1>
        </div>
        
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ 'error' if category == 'error' else 'success' }}">
                        {{ message|safe }}
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}
        
        <div class="download-section">
            <h3>📥 Your Results Are Ready</h3>
            <a href="/download/{{ output_filename }}" class="btn btn-primary">
                📊 Download Excel File
            </a>
            <p style="margin-top: 15px; color: #666;">
                The Excel file contains two sheets:<br>
                • <strong>Matched_Pairs:</strong> All matched Assignment-Memo pairs<br>
                • <strong>Summary:</strong> Processing statistics
            </p>
        </div>
        
        <div class="insights">{{ insights }}</div>
        
        <div style="text-align: center; margin-top: 30px;">
            <a href="/" class="btn btn-secondary">🔄 Process Another File</a>
        </div>
    </div>
</body>
</html>
        """, output_filename=output_filename, insights=insights)
        
    except Exception as e:
        error_msg = f"An unexpected error occurred: {str(e)}"
        logger.error(f"Upload processing error: {error_msg}")
        logger.error(traceback.format_exc())
        flash(error_msg, 'error')
        return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    """Handle file download."""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        
        if not os.path.exists(file_path):
            flash('File not found or has expired', 'error')
            logger.warning(f"Download requested for non-existent file: {filename}")
            return redirect(url_for('index'))
        
        logger.info(f"File downloaded: {filename}")
        return send_file(file_path, as_attachment=True)
        
    except Exception as e:
        error_msg = f"Error downloading file: {str(e)}"
        logger.error(error_msg)
        flash(error_msg, 'error')
        return redirect(url_for('index'))

@app.route('/health')
def health_check():
    """Health check endpoint."""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat(),
        'version': '1.0.0'
    })

@app.errorhandler(413)
def too_large(e):
    flash("File is too large. Maximum size is 16MB.", 'error')
    return redirect(url_for('index'))

@app.errorhandler(500)
def internal_error(error):
    logger.error(f"Internal server error: {str(error)}")
    flash("An internal error occurred. Please try again.", 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    import webbrowser
    import threading
    import time
    
    logger.info("=== EXCEL ASSIGNMENT-MEMO MATCHER STARTED ===")
    logger.info("Application initialized successfully")
    
    def open_browser():
        """Open browser after a short delay to ensure server is running."""
        time.sleep(1.5)
        url = "http://localhost:5000"
        logger.info(f"Opening browser at {url}")
        try:
            webbrowser.open(url)
        except Exception as e:
            logger.warning(f"Could not automatically open browser: {e}")
            print(f"Please manually open your browser and go to: {url}")
    
    # Start browser opening in a separate thread
    threading.Thread(target=open_browser, daemon=True).start()
    
    print("\n" + "="*60)
    print("🚀 EXCEL ASSIGNMENT-MEMO MATCHER")
    print("="*60)
    print("✅ Server starting...")
    print("🌐 Application will open automatically in your browser")
    print("📍 Manual URL: http://localhost:5000")
    print("🔧 Press Ctrl+C to stop the server")
    print("="*60 + "\n")
    
    # Run the application
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=False)