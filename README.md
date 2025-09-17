# 📊 Excel Assignment-Memo Matcher

A powerful web-based tool for processing Excel files to find matches between Assignment values and Memo Line entries. The application automatically extracts matching pairs and generates a filtered output file with comprehensive logging and AI-powered insights.

## 🚀 Features

### Core Functionality
- **Smart Matching Algorithm**: Compares each Assignment value against all Memo Line entries across rows
- **Paired Output Generation**: Creates filtered Excel files with matched Assignment-Memo pairs
- **Mixed Data Type Support**: Handles alphanumeric strings, numbers, and text seamlessly
- **Deterministic Results**: Consistent, reliable matching with no randomization

### User Experience
- **Drag-and-Drop Interface**: Modern, intuitive file upload experience
- **Real-time Progress Tracking**: Visual progress bar with step-by-step updates
- **Comprehensive Logging**: Detailed app.log with timestamps and status levels
- **AI-Powered Insights**: Intelligent summary generation with match analysis

### Error Handling
- **File Validation**: Checks file format, size, and structure
- **Column Verification**: Ensures required columns exist
- **Graceful Error Recovery**: Clear error messages with recovery suggestions
- **Data Sanitization**: Handles null, undefined, and empty values

## 📋 Requirements

### Input File Format
- **File Type**: Excel files (.xlsx or .xls)
- **Required Columns**:
  - `Assignment`: Contains assignment identifiers/values
  - `Memo Line`: Contains memo text that may reference assignments

### Browser Support
- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

## 🔧 Installation & Setup

### Option 1: Direct Use
1. Save the HTML file as `excel-processor.html`
2. Open in any modern web browser
3. Start processing files immediately

### Option 2: Web Server Deployment
```bash
# Using Python (recommended for local development)
python -m http.server 8000

# Using Node.js
npx serve .

# Using PHP
php -S localhost:8000
```

Then navigate to `http://localhost:8000`

## 📖 Usage Guide

### Step-by-Step Process

1. **File Upload**
   - Click "Choose File" or drag and drop your Excel file
   - Supported formats: .xlsx, .xls
   - Maximum recommended size: 50MB

2. **Processing**
   - Click "Process File" to start the matching algorithm
   - Monitor real-time progress and status updates
   - View detailed logs in the application log section

3. **Results Review**
   - Check the AI-generated summary and insights
   - Review match statistics and performance metrics
   - Analyze data quality recommendations

4. **Download Output**
   - Click "Download Results" to get your `filtered_output.xlsx`
   - File contains matched pairs with metadata columns

### Example Input Data

| Assignment | Memo Line | Other Data |
|------------|-----------|------------|
| INV-001 | Payment for INV-001 received | Customer A |
| INV-002 | Processing order | Customer B |
| INV-003 | INV-003 shipped today | Customer C |

### Example Output Data

| Assignment | Memo Line | Other Data | _MatchGroup | _RowType |
|------------|-----------|------------|-------------|----------|
| INV-001 | Payment for INV-001 received | Customer A | 1 | Assignment |
| INV-001 | Payment for INV-001 received | Customer A | 1 | Memo |
| INV-003 | INV-003 shipped today | Customer C | 2 | Assignment |
| INV-003 | INV-003 shipped today | Customer C | 2 | Memo |

## 🔍 Technical Details

### Architecture
- **Frontend**: Pure HTML5/CSS3/JavaScript (ES6+)
- **Excel Processing**: SheetJS library for robust file handling
- **Data Processing**: Custom matching algorithms with normalization
- **File Generation**: Client-side Excel file creation using Blob API

### Key Components

#### ExcelProcessor Class
```javascript
class ExcelProcessor {
    // Main application controller
    // Handles file processing, matching, and output generation
}
```

#### Matching Algorithm
1. **Data Normalization**: Converts all values to lowercase strings
2. **Cross-Reference Matching**: Compares each assignment against all memo lines
3. **Duplicate Prevention**: Avoids duplicate match pairs
4. **Result Pairing**: Groups matches with original row data

#### Logging System
- **Timestamp Tracking**: All actions logged with precise timestamps
- **Level Classification**: Info, Success, Warning, Error levels
- **Real-time Display**: Live log updates in application interface

### Data Flow
```
Excel File → Validation → Parsing → Normalization → Matching → Output Generation → Download
```

## 🛠️ Configuration Options

### Customizable Parameters
- **Match Sensitivity**: Modify the matching algorithm for partial/exact matches
- **Output Format**: Customize column names and metadata
- **Logging Levels**: Adjust log verbosity
- **File Size Limits**: Configure maximum upload size

### Environment Variables
```javascript
// Configurable constants
const CONFIG = {
    MAX_FILE_SIZE: 50 * 1024 * 1024, // 50MB
    SUPPORTED_FORMATS: ['.xlsx', '.xls'],
    LOG_LEVELS: ['info', 'success', 'warning', 'error']
};
```

## 🔧 Troubleshooting

### Common Issues

#### "Missing Required Columns" Error
- **Cause**: Excel file doesn't contain "Assignment" and/or "Memo Line" columns
- **Solution**: Ensure column headers match exactly (case-sensitive)

#### "No Data Found" Error
- **Cause**: Empty Excel file or all rows are blank
- **Solution**: Verify file contains data in the required columns

#### "File Format Not Supported" Error
- **Cause**: Uploaded file is not .xlsx or .xls format
- **Solution**: Convert file to Excel format or check file extension

#### Browser Memory Issues
- **Cause**: Large files exceeding browser memory limits
- **Solution**: Split large files or use smaller datasets

### Performance Optimization

#### For Large Files (10,000+ rows)
- Process files in smaller batches
- Consider server-side processing for very large datasets
- Monitor browser memory usage

#### Memory Management
- Files are processed entirely in memory
- Automatic cleanup after processing
- No persistent storage used

## 🧪 Testing

### Test Cases
1. **Valid Files**: Standard Excel files with required columns
2. **Invalid Formats**: Non-Excel files, corrupted files
3. **Missing Columns**: Files without required column headers
4. **Empty Files**: Files with no data or empty sheets
5. **Large Files**: Files with thousands of rows
6. **Mixed Data Types**: Files with numbers, text, and special characters

### Sample Test Data
```csv
Assignment,Memo Line,Description
TEST-001,Processing TEST-001 order,Sample order
TEST-002,Customer inquiry,General inquiry
TEST-003,TEST-003 completed successfully,Completed task
```

## 🔮 Future Enhancements

### Planned Features
- [ ] **Batch Processing**: Multiple file processing
- [ ] **Advanced Matching**: Fuzzy matching algorithms
- [ ] **Export Formats**: CSV, JSON, PDF export options
- [ ] **API Integration**: RESTful API for programmatic access
- [ ] **Cloud Storage**: Integration with cloud storage providers
- [ ] **Real AI Integration**: Azure OpenAI or LangGraph integration

### LLM Integration Options
```javascript
// Future AI integration example
const aiInsights = await generateInsights(matches, {
    provider: 'azure-openai',
    model: 'gpt-4',
    temperature: 0.1
});
```

## 📜 License

MIT License - Feel free to use, modify, and distribute.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📞 Support

For issues, questions, or contributions:
- Open an issue on GitHub
- Contact the development team
- Check the troubleshooting guide above

## 📊 Performance Metrics

### Typical Performance
- **Small Files** (< 1,000 rows): < 2 seconds
- **Medium Files** (1,000-10,000 rows): 2-10 seconds
- **Large Files** (10,000+ rows): 10+ seconds

### System Requirements
- **RAM**: Minimum 4GB, Recommended 8GB+
- **Browser**: Modern browser with JavaScript enabled
- **Storage**: Temporary storage for file processing

---

**Made with ❤️ for efficient Excel data processing**