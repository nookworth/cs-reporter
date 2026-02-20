# CS Reporter Web UI - Quick Start

## 🚀 Launch the Web Interface

### Windows:
```cmd
run_web.bat
```

### macOS/Linux:
```bash
chmod +x run_web.sh
./run_web.sh
```

### Or directly:
```bash
streamlit run app.py
```

The web interface will automatically open in your browser at `http://localhost:8501`

---

## 📖 How to Use

### Step 1: Upload Files
1. Click "Browse files" under **Current Month**
2. Select your current month Excel file
3. Click "Browse files" under **Previous Month**
4. Select your previous month Excel file

### Step 2: Preview Data
- Expand "Preview Data" to see your Excel sheets
- Verify the correct sheets are loaded

### Step 3: Configure (Optional)
- Use sidebar to select V1 or V2
- Choose config file (demo or production)

### Step 4: Generate Report
1. Click **"🚀 Generate Report"** button
2. Wait for processing (usually 5-10 seconds)
3. Expand "📊 Extracted Data Preview" to verify data

### Step 5: Download
- Click **"📥 Download PowerPoint Report"**
- Save the .pptx file to your computer

---

## ✨ Features

### Drag & Drop
- Simply drag Excel files onto the upload areas
- Supports .xlsx and .xls formats

### Live Preview
- See your Excel data before processing
- Preview extracted metrics and tables
- Verify data accuracy

### Instant Download
- No need to find output folder
- Direct download from browser
- Timestamped filenames

### Version Selection
- Switch between V1 and V2
- Choose different config files
- Test configurations easily

---

## 🔧 Configuration

### Using Custom Configs
1. Place your config in `config/` folder
2. Restart the web interface
3. Select from dropdown in sidebar

### Switching Versions
- **V2 (Recommended)**: Operation-based, with filters
- **V1 (Legacy)**: Original suffix-based system

---

## 🐛 Troubleshooting

### Port Already in Use
If you see "Address already in use":
```bash
streamlit run app.py --server.port 8502
```

### Module Not Found
Install dependencies:
```bash
pip install -r requirements.txt
```

### File Upload Fails
- Check file size (max 200MB by default)
- Ensure file is not corrupted
- Try closing and reopening Excel file

### Report Generation Fails
- Check console for error messages
- Verify Excel sheet names match config
- Ensure both files have required columns

---

## 🌐 Remote Access

### Share with Team
```bash
streamlit run app.py --server.address 0.0.0.0
```
Then share: `http://YOUR_IP:8501`

### Deploy to Cloud
- **Streamlit Cloud**: Free hosting
- **Heroku**: Easy deployment
- **AWS/Azure**: Full control

See `BT-Docs/deployment.md` for details.

---

## 💡 Tips

1. **Bookmark the URL** for quick access
2. **Keep browser tab open** while generating
3. **Use demo config** for testing
4. **Check preview** before generating
5. **Download immediately** - files are temporary

---

## 🆚 Web UI vs Command Line

| Feature | Web UI | Command Line |
|---------|--------|--------------|
| Ease of use | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ |
| File selection | Drag & drop | File dialog |
| Data preview | ✅ Yes | ❌ No |
| Download | Direct | Find in folder |
| Automation | ❌ No | ✅ Yes |
| Remote access | ✅ Yes | ❌ No |

**Use Web UI for:** Manual monthly reports, sharing with team, non-technical users

**Use Command Line for:** Automation, scripting, scheduled tasks

---

## 📚 Next Steps

- Customize your PowerPoint template
- Create custom config files
- Set up scheduled automation
- Deploy to cloud for team access

See `README.md` for full documentation.
