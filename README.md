# **Enrollment Funnel Analysis Report - Complete Guide**

---

## **📋 TABLE OF CONTENTS**
1. [One-Time Setup](#one-time-setup) (15 minutes)
2. [Every Time You Generate a Report](#generating-reports) (5 minutes)
3. [Publishing Your Report](#publishing-report) (2 minutes)
4. [Troubleshooting](#troubleshooting)

---

## **ONE-TIME SETUP** <a name="one-time-setup"></a>

### **PART 1: Create Your Google Sheet (5 minutes)**

1. **Open Google Sheets**
   - Go to: https://sheets.google.com
   - Click **"+ Blank"** (green button with plus sign)

2. **Set Up Column Headers**
   - Click cell **A1** and type: `Year`
   - Click cell **B1** and type: `Weeks_Out`
   - Click cell **C1** and type: `Interests`
   - Click cell **D1** and type: `Abandons`
   - Click cell **E1** and type: `Submits`
   - Click cell **F1** and type: `Admits`
   - Click cell **G1** and type: `Expect`

3. **Make Headers Bold** (optional but recommended)
   - Select row 1 (click the "1" on the left)
   - Click the **Bold** button (B icon)

4. **Enter Your Data**
   - Starting in row 2, enter your data
   - Example:
   ```
   Row 2: 2023 | 40 | 283 | 1 | 1 | 0 | 0
   Row 3: 2023 | 39 | 293 | 9 | 3 | 0 | 0
   Row 4: 2023 | 38 | 355 | 198 | 12 | 0 | 0
   ```
   - Continue for all weeks and all years

5. **Rename Your Sheet**
   - Click on "Untitled spreadsheet" at the top
   - Type: `Enrollment Data`
   - Press Enter

6. **Share Your Sheet**
   - Click the blue **"Share"** button (top right)
   - Under "General access" click **"Restricted"**
   - Change to **"Anyone with the link"** → **"Viewer"**
   - Click **"Done"**

7. **Copy the Share Link**
   - Click **"Share"** button again
   - Click **"Copy link"**
   - **SAVE THIS LINK** - you'll need it in the next step

---

### **PART 2: Create Your Google Colab Notebook (10 minutes)**

1. **Open Google Colab**
   - Go to: https://colab.research.google.com
   - Sign in with your Google account

2. **Create New Notebook**
   - Click **"File"** → **"New notebook"**
   - Or click **"+ New notebook"** if you see it

3. **Rename Your Notebook**
   - Click "Untitled0.ipynb" at the top
   - Type: `Enrollment Report Generator`
   - Press Enter

4. **Delete the Default Cell** (if any code is there)
   - Click the cell
   - Click the trash icon on the right

5. **Copy and Paste the Code**
   - Copy ALL of the code below
   - Click in the empty cell in Colab
   - Paste (Ctrl+V or Cmd+V)

```python
# ============================================
# ENROLLMENT FUNNEL ANALYSIS - GOOGLE COLAB
# ============================================

# Run this cell first (click play button)
!pip install plotly -q

import pandas as pd
import plotly.graph_objects as go
from google.colab import files
from datetime import datetime
import json

# ============================================
# CONFIGURATION - EDIT THESE VALUES
# ============================================

# Paste your Google Sheets share link here (between the quotes)
SHEET_URL = "PASTE_YOUR_GOOGLE_SHEETS_LINK_HERE"

# Current year (the year you're analyzing)
CURRENT_YEAR = 2027

# Weeks out to compare (change this each time you run the report)
WEEKS_OUT = 23

# ============================================
# FUNCTIONS (DON'T EDIT BELOW THIS LINE)
# ============================================

def load_google_sheet(url):
    """Load data from Google Sheets"""
    try:
        sheet_id = url.split('/d/')[1].split('/')[0]
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        df = pd.read_csv(csv_url)
        print(f"✓ Loaded {len(df)} rows from Google Sheets")
        print(f"✓ Years found: {sorted(df['Year'].unique())}")
        print(f"✓ Columns: {', '.join(df.columns)}")
        return df
    except Exception as e:
        print(f"✗ Error loading Google Sheet: {e}")
        print("\nMake sure:")
        print("  1. The sheet is shared as 'Anyone with the link'")
        print("  2. The link is correct")
        print("  3. Column names match exactly: Year, Weeks_Out, Interests, Abandons, Submits, Admits, Expect")
        return None

def get_historical_average(df, weeks_out, stage, current_year):
    """Calculate historical average"""
    historical = df[(df['Weeks_Out'] == weeks_out) & (df['Year'] < current_year)]
    return historical[stage].mean() if len(historical) > 0 else 0

def get_color(percentage):
    """Determine color based on percentage"""
    if percentage >= 90:
        return 'Green', '#99ff99'
    elif percentage >= 60:
        return 'Yellow', '#ffff99'
    else:
        return 'Red', '#ff9999'

def generate_comparison_table(df, weeks_out, current_year):
    """Generate comparison table"""
    stages = ['Interests', 'Abandons', 'Submits', 'Admits', 'Expect']
    current_data = df[(df['Year'] == current_year) & (df['Weeks_Out'] == weeks_out)]
    
    print(f"\n{'='*60}")
    print(f"COMPARISON TABLE - {current_year} at {weeks_out} Weeks Out")
    print(f"{'='*60}")
    
    rows = []
    for stage in stages:
        current_val = current_data[stage].values[0] if len(current_data) > 0 else 0
        hist_avg = get_historical_average(df, weeks_out, stage, current_year)
        percentage = (current_val / hist_avg * 100) if hist_avg > 0 else 0
        color_name, color_hex = get_color(percentage)
        
        rows.append({
            'stage': stage,
            'current': int(current_val),
            'avg': round(hist_avg, 2),
            'pct': round(percentage, 1),
            'color': color_name,
            'color_hex': color_hex
        })
        
        print(f"{stage:12} | {current_year}: {int(current_val):6} | Avg: {hist_avg:7.2f} | {percentage:5.1f}% | {color_name}")
    
    print(f"{'='*60}\n")
    return rows

def create_traces(df, current_year, stage, color):
    """Create Plotly traces for a stage"""
    years = sorted(df['Year'].unique())
    traces = []
    positions = ['top right', 'bottom right', 'middle right', 'top center']
    
    # Historical years
    for idx, year in enumerate([y for y in years if y < current_year]):
        year_data = df[df['Year'] == year].sort_values('Weeks_Out', ascending=False)
        text_list = [''] * (len(year_data) - 1) + [f'{year} {stage}']
        
        traces.append({
            'x': year_data['Weeks_Out'].tolist(),
            'y': year_data[stage].tolist(),
            'mode': 'lines+markers+text',
            'name': f'{stage} {year}',
            'line': {'color': color, 'width': 2},
            'marker': {'size': 8, 'symbol': 'circle'},
            'text': text_list,
            'textposition': positions[idx % len(positions)],
            'textfont': {'color': color, 'size': 12},
            'hovertemplate': '<b>%{text}</b><br>Weeks Out: %{x}<br>Value: %{y}<extra></extra>'
        })
    
    # Current year
    current_data = df[df['Year'] == current_year].sort_values('Weeks_Out', ascending=False)
    if len(current_data) > 0:
        text_list = [''] * (len(current_data) - 1) + [f'{current_year} {stage}']
        traces.append({
            'x': current_data['Weeks_Out'].tolist(),
            'y': current_data[stage].tolist(),
            'mode': 'lines+markers+text',
            'name': f'{stage} {current_year}',
            'line': {'color': color, 'width': 4},
            'marker': {'size': 12, 'symbol': 'star'},
            'text': text_list,
            'textposition': 'bottom center',
            'textfont': {'color': color, 'size': 12},
            'hovertemplate': '<b>%{text}</b><br>Weeks Out: %{x}<br>Value: %{y}<extra></extra>'
        })
    
    return traces

def generate_html_report(df, weeks_out, current_year):
    """Generate complete HTML report"""
    
    comparison_rows = generate_comparison_table(df, weeks_out, current_year)
    final_expect_avg = int(get_historical_average(df, 1, 'Expect', current_year))
    
    colors = {
        'Submits': '#0074D9',
        'Admits': '#2ECC40',
        'Expect': '#FF851B',
        'Interests': '#8e44ad',
        'Abandons': '#e74c3c'
    }
    
    # Create traces
    traces_chart1 = (create_traces(df, current_year, 'Submits', colors['Submits']) +
                     create_traces(df, current_year, 'Admits', colors['Admits']) +
                     create_traces(df, current_year, 'Expect', colors['Expect']))
    
    traces_chart2 = (create_traces(df, current_year, 'Interests', colors['Interests']) +
                     create_traces(df, current_year, 'Abandons', colors['Abandons']))
    
    html = f"""<html>
<head>
    <meta charset="UTF-8">
    <title>Enrollment Funnel Analysis - Week {weeks_out}</title>
</head>
<body>
    <h1 style="text-align:center;">Enrollment Funnel Analysis ({int(df['Year'].min())}–{current_year})</h1>
    
<h2 style="text-align:center">{current_year} Performance at {weeks_out} Weeks Out vs Historical Average</h2>
<table style="margin-left:auto;margin-right:auto;border-collapse:collapse;font-size:1.1em">
    <tr>
        <th style="border:1px solid #888;padding:6px">Weeks Out</th>
        <th style="border:1px solid #888;padding:6px">Stage</th>
        <th style="border:1px solid #888;padding:6px">{current_year} Value</th>
        <th style="border:1px solid #888;padding:6px">Historical Avg</th>
        <th style="border:1px solid #888;padding:6px">{current_year} vs Avg (%)</th>
        <th style="border:1px solid #888;padding:6px">Status</th>
    </tr>
"""
    
    for row in comparison_rows:
        html += f"""
    <tr style="background-color:{row['color_hex']}">
        <td style="border:1px solid #888;padding:6px;text-align:center">{weeks_out}</td>
        <td style="border:1px solid #888;padding:6px;text-align:center">{row['stage']}</td>
        <td style="border:1px solid #888;padding:6px;text-align:center">{row['current']}</td>
        <td style="border:1px solid #888;padding:6px;text-align:center">{row['avg']}</td>
        <td style="border:1px solid #888;padding:6px;text-align:center">{row['pct']}</td>
        <td style="border:1px solid #888;padding:6px;text-align:center">{row['color']}</td>
    </tr>
"""
    
    html += f"""</table><br>
<div style="text-align:center; font-size:1.2em; margin-bottom:1em;">
    <strong>Note:</strong> The historical average final expected number is: <b>{final_expect_avg}</b>
</div>
    <br>
    <div>
        <script src="https://cdn.plot.ly/plotly-2.4.1.min.js"></script>
        <div id="chart1" style="height:600px; width:100%;"></div>
        <script>
            Plotly.newPlot("chart1", {json.dumps(traces_chart1)}, {{
                "title": {{"text": "Submits, Admits & Expect by Weeks Out"}},
                "xaxis": {{"autorange": "reversed", "title": {{"text": "Weeks Out (Countdown to 1)"}}}},
                "yaxis": {{"gridcolor": "#eeeeee", "title": {{"text": "Count"}}}},
                "hovermode": "x unified",
                "plot_bgcolor": "white",
                "legend": {{"title": {{"text": "Stage & Year"}}}},
                "shapes": [{{"line": {{"color": "gray", "dash": "dash", "width": 2}}, "type": "line", "x0": {weeks_out}, "x1": {weeks_out}, "xref": "x", "y0": 0, "y1": 1, "yref": "y domain"}}],
                "annotations": [{{"showarrow": false, "text": "{weeks_out} Weeks Out", "x": {weeks_out}, "xanchor": "right", "xref": "x", "y": 1, "yanchor": "top", "yref": "y domain"}}]
            }}, {{"responsive": true}});
        </script>
    </div>

    <div>
        <div id="chart2" style="height:600px; width:100%;"></div>
        <script>
            Plotly.newPlot("chart2", {json.dumps(traces_chart2)}, {{
                "title": {{"text": "Interests & Abandons by Weeks Out"}},
                "xaxis": {{"autorange": "reversed", "title": {{"text": "Weeks Out (Countdown to 1)"}}}},
                "yaxis": {{"gridcolor": "#eeeeee", "title": {{"text": "Count"}}}},
                "hovermode": "x unified",
                "plot_bgcolor": "white",
                "legend": {{"title": {{"text": "Stage & Year"}}}},
                "shapes": [{{"line": {{"color": "gray", "dash": "dash", "width": 2}}, "type": "line", "x0": {weeks_out}, "x1": {weeks_out}, "xref": "x", "y0": 0, "y1": 1, "yref": "y domain"}}],
                "annotations": [{{"showarrow": false, "text": "{weeks_out} Weeks Out", "x": {weeks_out}, "xanchor": "right", "xref": "x", "y": 1, "yanchor": "top", "yref": "y domain"}}]
            }}, {{"responsive": true}});
        </script>
    </div>
<div style="text-align:center; margin-top:2em; color:#888; font-size:0.9em;">
    Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}
</div>
</body>
</html>"""
    
    return html

# ============================================
# MAIN EXECUTION
# ============================================

print("=" * 60)
print("  ENROLLMENT FUNNEL ANALYSIS REPORT GENERATOR")
print("=" * 60)
print()

# Load data
print("Loading data from Google Sheets...")
df = load_google_sheet(SHEET_URL)

if df is not None:
    # Generate report
    print(f"\nGenerating report for {CURRENT_YEAR} at {WEEKS_OUT} weeks out...")
    html_content = generate_html_report(df, WEEKS_OUT, CURRENT_YEAR)
    
    # Save and download
    filename = f"enrollment_report_week{WEEKS_OUT}_{datetime.now().strftime('%Y%m%d')}.html"
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"\n✓ Report generated successfully!")
    print(f"✓ Filename: {filename}")
    print("\nDownloading file to your computer...")
    files.download(filename)
    print("\n" + "=" * 60)
    print("✓ DONE! Open the downloaded HTML file in your web browser.")
    print("=" * 60)
else:
    print("\n✗ Could not generate report due to data loading error.")
```

6. **Edit the Configuration Section**
   - Find this line near the top:
     ```python
     SHEET_URL = "PASTE_YOUR_GOOGLE_SHEETS_LINK_HERE"
     ```
   - Delete `PASTE_YOUR_GOOGLE_SHEETS_LINK_HERE`
   - Paste your Google Sheets link (the one you copied earlier)
   - Make sure it stays between the quotes

   - Find these lines:
     ```python
     CURRENT_YEAR = 2027
     WEEKS_OUT = 23
     ```
   - Change `2027` to whatever year you're currently analyzing
   - Change `23` to whatever week comparison you want

7. **Save Your Notebook**
   - Click **"File"** → **"Save"**
   - Or press Ctrl+S (Cmd+S on Mac)

---

## **GENERATING REPORTS** <a name="generating-reports"></a>

### **Every Time You Want a New Report (5 minutes)**

1. **Update Your Google Sheet**
   - Go to your Google Sheet
   - Add new data or update existing data
   - No need to save - Google Sheets auto-saves

2. **Open Your Colab Notebook**
   - Go to: https://colab.research.google.com
   - Click on your saved notebook: "Enrollment Report Generator"

3. **Update the Week Number** (if needed)
   - Find this line:
     ```python
     WEEKS_OUT = 23
     ```
   - Change the number to the week you want to analyze

4. **Run the Report**
   - Click **"Runtime"** → **"Run all"**
   - Or click the ▶️ play button on the left of the code cell

5. **Wait for Completion**
   - You'll see text appearing below the code
   - Look for: "✓ DONE! Open the downloaded HTML file in your web browser"
   - The file will automatically download to your computer

6. **Open the HTML File**
   - Find the downloaded file (usually in your Downloads folder)
   - Double-click to open in your web browser
   - Review the report!

---

## **PUBLISHING YOUR REPORT** <a name="publishing-report"></a>

### **Share Via GitHub Pages (2 minutes)**

1. **Go to Your GitHub Repository**
   - Open: https://github.com/yourusername/your-repository-name

2. **Upload the HTML File**
   - Click **"Add file"** → **"Upload files"**
   - Drag your HTML file from Downloads
   - Or click "choose your files" and select it

3. **Commit the File**
   - Scroll down
   - Type a commit message like: "Add week 23 report"
   - Click **"Commit changes"**

4. **Get Your Public URL**
   - Your report is now at:
   ```
   https://yourusername.github.io/repository-name/filename.html
   ```
   - Example:
   ```
   https://john-smith.github.io/enrollment-reports/enrollment_report_week23_20250115.html
   ```

5. **Share the Link**
   - Copy the URL
   - Send to anyone who needs to see the report
   - They can view it in any web browser
   - No GitHub account needed to view

---

## **TROUBLESHOOTING** <a name="troubleshooting"></a>

### **"Error loading Google Sheet"**
**Problem:** Can't read your Google Sheet

**Solutions:**
- Make sure the sheet is shared as "Anyone with the link"
- Check that column names are exactly: `Year`, `Weeks_Out`, `Interests`, `Abandons`, `Submits`, `Admits`, `Expect` (case-sensitive!)
- Make sure the link is between quotes: `"https://..."`

---

### **"No data found for 2027 at 23 weeks out"**
**Problem:** Missing data for that specific week

**Solutions:**
- Check your Google Sheet - do you have a row for that year and week?
- Make sure `CURRENT_YEAR` matches a year in your data
- Make sure `WEEKS_OUT` matches weeks in your data

---

### **Report shows all zeros**
**Problem:** Data not loading correctly

**Solutions:**
- Check column names match exactly (case-sensitive)
- Make sure data starts in row 2 (row 1 is headers)
- No empty columns between data

---

### **GitHub Pages shows 404**
**Problem:** Can't see the published report

**Solutions:**
- Wait 2-3 minutes after uploading
- Check that GitHub Pages is enabled (Settings → Pages → Source = main)
- Make sure repository is Public (not Private)
- Check exact filename (case-sensitive)

---

### **Download doesn't start in Colab**
**Problem:** File won't download

**Solutions:**
- Check if your browser blocked the download (look for popup blocker message)
- Make sure pop-ups are allowed for colab.research.google.com
- Try a different browser (Chrome works best)

---

### **"ModuleNotFoundError" or package errors**
**Problem:** Missing Python packages

**Solutions:**
- Make sure the first line `!pip install plotly -q` ran successfully
- Click **"Runtime"** → **"Restart runtime"**
- Run the code again

---

## **QUICK REFERENCE CHECKLIST**

### **One-Time Setup ✓**
- [ ] Create Google Sheet with correct column names
- [ ] Share sheet as "Anyone with the link"
- [ ] Copy share link
- [ ] Create Google Colab notebook
- [ ] Paste code into Colab
- [ ] Update `SHEET_URL` with your link
- [ ] Save notebook
- [ ] Create GitHub repository
- [ ] Enable GitHub Pages

### **Each Time You Generate a Report ✓**
- [ ] Update data in Google Sheet
- [ ] Open Colab notebook
- [ ] Update `WEEKS_OUT` number (if needed)
- [ ] Click "Runtime" → "Run all"
- [ ] Download HTML file
- [ ] Upload to GitHub repository
- [ ] Share the public URL

---

## **TIPS & BEST PRACTICES**

1. **Naming Files:** Use descriptive names like `report_week23_2027.html`

2. **Keep History:** Don't delete old reports from GitHub - they create a historical record

3. **Regular Updates:** Update your Google Sheet as soon as new data comes in

4. **Test First:** After setup, generate a test report to make sure everything works

5. **Bookmark Links:**
   - Your Google Sheet
   - Your Colab notebook
   - Your GitHub repository

6. **Data Entry:** Always enter data in the same format (no commas in numbers, use 0 for missing data)

---

**Questions or issues?** Check the Troubleshooting section above, or verify each step carefully.

**Need to change settings?** Just edit the configuration section at the top of your Colab notebook and re-run.
