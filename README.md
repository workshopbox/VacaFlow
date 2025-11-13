# DAP8 Vacation Analytics Dashboard

## 📊 Complete Absence Management System

Track and analyze all 7 absence types from your Amazon DAP8 logistics calendar.

---

## 🚀 Quick Start

1. **Open `index.html` in your web browser**
2. **Upload your calendar HTML file** (exported from Excel)
3. **View comprehensive analytics** across all absence types

That's it! No installation or server required.

---

## 📁 Project Structure

```
dap8_dashboard/
├── index.html       # Main HTML structure
├── style.css        # All styling and layout
├── script.js        # Dashboard logic and charts
└── README.md        # This file
```

---

## 🎯 Features

### Tracked Day Types

| Icon | Type | Color |
|------|------|-------|
| 🏖️ | Vacation | Lime |
| ⏰ | Flextime Day | Purple |
| 👶 | Papamonat | Magenta |
| 📋 | Special Leave | Cyan |
| 🏥 | Long Sick Leave | Red |
| 🤒 | Short Sick Leave | Brown |
| ✈️ | Active Travel Time | Light Cyan |

### Dashboard Components

**Stats Cards**
- Total employees
- Total absence days
- Breakdown by each type (with icons)

**4 Interactive Charts**
1. Absence Types Distribution (Doughnut)
2. Absence Days by Shift (Stacked Bar)
3. Daily Absence Coverage (Stacked Area)
4. Weekly Distribution (Stacked Bar)

**Employee Table**
- Detailed breakdown per employee
- Color-coded columns for each type
- Sortable and searchable
- CSV export functionality

**Smart Filters**
- Filter by Department (DAP8, DAP9, etc.)
- Filter by Shift (Früh, Nacht, Hybrid, Spät, ORM)
- Filter by Job Level (JL1, JL3, JL4, etc.)

---

## 📋 How to Use

### 1. Prepare Your Calendar File

Export your team calendar from Excel as HTML:
- File → Save As → Web Page (.htm or .html)
- Keep the stylesheet.css in the same folder (or use embedded styles)

### 2. Upload the File

- Drag & drop onto the upload area, OR
- Click "Choose File" to browse

### 3. Analyze the Data

The dashboard automatically:
- Detects all 7 day types based on cell colors
- Separates Department, Shift, and Job Level
- Generates charts and statistics
- Populates the employee table

### 4. Filter and Export

- Use dropdowns to filter data
- Search for specific employees
- Export filtered data to CSV

---

## 🎨 Color Coding

The application uses your Excel stylesheet colors:

- **xl68 (lime)** → Vacation
- **xl70 (purple)** → Flextime Day
- **xl71 (magenta)** → Papamonat
- **xl72 (cyan)** → Special Leave
- **xl73 (red)** → Long Sick Leave
- **xl74 (brown)** → Short Sick Leave
- **xl75 (light cyan)** → Active Travel Time

---

## 🔧 Technical Details

**Dependencies (loaded from CDN)**
- Chart.js 4.4.0 - For all charts and graphs
- XLSX.js 0.18.5 - For Excel file processing

**Browser Compatibility**
- Chrome/Edge (recommended)
- Firefox
- Safari
- Modern browsers with ES6 support

**File Format**
- Accepts .htm and .html files
- Must be Excel-exported HTML with color classes
- Supports standard Excel HTML format

---

## 📊 Data Structure

### Expected HTML Format

```html
<table>
  <tr>
    <td>Department Column</td>
    <td>Name Column</td>
    <td class="xl68">Day 1</td>
    <td class="xl67">Day 2</td>
    ...
  </tr>
</table>
```

### Department Column Format

Format: `DAP8 JL1 Früh (DAP8 JL1 Früh)`

Extracts:
- **Department**: First 4 chars (DAP8)
- **Job Level**: JL followed by number (JL1)
- **Shift**: Früh, Nacht, Hybrid, Spät, or ORM

### Name Column Format

Format: `Frühschicht 1` or `L3 Früh 2`

Used for shift detection when department column is ambiguous.

---

## 🎯 Use Cases

**HR & Management**
- Monthly absence reports
- Capacity planning
- Sick leave analysis
- Vacation coverage tracking

**Team Leads**
- Daily coverage overview
- Week-by-week planning
- Identify absence patterns
- Resource allocation

**Compliance**
- Papamonat tracking
- Sick leave documentation
- Work-life balance metrics

---

## 💡 Tips

**For Best Results**
- Keep your Excel calendar format consistent
- Use the standard color coding
- Include all employee information in the first two columns
- Export with "Web Page" option (not "Web Page, Filtered")

**Performance**
- Handles 100+ employees easily
- Processes files instantly
- Charts render in real-time
- Filters update immediately

---

## 🐛 Troubleshooting

**No data showing?**
- Check that your file is HTML format (.htm or .html)
- Verify cells are colored using the xl68-xl75 classes
- Make sure the first row contains headers

**Wrong shift/department detected?**
- Verify the department column follows the format: `DAP8 JL1 Früh`
- Check that employee names contain shift information

**Charts not loading?**
- Ensure you have internet connection (for CDN libraries)
- Try a different browser
- Check browser console for errors (F12)

---

## 📄 License

Internal use for Amazon DAP8 Logistics Operations

---

## 👨‍💻 Development

**To modify styles:**
- Edit `style.css`

**To change functionality:**
- Edit `script.js`

**To adjust layout:**
- Edit `index.html`

All files are well-commented for easy customization.

---

## 🎉 Credits

Built for Amazon DAP8 Logistics Operations
Version: 2.0 - All Day Types
Last Updated: November 2025

---

**Need Help?** Check the browser console (F12) for detailed error messages.

**Questions?** Review the code comments in `script.js` for implementation details.
