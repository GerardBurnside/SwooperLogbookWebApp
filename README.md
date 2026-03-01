# Skydiving Logbook Web App

A personal mobile-friendly web application for tracking skydiving jumps. Data is stored in Google Spreadsheet for easy access and backup.

This app was conceived for swoopers to keep track on the number of jumps on their linesets.

## Features
- 📱 Mobile-responsive design
- 🪂 Jump entry with location and equipment tracking
- 📊 Auto-incrementing jump numbers (configurable starting number)
- 🔧 **Advanced equipment system** - Manage harnesses, canopies, and linesets separately
- ⚙️ **Equipment rigs** - Combine components into complete setups
- 📊 **Component-level statistics** - Track usage by harness, canopy, and lineset
- 💫 **Equipment archiving** - Archive old rigs while preserving statistics
- 📍 **Drop zone statistics** - Track your most visited locations
- �📋 Google Sheets integration for data storage
- 🔄 Offline-capable with sync when online
- 📤 Export data to CSV
- ⚙️ Configurable settings

## Quick Start
1. Open `index.html` in a web browser
2. **Set up equipment components**: 
   - Go to Equipment tab → Harnesses to add your harness(es) (e.g., Javelin, Mutant)
   - Go to Canopies to add your canopies (e.g., Petra64, Petra68)
   - Go to Linesets to add your linesets (e.g., #1, #2, #3)
3. **Create equipment rigs**:
   - Go to Rigs to combine your components into complete setups
   - Name each rig (e.g., "Main Setup", "Backup Rig")
4. **Log jumps**: Return to Jumps tab and select from your equipment rigs
5. **View detailed stats**: Check Statistics tab for component-level analytics
6. Optional: Configure Google Sheets integration for cloud backup

## Equipment Management

The app uses a realistic three-component equipment system:

### **Components**
- **Harnesses**: Your harness/container systems (e.g., Javelin Odyssey, Mutant)
- **Canopies**: Your parachutes (e.g., Petra64, Petra68, Sabre2 120)  
- **Linesets**: Your line sets, numbered or named (e.g., #1, #2, #3, A-lines)

### **Rigs**
- Combine any harness + canopy + lineset into a complete equipment setup
- Name your rigs (e.g., "Primary Setup", "Backup Rig", "Student Gear")
- Use these rigs when logging jumps

### **Archiving**
- Archive old equipment rigs when you stop using them
- Archived equipment won't appear in the jump logging dropdown
- All statistics are preserved for archived equipment

### **Statistics**
- **Equipment Rigs**: See usage for each complete setup
- **Component Level**: Track individual harness, canopy, and lineset usage
- **Drop Zones**: Your most visited locations

This system gives you detailed insights into your gear usage patterns and helps track equipment lifecycles realistically.

## Google Sheets Setup (Optional)
1. See `config/README.md` for detailed setup instructions
2. Copy `config/sheets-config.example.json` to `config/sheets-config.json`
3. Add your Google Sheets API credentials

## Free Hosting
This app can be hosted for free on:
- **GitHub Pages** (recommended)
- Netlify
- Vercel  
- Firebase Hosting

See `HOSTING.md` for detailed deployment instructions.

## Mobile Usage
Add to your phone's home screen:
1. Open the app in your mobile browser
2. Use "Add to Home Screen" option
3. Enjoy native app-like experience

## Data Storage
- **Local:** Data automatically saved to browser storage
- **Cloud:** Optional Google Sheets sync for backup
- **Export:** Download your data as CSV anytime

## Stack
- Frontend: HTML, CSS, JavaScript (Vanilla)
- Data Storage: LocalStorage + Google Sheets API
- PWA: Manifest + Service Worker for offline use

## Project Structure
```
├── index.html              # Main application
├── css/
│   └── style.css          # Responsive styling
├── js/
│   ├── app.js            # Main application logic
│   └── sheets.js         # Google Sheets integration
├── config/
│   ├── README.md         # Sheets setup guide
│   └── sheets-config.example.json
├── manifest.json         # PWA manifest
├── sw.js                # Service worker
├── HOSTING.md           # Deployment guide
└── README.md           # This file
```

## Contributing
This is a personal project, but feel free to fork and customize for your needs!

## License
Open source - use freely for personal projects.
