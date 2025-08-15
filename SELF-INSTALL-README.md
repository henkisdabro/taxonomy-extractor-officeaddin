# 🚀 IPG Taxonomy Extractor - Self-Installation Guide

> **Install the add-in yourself without waiting for IT deployment**

Since your IT team hasn't responded to central deployment requests, this guide provides everything you need to install the IPG Taxonomy Extractor add-in individually on both Excel Desktop and Excel Web.

## 📦 What's in This Package

```
📁 Self-Installation Package
├── 🖥️ install-excel-desktop.ps1      # Automated Desktop installation
├── 🌐 install-excel-web-guide.md     # Step-by-step Web installation
├── 📊 installation-comparison.md     # Detailed pros/cons comparison
├── 📋 SELF-INSTALL-README.md         # This file
└── 🔗 manifest.xml                   # Add-in configuration file
```

## ⚡ Quick Start - Choose Your Platform

### 🖥️ Excel Desktop (Windows)
**Best for:** Windows users who want permanent installation

1. **Download**: Save `install-excel-desktop.ps1` to your computer
2. **Run**: Right-click → "Run with PowerShell"
3. **Follow**: The interactive installation wizard
4. **Time**: 2-5 minutes

[**📖 Detailed Desktop Guide →**](install-excel-desktop.ps1)

### 🌐 Excel Web (All Platforms)
**Best for:** Cross-platform users, mobile access, quick setup

1. **Download**: Save `manifest.xml` to your computer
2. **Open**: Excel Web at [office.com](https://office.com)
3. **Upload**: Home → Add-ins → Upload My Add-in
4. **Select**: Choose the downloaded `manifest.xml`
5. **Time**: 1-2 minutes

[**📖 Detailed Web Guide →**](install-excel-web-guide.md)

## 🤔 Which Method Should I Choose?

| Your Situation | Recommended Method | Why |
|------------------|-------------------|-----|
| Windows user, uses Excel daily | 🖥️ **Desktop** | Permanent, faster, offline capable |
| Mac/Linux user | 🌐 **Web** | Only cross-platform option |
| Mobile/tablet user | 🌐 **Web** | Works on mobile browsers |
| Occasional Excel user | 🌐 **Web** | Simpler, no system changes |
| Technical/IT background | 🖥️ **Desktop** | More control and features |
| Non-technical user | 🌐 **Web** | Easier, point-and-click |

[**📊 Detailed Comparison →**](installation-comparison.md)

## 🔧 Prerequisites & Requirements

### For Excel Desktop Installation
- ✅ Windows 7 SP1 or higher
- ✅ Microsoft Excel 2016 or higher
- ✅ PowerShell 5.0+ (pre-installed on Windows 10/11)
- ✅ Internet connection (for download only)

### For Excel Web Installation
- ✅ Modern web browser (Chrome, Edge, Firefox, Safari)
- ✅ Microsoft 365 account
- ✅ Internet connection (always required)
- ✅ Organization allows custom add-ins

## ⚠️ Common Issues & Solutions

### Issue: "Can't find Excel" or "Office not detected"
**Solution:** Ensure Microsoft Office/Excel is properly installed and licensed.

### Issue: "PowerShell execution blocked"
**Solution:** Run this command first:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Issue: "Upload My Add-in button missing" (Excel Web)
**Solution:** Your organization may have disabled custom add-ins. Contact IT or try Excel Desktop method.

### Issue: Add-in appears but doesn't work
**Solution:** 
1. Restart Excel completely
2. Check internet connection
3. Ensure you're using a supported Excel version

## 🆘 Need Help?

### 📚 Documentation
- [Installation Comparison](installation-comparison.md) - Detailed pros/cons
- [Excel Web Guide](install-excel-web-guide.md) - Complete web installation
- [Project Repository](https://github.com/henkisdabro/taxonomy-extractor-officeaddin) - Full documentation

### 🐛 Reporting Issues
1. Visit: [GitHub Issues](https://github.com/henkisdabro/taxonomy-extractor-officeaddin/issues)
2. Click "New Issue"
3. Describe your problem with:
   - Operating system
   - Excel version
   - Installation method attempted
   - Error messages (if any)

### 📧 Direct Support
Create an issue on GitHub with your question - we typically respond within 24 hours.

## 🔒 Security & Privacy

### Is This Safe?
- ✅ **Microsoft Approved**: Uses official Office Add-in framework
- ✅ **Open Source**: All code is publicly available for review
- ✅ **No Data Collection**: Processes your data locally only
- ✅ **HTTPS Only**: All communications are encrypted

### What Permissions Does It Need?
- **Read/Write Document**: To process your Excel data
- **Internet Access**: To load the add-in interface
- **No System Access**: Cannot access files outside Excel

## 🎯 What This Add-in Does

The IPG Taxonomy Extractor helps you:

- ✨ **Extract specific segments** from pipe-delimited taxonomy data
- 🔄 **Process Activation IDs** with post-colon identifier extraction  
- 🎯 **Handle targeting patterns** with `^ABC^` pattern detection
- ↩️ **Undo operations** with multi-step undo system
- ⚡ **Real-time updates** with context-aware interface

**Example Input:**
```
Brand^123^Campaign^456^Creative^789^Platform^ABC^Audience^XYZ: ActivationID_2024
```

**Extracted Results:**
- Segment 1: Brand
- Segment 2: 123  
- Segment 3: Campaign
- Activation ID: ActivationID_2024

## 🚀 Getting Started After Installation

1. **Open Excel** (Desktop or Web)
2. **Look for "IPG Tools"** in the Home ribbon
3. **Click "IPG Taxonomy Extractor"** to open the task pane
4. **Select your taxonomy data** in Excel
5. **Choose extraction options** in the task pane
6. **Click "Extract"** to process your data

## 📈 Success Tips

### Before Installation
- Close all Excel windows
- Ensure stable internet connection
- Have your Microsoft 365 credentials ready

### After Installation  
- Test with sample data first
- Use the undo feature to reverse changes
- Keep the task pane open while working

### For Organizations
- Share this guide with colleagues
- Consider requesting IT to deploy centrally
- Document any organization-specific issues

---

## 📊 Quick Reference Card

| Task | Excel Desktop | Excel Web |
|------|---------------|-----------|
| **Install** | Run PowerShell script | Upload manifest.xml |
| **Uninstall** | Re-run script with `-Uninstall` | Clear browser cache |
| **Update** | Re-run script | Re-upload new manifest |
| **Troubleshoot** | Check PowerShell output | Check browser console |
| **Support** | GitHub Issues | GitHub Issues |

---

**Last Updated:** January 2025  
**Add-in Version:** 2.0.0  
**Compatibility:** Excel 2016+ (Desktop), Excel Web (All browsers)

*This self-installation package was created because central IT deployment wasn't available. For future updates, check the [project repository](https://github.com/henkisdabro/taxonomy-extractor-officeaddin) for new versions.*