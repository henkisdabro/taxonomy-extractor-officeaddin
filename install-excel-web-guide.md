# 🌐 IPG Taxonomy Extractor - Excel Web Installation Guide

> **Individual User Self-Installation for Excel Online**

## 📋 Quick Installation Steps

### Method 1: Direct Upload (Recommended)

1. **Open Excel Web**
   - Go to [office.com](https://office.com) and sign in
   - Open any Excel workbook or create a new one

2. **Access Add-ins Menu**
   - Click **Home** → **Add-ins** → **More Add-ins**
   - OR Click **Insert** → **Office Add-ins**

3. **Upload the Add-in**
   - Click **Upload My Add-in** (top right corner)
   - Click **Browse** and select the manifest file
   - Download manifest: [manifest.xml](https://ipg-taxonomy-extractor-addin.wookstar.com/manifest.xml)
   - Click **Upload**

4. **Activate the Add-in**
   - The add-in should appear in your ribbon under "IPG Tools"
   - Click **IPG Taxonomy Extractor** to open the task pane

### Method 2: Organization Store (If Available)

1. **Check Organization Add-ins**
   - Go to **Home** → **Add-ins** → **More Add-ins**
   - Look in the **My Organization** section
   - If available, click **Add** next to "IPG Taxonomy Extractor"

## 🔧 Alternative Installation Methods

### Browser Bookmarklet (Advanced Users)

Create a bookmark with this JavaScript code to auto-install:

```javascript
javascript:(function(){
  var manifestUrl = 'https://ipg-taxonomy-extractor-addin.wookstar.com/manifest.xml';
  
  if (typeof Office !== 'undefined' && Office.addin) {
    try {
      Office.addin.showAsTaskpane(manifestUrl);
      alert('IPG Taxonomy Extractor installation initiated!');
    } catch(e) {
      window.open('https://ipg-taxonomy-extractor-addin.wookstar.com/install-guide.html', '_blank');
    }
  } else {
    alert('Please run this bookmarklet while in Excel Web');
  }
})();
```

### URL-Based Installation

Visit this URL while logged into Excel Web:
```
https://office.com/launch/excel?auth=1&addins=https://ipg-taxonomy-extractor-addin.wookstar.com/manifest.xml
```

## ⚠️ Troubleshooting Common Issues

### Issue: "Upload My Add-in" Button Missing

**Possible Causes:**
- Organization policy blocks custom add-ins
- Browser compatibility issues
- Insufficient permissions

**Solutions:**
1. **Check Browser**: Use Chrome, Edge, or Firefox (latest versions)
2. **Clear Cache**: Clear browser cache and cookies for office.com
3. **Try Incognito**: Test in private/incognito browsing mode
4. **Contact IT**: Your organization may have disabled custom add-ins

### Issue: Manifest Upload Fails

**Common Errors:**
- "Invalid manifest format"
- "Manifest cannot be parsed"
- "Security policy violation"

**Solutions:**
1. **Re-download Manifest**: Get fresh copy from [here](https://ipg-taxonomy-extractor-addin.wookstar.com/manifest.xml)
2. **Check File Extension**: Ensure file ends with `.xml`
3. **Validate Manifest**: File should start with `<?xml version="1.0"`

### Issue: Add-in Doesn't Appear After Upload

**Solutions:**
1. **Refresh Page**: Press F5 or Ctrl+R to reload Excel Web
2. **Sign Out/In**: Sign out of Office.com and sign back in
3. **Try Different Browser**: Switch to a different web browser
4. **Check Ribbon**: Look for "IPG Tools" group in Home tab

### Issue: Add-in Loads But Doesn't Function

**Solutions:**
1. **Check Internet**: Ensure stable internet connection
2. **Disable Browser Extensions**: Turn off ad blockers or script blockers
3. **Allow Popups**: Enable popups for office.com domain
4. **Update Browser**: Use the latest browser version

## 🔒 Security & Privacy

### Data Handling
- **No Data Storage**: Add-in processes data locally, nothing sent to external servers
- **HTTPS Only**: All communications use secure encrypted connections
- **Microsoft Approved**: Follows Microsoft Office Add-in security standards

### Browser Requirements
- **Modern Browser**: Chrome 90+, Edge 90+, Firefox 88+, Safari 14+
- **JavaScript Enabled**: Required for Office Add-ins functionality
- **Cookies Allowed**: office.com domain must accept cookies

## 📱 Mobile Device Support

### Excel Mobile Apps
- **iOS**: Limited support (view-only functionality)
- **Android**: Limited support (view-only functionality)
- **Web Mobile**: Full functionality in mobile browsers

### Recommended Approach for Mobile
Use Excel Web in your mobile browser rather than the native apps for full add-in functionality.

## 🚫 Removal Instructions

### Uninstall from Excel Web

1. **Open Excel Web**
2. **Go to Add-ins**: Home → Add-ins → More Add-ins
3. **Manage Add-ins**: Click "Manage My Add-ins" or "My Add-ins"
4. **Remove**: Find "IPG Taxonomy Extractor" and click "Remove"

### Clear Browser Cache (Complete Removal)

1. **Chrome/Edge**: Ctrl+Shift+Delete → Clear browsing data
2. **Firefox**: Ctrl+Shift+Delete → Clear recent history
3. **Safari**: Develop → Empty Caches

## 🆘 Support Resources

### Quick Help
- **Documentation**: [Project README](https://github.com/henkisdabro/taxonomy-extractor-officeaddin)
- **Video Tutorials**: Coming soon
- **FAQ**: See below

### Contact Support
- **GitHub Issues**: [Report Problems](https://github.com/henkisdabro/taxonomy-extractor-officeaddin/issues)
- **Email Support**: Available through GitHub

## ❓ Frequently Asked Questions

### Q: Do I need admin rights to install this?
**A:** No, this is a user-level installation that doesn't require administrator privileges.

### Q: Will this work with my organization's Office 365?
**A:** Yes, unless your IT department has specifically blocked custom add-ins.

### Q: Is my data secure?
**A:** Yes, all processing happens locally in your browser. No data is sent to external servers.

### Q: Can I use this offline?
**A:** No, Excel Web requires an internet connection, and the add-in needs to connect to its hosting server.

### Q: Why doesn't it appear in Excel Desktop?
**A:** This is the web-specific installation guide. For Excel Desktop, use the PowerShell installation script.

### Q: Can multiple people in my organization use this?
**A:** Yes, each user needs to install it individually using this guide, or your IT team can deploy it centrally.

---

**Last Updated:** January 2025  
**Version:** 2.0.0  
**Compatibility:** Excel Web (All browsers)