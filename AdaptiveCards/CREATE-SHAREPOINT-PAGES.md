# 📝 Creating SharePoint Pages for Teams Integration

## 🎯 **Page 1: TeamsMessageCreator Page**

### **Step 1: Navigate to SharePoint**
1. **Open browser** and go to:
   ```
   https://gustafkliniken.sharepoint.com/sites/Gustafkliniken
   ```

### **Step 2: Create New Page**
1. **Click "Site Pages"** in the left navigation
2. **Click "+ New"** → **"Site Page"**
3. **Page Name**: `TeamsMessageCreator`
4. **Choose layout**: "Blank" (full width)

### **Step 3: Add the Web Part**
1. **Click the "+" (plus icon)** on the page
2. **Search for**: "Adaptive Card Viewer"
3. **Click to add** the web part to the page

### **Step 4: Configure the Web Part**
1. **Click the pencil icon** (edit) on the web part
2. **In the property panel** (should open on the right):
   - Look for **"Component Mode"** or **"Display Mode"**
   - Set it to **"Teams Message Creator"**
   - OR look for **"Card Source"** and set to **"TeamsMessageCreator"**

### **Step 5: Save and Publish**
1. **Click "Save as draft"** (top right)
2. **Click "Publish"** 
3. **Click "Publish"** again to confirm

### **Step 6: Test the Page**
1. **Visit the page directly**:
   ```
   https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/SitePages/TeamsMessageCreator.aspx
   ```
2. **You should see**:
   - Site selector dropdown
   - Message form fields
   - Teams webhook URL input
   - "Create & Distribute" button

---

## 📊 **Page 2: ManagerDashboard Page**

### **Step 1: Create Second Page**
1. **Back to Site Pages** → **"+ New"** → **"Site Page"**
2. **Page Name**: `ManagerDashboard`
3. **Layout**: "Blank"

### **Step 2: Add Web Part**
1. **Click "+"** → Search **"Adaptive Card Viewer"**
2. **Add to page**

### **Step 3: Configure for Dashboard**
1. **Edit the web part** (pencil icon)
2. **Set component to**: "Manager Dashboard" or "ManagerDashboard"

### **Step 4: Save and Publish**
1. **Save as draft** → **Publish**

### **Step 5: Test Dashboard Page**
1. **Visit**:
   ```
   https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/SitePages/ManagerDashboard.aspx
   ```
2. **Should show**: Message analytics, read tracking, user statistics

---

## ⚠️ **Troubleshooting Web Part Configuration**

### **If "Adaptive Card Viewer" web part is not found:**

1. **Try searching for**:
   - "Adaptive"
   - "Card"
   - "Message"
   - Your actual web part name

2. **Check App Catalog**:
   - Your SPFx solution must be deployed
   - Web part must be available in the catalog

3. **Alternative approach**:
   - Add a **"Text"** web part temporarily
   - We'll configure it properly after

### **If Web Part Properties Don't Show Right Options:**

The web part might need specific configuration. Let me know what options you see in the property panel, and I'll help you configure it correctly.

---

## ✅ **Expected Results**

### **TeamsMessageCreator page should show:**
- 🎯 Site selector: "Select Gustav Kliniken Site"
- 📝 Form fields: Title, Content, Priority
- 🔗 Teams webhook URLs textarea
- 📤 "Create & Distribute" button

### **ManagerDashboard page should show:**
- 📊 Message statistics
- 👥 Read tracking data
- 📈 Progress charts
- 🔍 Filter options

---

## 🚀 **Next Steps After Pages Are Created**

1. **Update Teams app** with new package (already prepared)
2. **Test Teams tabs** - should load proper forms
3. **Test message creation** with webhook URLs
4. **Verify read tracking** works with buttons

**Let me know when you've created the first page and what you see!** 📝
