# 📧 Create TeamsChannels SharePoint List - Step by Step Guide

## 🎯 Method 1: Manual Creation (Easiest - 5 minutes)

### Step 1: Navigate to SharePoint
1. Open your browser
2. Go to: `https://gustafkliniken.sharepoint.com/sites/Gustafkliniken`
3. Click on **"New"** button (top left)
4. Select **"List"**

### Step 2: Create the List
1. Choose **"Blank list"**
2. **Name**: `TeamsChannels`
3. **Description**: `Configuration for Teams channel email addresses`
4. Click **"Create"**

### Step 3: Add Columns (Click "Add column" for each)

#### Column 1: ChannelName
- Click **"+ Add column"** → **"Single line of text"**
- **Name**: `ChannelName`
- **Description**: `Name of the Teams channel (e.g., General, Announcements)`
- **Required**: ✅ Yes
- Click **"Save"**

#### Column 2: ChannelEmail  
- Click **"+ Add column"** → **"Single line of text"**
- **Name**: `ChannelEmail`
- **Description**: `Email address of the Teams channel`
- **Required**: ✅ Yes
- Click **"Save"**

#### Column 3: TeamName
- Click **"+ Add column"** → **"Single line of text"**
- **Name**: `TeamName` 
- **Description**: `Name of the Teams team (e.g., IT Department)`
- **Required**: ✅ Yes
- Click **"Save"**

#### Column 4: Department
- Click **"+ Add column"** → **"Single line of text"**
- **Name**: `Department`
- **Description**: `Department filter (e.g., IT, HR, Management, All)`
- **Required**: ❌ No
- Click **"Save"**

#### Column 5: MessageTypes
- Click **"+ Add column"** → **"Single line of text"**  
- **Name**: `MessageTypes`
- **Description**: `Priority levels to send (e.g., High,Medium,Low)`
- **Required**: ❌ No
- Click **"Save"**

#### Column 6: IsActive
- Click **"+ Add column"** → **"Yes/No"**
- **Name**: `IsActive`
- **Description**: `Whether to send messages to this channel`
- **Default value**: ✅ Yes
- **Required**: ✅ Yes
- Click **"Save"**

### Step 4: Add Sample Entries

Click **"+ New"** to add these sample entries:

#### Entry 1: IT General
- **Title**: `IT Department - General`
- **ChannelName**: `General`
- **ChannelEmail**: `[REPLACE WITH ACTUAL EMAIL]`
- **TeamName**: `IT Department`
- **Department**: `IT`
- **MessageTypes**: `High,Medium,Low`
- **IsActive**: ✅ Yes

#### Entry 2: HR Announcements  
- **Title**: `HR Department - Announcements`
- **ChannelName**: `Announcements`
- **ChannelEmail**: `[REPLACE WITH ACTUAL EMAIL]`
- **TeamName**: `HR Department`
- **Department**: `HR`
- **MessageTypes**: `High,Medium`
- **IsActive**: ✅ Yes

#### Entry 3: Management Urgent
- **Title**: `Management - Urgent Only`
- **ChannelName**: `Urgent Communications`
- **ChannelEmail**: `[REPLACE WITH ACTUAL EMAIL]`
- **TeamName**: `Management Team`
- **Department**: `Management`
- **MessageTypes**: `High`
- **IsActive**: ✅ Yes

## 🔗 How to Get Teams Channel Email Addresses

For each entry above, you need to get the actual email address:

### For Each Teams Channel:
1. **Open Microsoft Teams** (desktop or web)
2. **Navigate to the team** (e.g., "IT Department")
3. **Find the channel** (e.g., "General")
4. **Click the "..." (three dots)** next to the channel name
5. **Select "Get email address"**
6. **Copy the email address** (looks like: `general_abc123@gustafkliniken.onmicrosoft.com`)
7. **Go back to SharePoint** and edit the entry
8. **Replace `[REPLACE WITH ACTUAL EMAIL]`** with the real email address

## 📋 Your Completed List Should Look Like:

| Title | ChannelName | ChannelEmail | TeamName | Department | MessageTypes | IsActive |
|-------|-------------|--------------|----------|------------|--------------|----------|
| IT Department - General | General | general_abc123@gustafkliniken.onmicrosoft.com | IT Department | IT | High,Medium,Low | ✅ |
| HR Department - Announcements | Announcements | announcements_def456@gustafkliniken.onmicrosoft.com | HR Department | HR | High,Medium | ✅ |
| Management - Urgent Only | Urgent Communications | urgent_ghi789@gustafkliniken.onmicrosoft.com | Management Team | Management | High | ✅ |

## ✅ Testing Your Setup

Once you have real email addresses, you can test by:

1. **Send a test email** to one of the channel email addresses
2. **Check Teams** - the email should appear as a message
3. **Verify formatting** looks good

## 🚀 Next Steps

After creating the list:

1. ✅ **Get real Teams channel emails** (instructions above)
2. ✅ **Update the sample entries** with real email addresses  
3. ✅ **Add more channels** as needed
4. ✅ **Deploy the SPFx solution** with TeamsChannelService
5. ✅ **Test automated sending** from your solution

## 💡 Tips

- **Channel emails can change** - Teams may regenerate them occasionally
- **Test regularly** - send a test message to verify emails still work
- **Use IsActive column** - disable channels temporarily without deleting
- **Department filtering** - helps target messages to specific teams
- **MessageTypes** - controls which priority levels each channel receives

## 🆘 Need Help?

If you get stuck:
1. Make sure you have **contribute permissions** on the SharePoint site
2. Try using **SharePoint in a different browser** if you have issues
3. Contact your **SharePoint admin** if you can't create lists

---

**📝 Once this is set up, your code will automatically:**
- Read configured channels from this list
- Send messages to appropriate channels based on department/priority
- Handle success/failure reporting
- Allow easy management without code changes
