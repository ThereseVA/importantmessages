# 🔗 Getting Teams Channel Webhook URLs

## **Step-by-Step Guide**

### **Method 1: Teams Channel Connectors**

1. **Open Teams Channel**:
   - Go to the channel where you want to receive messages
   - Click the **three dots (⋯)** next to the channel name

2. **Add Connector**:
   - Select **"Connectors"**
   - Find **"Incoming Webhook"**
   - Click **"Configure"**

3. **Configure Webhook**:
   - **Name**: "Adaptive Card Messages" 
   - **Upload image**: Optional logo
   - Click **"Create"**

4. **Copy Webhook URL**:
   - Teams provides a URL like:
   ```
   https://outlook.office.com/webhook/abc123def456.../IncomingWebhook/xyz789...
   ```
   - **Copy this URL** - you'll paste it in your Teams tab

### **Method 2: Power Automate (Advanced)**

1. **Create Flow**:
   - Trigger: "When a HTTP request is received"
   - Action: "Post adaptive card in Teams channel"

2. **Get HTTP URL**:
   - Power Automate gives you a trigger URL
   - Use this instead of webhook URL

## **Using Webhook URLs in Your App**

### **In Create Message Tab**:

1. **Add Webhook URLs**:
   ```
   HR Team: https://outlook.office.com/webhook/hr-channel...
   IT Team: https://outlook.office.com/webhook/it-channel...
   All Staff: https://outlook.office.com/webhook/general-channel...
   ```

2. **Select Channels**:
   - Check which channels should receive the message
   - Your adaptive card will be sent to selected channels

3. **Distribution Happens Automatically**:
   - `TeamsDistributionService` sends to all selected webhooks
   - Each channel receives the same adaptive card
   - All buttons work the same way

## **Example Workflow**

```
1. Create message in Teams tab
2. Add webhook URLs:
   ✅ General Channel: https://outlook.office.com/webhook/general...
   ✅ HR Channel: https://outlook.office.com/webhook/hr...
   
3. Click "Create & Distribute"
4. Message appears in both channels with read buttons
5. Users click "I Have Read" → tracked in dashboard
```

## **Testing Your Setup**

### **Quick Test**:
1. **Get one webhook URL** from any Teams channel
2. **Create a test message** in your Teams tab
3. **Add the webhook URL**
4. **Send the message**
5. **Check the Teams channel** - your adaptive card should appear
6. **Click "I Have Read"** - it should record in your dashboard

### **Success Indicators**:
- ✅ Adaptive card appears in Teams channel
- ✅ Card shows priority icon and message content
- ✅ "I Have Read" button is clickable
- ✅ Clicking button records in SharePoint
- ✅ Manager dashboard updates with read status

**Your Teams integration is working! 🎉**
