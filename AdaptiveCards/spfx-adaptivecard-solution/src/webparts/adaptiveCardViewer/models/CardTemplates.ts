// Card templates as TypeScript objects
export const sampleCardTemplate = {
  "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.3",
  "body": [
    {
      "type": "TextBlock",
      "text": "Welcome Card",
      "weight": "Bolder",
      "size": "Large"
    },
    {
      "type": "TextBlock",
      "text": "This card is loaded from TypeScript!",
      "wrap": true
    }
  ]
};

/**
 * Generate Adaptive Card JSON from SharePoint message for Teams/Email distribution
 */
export function generateMessageCard(message: any): any {
  const priorityColor = message.Priority === 'High' ? 'Attention' : 
                       message.Priority === 'Medium' ? 'Warning' : 'Good';
  
  const priorityIcon = message.Priority === 'High' ? '🚨' : 
                      message.Priority === 'Medium' ? '⚠️' : 'ℹ️';

  return {
    "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
      {
        "type": "ColumnSet",
        "columns": [
          {
            "type": "Column",
            "width": "auto",
            "items": [
              {
                "type": "TextBlock",
                "text": priorityIcon,
                "size": "Large"
              }
            ]
          },
          {
            "type": "Column",
            "width": "stretch",
            "items": [
              {
                "type": "TextBlock",
                "text": message.Title,
                "weight": "Bolder",
                "size": "Large",
                "color": priorityColor,
                "wrap": true
              },
              {
                "type": "TextBlock",
                "text": `Priority: ${message.Priority}`,
                "size": "Small",
                "color": priorityColor,
                "weight": "Bolder"
              }
            ]
          }
        ]
      },
      {
        "type": "TextBlock",
        "text": message.MessageContent,
        "wrap": true,
        "spacing": "Medium"
      },
      {
        "type": "FactSet",
        "facts": [
          {
            "title": "From:",
            "value": message.Author?.Title || "System"
          },
          {
            "title": "Target:",
            "value": message.TargetAudience
          },
          {
            "title": "Expires:",
            "value": new Date(message.ExpiryDate).toLocaleDateString()
          }
        ],
        "spacing": "Medium"
      }
    ],
    "actions": [
      {
        "type": "Action.Http",
        "title": "✅ Mark as Read",
        "method": "POST",
        "url": `https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/_api/web/lists/getbytitle('Message Read Actions')/items`,
        "headers": [
          {
            "name": "Content-Type",
            "value": "application/json"
          }
        ],
        "body": JSON.stringify({
          MessageId: message.Id,
          UserEmail: "${userEmail}",
          UserDisplayName: "${userName}",
          ReadTimestamp: new Date().toISOString(),
          DeviceInfo: "Teams/Email"
        })
      },
      {
        "type": "Action.OpenUrl",
        "title": "📊 View Dashboard",
        "url": `https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/SitePages/Dashboard.aspx`
      }
    ]
  };
}

/**
 * Generate simplified card for Teams (with enhanced read tracking)
 */
export function generateTeamsCard(message: any): any {
  const priorityIcon = message.Priority === 'High' ? '🚨' : 
                      message.Priority === 'Medium' ? '⚠️' : 'ℹ️';

  // Get the correct SharePoint site URL for proper linking
  // Always use the correct subsite regardless of context
  const siteUrl = 'https://gustafkliniken.sharepoint.com/sites/Gustafkliniken';

  return {
    "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
      {
        "type": "TextBlock",
        "text": `${priorityIcon} ${message.Title}`,
        "weight": "Bolder",
        "size": "Large",
        "wrap": true
      },
      {
        "type": "TextBlock",
        "text": message.MessageContent,
        "wrap": true,
        "spacing": "Medium"
      },
      {
        "type": "FactSet",
        "facts": [
          {
            "title": "Priority:",
            "value": message.Priority
          },
          {
            "title": "From:",
            "value": message.Author?.Title || "System"
          },
          {
            "title": "Expires:",
            "value": new Date(message.ExpiryDate).toLocaleDateString()
          }
        ]
      },
      {
        "type": "Container",
        "style": "emphasis",
        "items": [
          {
            "type": "TextBlock",
            "text": "� **Action Required:** Please confirm you have read this message",
            "wrap": true,
            "weight": "Bolder",
            "size": "Small"
          }
        ],
        "spacing": "Medium"
      }
    ],
    "actions": [
      {
        "type": "Action.Submit",
        "title": "✅ I Have Read This Message",
        "data": {
          "action": "markAsRead",
          "messageId": message.Id,
          "messageTitle": message.Title
        }
      },
      {
        "type": "Action.OpenUrl", 
        "title": "📊 View Dashboard",
        "url": `https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/SitePages/Dashboard.aspx`
      },
      {
        "type": "Action.OpenUrl",
        "title": "📋 All Messages",
        "url": `https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/Lists/Important%20Messages/AllItems.aspx`
      }
    ]
  };
}

export const dashboardCardTemplate = {
  "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.3",
  "body": [
    {
      "type": "TextBlock",
      "text": "Dashboard Metrics",
      "weight": "Bolder",
      "size": "Large",
      "color": "Accent"
    },
    {
      "type": "FactSet",
      "facts": [
        {
          "title": "Active Users:",
          "value": "${activeUsers}"
        },
        {
          "title": "Total Sales:",
          "value": "${totalSales}"
        },
        {
          "title": "Last Updated:",
          "value": "${lastUpdated}"
        }
      ]
    }
  ],
  "actions": [
    {
      "type": "Action.Submit",
      "title": "Refresh Data",
      "data": {
        "action": "refresh"
      }
    }
  ]
};

export const cardTemplates = {
  sample: sampleCardTemplate,
  dashboard: dashboardCardTemplate
};
