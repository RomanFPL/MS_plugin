import React, { useState } from 'react';
import './TaskPane.css';

interface EmailData {
  subject: string;
  body: string;
  attachments: string[];
}

const TaskPane: React.FC = () => {
  const [apiUrl, setApiUrl] = useState('');
  const [replyText, setReplyText] = useState('');
  const [log, setLog] = useState('');
  const [emailData, setEmailData] = useState<EmailData | null>(null);

  const logMessage = (message: string) => {
    const timestamp = new Date().toLocaleTimeString();
    setLog((prev: string) => `${prev}[${timestamp}] ${message}\n`);
  };

  const analyzeEmail = async () => {
    // Check if running in Office environment
    if (typeof Office === 'undefined' || !Office.context?.mailbox?.item) {
      logMessage("Testing mode: Using mock data");
      const mockData: EmailData = {
        subject: "Test Subject: Project Discussion",
        body: "Hello, I would like to discuss the upcoming project timeline and deliverables. Please let me know your availability for a meeting this week.",
        attachments: ["project_requirements.pdf", "timeline.xlsx"]
      };
      setEmailData(mockData);
      logMessage("Email data loaded (mock)");
      return;
    }

    const item = Office.context.mailbox.item;
    logMessage("Reading email content...");
    
    // Get email body
    item.body.getAsync(Office.CoercionType.Text, {}, async (bodyResult: any) => {
      if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
        logMessage("Error reading email body");
        return;
      }

      // Get attachments
      const attachmentNames: string[] = [];
      if (item.attachments && item.attachments.length > 0) {
        for (let i = 0; i < item.attachments.length; i++) {
          attachmentNames.push(item.attachments[i].name);
        }
        logMessage(`Found ${item.attachments.length} attachment(s)`);
      } else {
        logMessage("No attachments found");
      }

      const emailData: EmailData = {
        subject: item.subject || "",
        body: bodyResult.value || "",
        attachments: attachmentNames
      };

      setEmailData(emailData);
      logMessage("Email data extracted successfully");
    });
  };

  const handleFetch = async () => {
    if (!apiUrl.trim()) {
      alert("Enter API URL");
      return;
    }

    if (!emailData) {
      alert("Please analyze email first");
      return;
    }

    try {
      logMessage("Sending data to AI backend...");
      
      const resp = await fetch(apiUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(emailData)
      });
      
      if (!resp.ok) throw new Error("API responded " + resp.status);
      
      const text = await resp.text();
      setReplyText(text);
      logMessage("Received AI-generated reply");
    } catch (e: any) {
      logMessage("Error: " + e.message);
    }
  };

  const handleInsert = () => {
    if (!replyText.trim()) {
      logMessage("No reply text to insert.");
      return;
    }

    // Check if running in Office environment
    if (typeof Office === 'undefined' || !Office.context?.mailbox?.item) {
      logMessage("Testing mode: Reply would be inserted into draft.");
      return;
    }

    const item = Office.context.mailbox.item;
    if (item.itemType === Office.MailboxEnums.ItemType.Message && item.body) {
      item.body.setAsync(replyText, { coercionType: Office.CoercionType.Text }, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          logMessage("Reply inserted into draft.");
        } else {
          logMessage("Error inserting reply: " + result.error?.message);
        }
      });
    } else {
      logMessage("Unable to insert - not in compose mode.");
    }
  };

  return (
    <div className="font-sans m-4">
      <h3 className="text-lg font-semibold mb-4">Visarsoft AI Assistant</h3>
      
      <div className="mb-3">
        <label htmlFor="apiUrl" className="block font-semibold mb-1">Backend API URL</label>
        <input
          id="apiUrl"
          type="text"
          value={apiUrl}
          onChange={(e: React.ChangeEvent<HTMLInputElement>) => setApiUrl(e.target.value)}
          placeholder="https://your-backend.com/api/analyze"
          className="w-full px-2 py-2 border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
      </div>

      <div className="mb-3">
        <button 
          onClick={analyzeEmail}
          className="w-full bg-green-600 text-white py-2 px-4 rounded hover:bg-green-700 active:bg-green-800 focus:outline-none focus:ring-2 focus:ring-green-500"
        >
          Analyze Email Content
        </button>
      </div>

      {emailData && (
        <div className="mb-3 p-3 bg-blue-50 rounded border">
          <h4 className="font-semibold text-sm mb-2">Email Data:</h4>
          <div className="text-xs space-y-1">
            <div><strong>Subject:</strong> {emailData.subject}</div>
            <div><strong>Body:</strong> {emailData.body.slice(0, 100)}...</div>
            <div><strong>Attachments:</strong> {emailData.attachments.length > 0 ? emailData.attachments.join(', ') : 'None'}</div>
          </div>
        </div>
      )}
      
      <div className="mb-3">
        <button 
          onClick={handleFetch}
          disabled={!emailData}
          className={`w-full py-2 px-4 rounded focus:outline-none focus:ring-2 ${
            emailData 
              ? 'bg-blue-600 text-white hover:bg-blue-700 active:bg-blue-800 focus:ring-blue-500' 
              : 'bg-gray-300 text-gray-500 cursor-not-allowed'
          }`}
        >
          Generate AI Reply
        </button>
      </div>
      
      <div className="mb-3">
        <label htmlFor="replyText" className="block font-semibold mb-1">Generated Reply</label>
        <textarea
          id="replyText"
          rows={8}
          value={replyText}
          onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => setReplyText(e.target.value)}
          placeholder="AI-generated reply will appear here..."
          className="w-full px-2 py-2 border border-gray-300 rounded resize-y focus:outline-none focus:ring-2 focus:ring-blue-500 font-inherit"
        />
      </div>
      
      <div className="mb-3">
        <button 
          onClick={handleInsert}
          disabled={!replyText.trim()}
          className={`w-full py-2 px-4 rounded focus:outline-none focus:ring-2 ${
            replyText.trim() 
              ? 'bg-purple-600 text-white hover:bg-purple-700 active:bg-purple-800 focus:ring-purple-500' 
              : 'bg-gray-300 text-gray-500 cursor-not-allowed'
          }`}
        >
          Insert Reply into Draft
        </button>
      </div>
      
      <div className="mb-3">
        <div className="whitespace-pre-wrap bg-gray-100 p-2 rounded min-h-[60px] max-h-[150px] overflow-y-auto text-xs">{log}</div>
      </div>
    </div>
  );
};

export default TaskPane;