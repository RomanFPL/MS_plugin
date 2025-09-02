import React, { useState } from 'react';
import './TaskPane.css';

const TaskPane: React.FC = () => {
  const [apiUrl, setApiUrl] = useState('');
  const [replyText, setReplyText] = useState('');
  const [log, setLog] = useState('');

  const logMessage = (message: string) => {
    const timestamp = new Date().toLocaleTimeString();
    setLog((prev: string) => `${prev}[${timestamp}] ${message}\n`);
  };

  const handleFetch = async () => {
    if (!apiUrl.trim()) {
      alert("Enter API URL");
      return;
    }

    // Check if running in Office environment
    if (typeof Office === 'undefined' || !Office.context?.mailbox?.item) {
      logMessage("Testing mode: Using mock data");
      const subject = "Test Subject";
      const bodyPreview = "Test email body content";
      
      try {
        logMessage("Fetching from API...");
        const resp = await fetch(apiUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ subject, bodyPreview })
        });
        
        if (!resp.ok) throw new Error("API responded " + resp.status);
        
        const text = await resp.text();
        setReplyText(text);
        logMessage("Received reply from API.");
      } catch (e: any) {
        logMessage("Error: " + e.message);
      }
      return;
    }

    const item = Office.context.mailbox.item;
    item.body.getAsync(Office.CoercionType.Text, {}, async (res: any) => {
      try {
        const subject = item.subject || "";
        const bodyPreview = res.status === Office.AsyncResultStatus.Succeeded 
          ? (res.value || "").slice(0, 5000) 
          : "";
        
        logMessage("Fetching from API...");
        
        const resp = await fetch(apiUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ subject, bodyPreview })
        });
        
        if (!resp.ok) throw new Error("API responded " + resp.status);
        
        const text = await resp.text();
        setReplyText(text);
        logMessage("Received reply from API.");
      } catch (e: any) {
        logMessage("Error: " + e.message);
      }
    });
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
      <h3 className="text-lg font-semibold mb-4">Simple Reply Generator</h3>
      
      <div className="mb-3">
        <label htmlFor="apiUrl" className="block font-semibold mb-1">API URL</label>
        <input
          id="apiUrl"
          type="text"
          value={apiUrl}
          onChange={(e: React.ChangeEvent<HTMLInputElement>) => setApiUrl(e.target.value)}
          placeholder="https://api.example.com/generate"
          className="w-full px-2 py-2 border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
      </div>
      
      <div className="mb-3">
        <button 
          onClick={handleFetch}
          className="w-full bg-blue-600 text-white py-2 px-4 rounded hover:bg-blue-700 active:bg-blue-800 focus:outline-none focus:ring-2 focus:ring-blue-500"
        >
          Fetch suggestion
        </button>
      </div>
      
      <div className="mb-3">
        <label htmlFor="replyText" className="block font-semibold mb-1">Reply preview</label>
        <textarea
          id="replyText"
          rows={8}
          value={replyText}
          onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => setReplyText(e.target.value)}
          placeholder="Your generated reply will appear here..."
          className="w-full px-2 py-2 border border-gray-300 rounded resize-y focus:outline-none focus:ring-2 focus:ring-blue-500 font-inherit"
        />
      </div>
      
      <div className="mb-3">
        <button 
          onClick={handleInsert}
          className="w-full bg-blue-600 text-white py-2 px-4 rounded hover:bg-blue-700 active:bg-blue-800 focus:outline-none focus:ring-2 focus:ring-blue-500"
        >
          Insert into draft
        </button>
      </div>
      
      <div className="mb-3">
        <div className="whitespace-pre-wrap bg-gray-100 p-2 rounded min-h-[60px] max-h-[150px] overflow-y-auto">{log}</div>
      </div>
    </div>
  );
};

export default TaskPane;