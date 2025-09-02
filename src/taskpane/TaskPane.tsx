import React, { useState } from 'react';
import './TaskPane.css';
import OpenAI from 'openai';

interface EmailData {
  subject: string;
  body: string;
  from: string;
  to: string[];
  attachments: string[];
  dateReceived: string;
}

interface OpenAIMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

const TaskPane: React.FC = () => {
  const [replyText, setReplyText] = useState('');
  const [log, setLog] = useState('');
  const [emailData, setEmailData] = useState<EmailData | null>(null);
  const [responseDescription, setResponseDescription] = useState('');
  const [responseTone, setResponseTone] = useState('formal');
  const [selectedModel, setSelectedModel] = useState('gpt-4o-mini');
  const [useMockMode, setUseMockMode] = useState(false);

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
        from: "sender@example.com",
        to: ["recipient@example.com"],
        attachments: ["project_requirements.pdf", "timeline.xlsx"],
        dateReceived: new Date().toISOString()
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
        from: item.from ? `${item.from.displayName} <${item.from.emailAddress}>` : "Unknown",
        to: item.to ? item.to.map(recipient => `${recipient.displayName} <${recipient.emailAddress}>`) : [],
        attachments: attachmentNames,
        dateReceived: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : new Date().toISOString()
      };

      setEmailData(emailData);
      logMessage("Email data extracted successfully");
    });
  };

  const generateWithOpenAI = async () => {
    logMessage("OpenAI generation started...");
    const apiKey = import.meta.env.VITE_OPENAI_API_KEY;
    logMessage(`API Key available: ${apiKey ? 'YES' : 'NO'}`);
    logMessage(`Full env check: ${JSON.stringify(import.meta.env)}`);
    
    if (!apiKey) {
      logMessage("Error: OpenAI API Key not configured");
      logMessage("Tip: Restart the dev server with 'npm run dev' to load .env changes");
      logMessage("Or enable Mock Mode for testing without API costs.");
      return;
    }

    if (!emailData) {
      logMessage("Error: Please analyze email first");
      return;
    }

    try {
      logMessage("Generating reply with OpenAI...");
      
      const openai = new OpenAI({
        apiKey: apiKey,
        dangerouslyAllowBrowser: true
      });

      const getToneInstruction = () => {
        switch (responseTone) {
          case 'informal': return 'Reply in a friendly, casual tone.';
          case 'harsh': return 'Reply in a direct, assertive tone.';
          default: return 'Reply in a professional, formal tone.';
        }
      };

      const systemInstruction = `You are a helpful assistant that generates email replies. ${getToneInstruction()} Always reply in the same language as the original email.`;
      const userDescription = responseDescription.trim() ? `\n\nAdditional context: ${responseDescription}` : '';

      const messages: OpenAIMessage[] = [
        {
          role: 'system',
          content: systemInstruction
        },
        {
          role: 'user',
          content: `Please generate a reply to this email:\n\nSubject: ${emailData.subject}\nFrom: ${emailData.from}\nDate: ${emailData.dateReceived}\nAttachments: ${emailData.attachments.length > 0 ? emailData.attachments.join(', ') : 'None'}\n\nBody:\n${emailData.body}${userDescription}`
        }
      ];

      const completion = await openai.chat.completions.create({
        model: selectedModel,
        messages: messages,
        max_tokens: 800,
        temperature: 0.7
      });

      const reply = completion.choices[0]?.message?.content || 'No response generated';
      setReplyText(reply);
      logMessage("OpenAI reply generated successfully");
    } catch (e: any) {
      if (e.error?.code === 'invalid_api_key') {
        logMessage("Error: Invalid OpenAI API key. Please check your API key at https://platform.openai.com/account/api-keys");
        logMessage("Tip: You can enable Mock Mode for testing without API costs.");
      } else if (e.error?.type === 'insufficient_quota') {
        logMessage("Error: OpenAI API quota exceeded. Please check your billing at https://platform.openai.com/account/billing");
        logMessage("Tip: You can enable Mock Mode for testing without API costs.");
      } else {
        logMessage("OpenAI Error: " + (e.error?.message || e.message));
      }
    }
  };

  const generateMockResponse = () => {
    logMessage("Mock generation started...");
    logMessage(`Email data: ${emailData ? 'Available' : 'Missing'}`);
    logMessage(`Response tone: ${responseTone}`);
    logMessage(`Response description: "${responseDescription}"`);
    
    if (!emailData) {
      logMessage("Error: Please analyze email first");
      return;
    }

    logMessage("Generating mock response...");
    
    const toneMap = {
      formal: "Thank you for your email regarding",
      informal: "Hi! Thanks for reaching out about",
      harsh: "I need to address your email about"
    };
    
    const baseResponse = toneMap[responseTone as keyof typeof toneMap] || toneMap.formal;
    const context = responseDescription.trim() ? ` ${responseDescription}.` : " your inquiry.";
    
    const mockReply = `${baseResponse} ${emailData.subject}.${context}\n\nI understand your message and will get back to you with more details soon.\n\nBest regards,\n[Your Name]\n\n[This is a mock response - OpenAI API not used]`;
    
    setReplyText(mockReply);
    logMessage("Mock response generated successfully");
  };

  const handleGenerate = () => {
    logMessage(`Starting generation... Mock Mode: ${useMockMode ? 'ON' : 'OFF'}`);
    logMessage(`Email Data available: ${emailData ? 'YES' : 'NO'}`);
    
    if (useMockMode) {
      logMessage("Using Mock generation");
      generateMockResponse();
    } else {
      logMessage("Using OpenAI generation");
      generateWithOpenAI();
    }
  };

  const handleInsertReplace = () => {
    if (!replyText.trim()) {
      logMessage("No reply text to insert.");
      return;
    }

    // Check if running in Office environment
    if (typeof Office === 'undefined' || !Office.context?.mailbox?.item) {
      logMessage("Testing mode: Reply would replace email content.");
      return;
    }

    const item = Office.context.mailbox.item;
    if (item.itemType === Office.MailboxEnums.ItemType.Message && item.body) {
      item.body.setAsync(replyText, { coercionType: Office.CoercionType.Text }, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          logMessage("Reply replaced email content.");
        } else {
          logMessage("Error replacing content: " + result.error?.message);
        }
      });
    } else {
      logMessage("Unable to replace - not in compose mode.");
    }
  };

  const handleInsertPrepend = () => {
    if (!replyText.trim()) {
      logMessage("No reply text to prepend.");
      return;
    }

    // Check if running in Office environment
    if (typeof Office === 'undefined' || !Office.context?.mailbox?.item) {
      logMessage("Testing mode: Reply would be prepended to email.");
      return;
    }

    const item = Office.context.mailbox.item;
    if (item.itemType === Office.MailboxEnums.ItemType.Message && item.body) {
      const formattedReply = `${replyText}\n\n---\n\n`;
      item.body.prependAsync(formattedReply, { coercionType: Office.CoercionType.Text }, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          logMessage("Reply prepended to email.");
        } else {
          logMessage("Error prepending reply: " + result.error?.message);
        }
      });
    } else {
      logMessage("Unable to prepend - not in compose mode.");
    }
  };

  const createNewReply = () => {
    if (!replyText.trim()) {
      logMessage("No reply text to create new email.");
      return;
    }

    // Check if running in Office environment
    if (typeof Office === 'undefined' || !Office.context?.mailbox?.item) {
      logMessage("Testing mode: New reply email would be created.");
      return;
    }

    if (emailData) {
      // Create a new email with reply content
      const replySubject = emailData.subject.toLowerCase().startsWith('re:') 
        ? emailData.subject 
        : `Re: ${emailData.subject}`;
      
      logMessage(`Creating new reply with subject: ${replySubject}`);
      logMessage("New reply functionality requires Office.js compose mode.");
      
      // In a real implementation, you would use Office.context.mailbox.displayNewMessageForm
      // For now, we'll copy the reply to clipboard as a fallback
      navigator.clipboard.writeText(`Subject: ${replySubject}\n\n${replyText}`)
        .then(() => logMessage("Reply copied to clipboard - paste into new email"))
        .catch(() => logMessage("Could not copy to clipboard"));
    }
  };

  return (
    <div className="font-sans m-4">
      <h3 className="text-lg font-semibold mb-4">AI Email Assistant</h3>
      
      <div className="mb-3">
        <div className="flex items-center gap-2 mb-2">
          <span className="text-sm font-semibold">AI Mode</span>
          {!useMockMode && (
            <span className="text-xs text-green-600 bg-green-100 px-2 py-1 rounded">
              ✓ OpenAI Ready
            </span>
          )}
          {useMockMode && (
            <span className="text-xs text-orange-600 bg-orange-100 px-2 py-1 rounded">
              📝 Mock Mode
            </span>
          )}
        </div>
        <label className="flex items-center">
          <input
            type="checkbox"
            checked={useMockMode}
            onChange={(e) => setUseMockMode(e.target.checked)}
            className="mr-2"
          />
          <span className="text-sm">Use Mock Mode (for testing without OpenAI API costs)</span>
        </label>
      </div>
      
      <div className="mb-3">
        <label htmlFor="responseDescription" className="block font-semibold mb-1">Response Description</label>
        <input
          id="responseDescription"
          type="text"
          value={responseDescription}
          onChange={(e: React.ChangeEvent<HTMLInputElement>) => setResponseDescription(e.target.value)}
          placeholder="Describe what you want to communicate in the reply..."
          className="w-full px-2 py-2 border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
      </div>
      
      <div className="mb-3">
        <label className="block font-semibold mb-2">Response Tone</label>
        <div className="flex gap-4">
          <label className="flex items-center">
            <input
              type="radio"
              name="tone"
              value="formal"
              checked={responseTone === 'formal'}
              onChange={(e) => setResponseTone(e.target.value)}
              className="mr-2"
            />
            Formal
          </label>
          <label className="flex items-center">
            <input
              type="radio"
              name="tone"
              value="informal"
              checked={responseTone === 'informal'}
              onChange={(e) => setResponseTone(e.target.value)}
              className="mr-2"
            />
            Informal
          </label>
          <label className="flex items-center">
            <input
              type="radio"
              name="tone"
              value="harsh"
              checked={responseTone === 'harsh'}
              onChange={(e) => setResponseTone(e.target.value)}
              className="mr-2"
            />
            Direct
          </label>
        </div>
      </div>
      
      {!useMockMode && (
        <div className="mb-3">
          <label htmlFor="modelSelect" className="block font-semibold mb-1">AI Model</label>
          <select
            id="modelSelect"
            value={selectedModel}
            onChange={(e) => setSelectedModel(e.target.value)}
            className="w-full px-2 py-2 border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            <option value="gpt-4o-mini">GPT-4o Mini (Fast & Cheap)</option>
            <option value="gpt-4o">GPT-4o (Best Quality)</option>
            <option value="gpt-3.5-turbo">GPT-3.5 Turbo (Legacy)</option>
          </select>
        </div>
      )}

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
            <div><strong>From:</strong> {emailData.from}</div>
            <div><strong>To:</strong> {emailData.to.join(', ') || 'N/A'}</div>
            <div><strong>Date:</strong> {new Date(emailData.dateReceived).toLocaleString()}</div>
            <div><strong>Body:</strong> {emailData.body.slice(0, 150)}...</div>
            <div><strong>Attachments:</strong> {emailData.attachments.length > 0 ? emailData.attachments.join(', ') : 'None'}</div>
          </div>
        </div>
      )}
      
      <div className="mb-3">
        <button 
          onClick={handleGenerate}
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
        <div className="grid grid-cols-1 gap-2">
          <button 
            onClick={handleInsertReplace}
            disabled={!replyText.trim()}
            className={`w-full py-2 px-4 rounded focus:outline-none focus:ring-2 text-sm ${
              replyText.trim() 
                ? 'bg-purple-600 text-white hover:bg-purple-700 active:bg-purple-800 focus:ring-purple-500' 
                : 'bg-gray-300 text-gray-500 cursor-not-allowed'
            }`}
          >
            Replace Email Content
          </button>
          
          <button 
            onClick={handleInsertPrepend}
            disabled={!replyText.trim()}
            className={`w-full py-2 px-4 rounded focus:outline-none focus:ring-2 text-sm ${
              replyText.trim() 
                ? 'bg-indigo-600 text-white hover:bg-indigo-700 active:bg-indigo-800 focus:ring-indigo-500' 
                : 'bg-gray-300 text-gray-500 cursor-not-allowed'
            }`}
          >
            Prepend to Email
          </button>
          
          <button 
            onClick={createNewReply}
            disabled={!replyText.trim()}
            className={`w-full py-2 px-4 rounded focus:outline-none focus:ring-2 text-sm ${
              replyText.trim() 
                ? 'bg-green-600 text-white hover:bg-green-700 active:bg-green-800 focus:ring-green-500' 
                : 'bg-gray-300 text-gray-500 cursor-not-allowed'
            }`}
          >
            Copy as New Reply
          </button>
        </div>
      </div>
      
      <div className="mb-3">
        <div className="whitespace-pre-wrap bg-gray-100 p-2 rounded min-h-[60px] max-h-[150px] overflow-y-auto text-xs">{log}</div>
      </div>
    </div>
  );
};

export default TaskPane;