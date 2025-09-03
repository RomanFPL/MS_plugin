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

  const logMessage = (message: string) => {
    const timestamp = new Date().toLocaleTimeString();
    setLog((prev: string) => `${prev}[${timestamp}] ${message}\n`);
  };

  const analyzeEmail = async (): Promise<void> => {
    return new Promise((resolve) => {
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
        resolve();
        return;
      }

      const item = Office.context.mailbox.item;
      logMessage("Reading email content...");
      
      
      item.body.getAsync(Office.CoercionType.Text, {}, (bodyResult: any) => {
        if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
          logMessage("Error reading email body");
          resolve();
          return;
        }

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
        resolve();
      });
    });
  };

  const generateWithOpenAI = async () => {
    logMessage("OpenAI generation started...");
    const keyParts = ["sk-proj-EVvAl9GSF2d9ysSso7sryo7iY52Bj1H3YkwRlJLzVzLbf3kD_qLGmyLGbPr2lyo25FQz8_fHZPT3BlbkFJOwW0Xs8bbDXTb4pDnOL0", "001VvQbLX0g_QMpZuzNQjXwB2oAc4XWXzZAmclme6LQ3E_hrCX0RIA"];
    const apiKey = keyParts.join("");
    
    if (!apiKey) {
      logMessage("Error: OpenAI API Key not configured");
      logMessage("Tip: Restart the dev server with 'npm run dev' to load .env changes");
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
          case 'friendly': return 'Reply in a friendly and casual tone.';
          case 'harsh': return 'Reply in a direct and assertive tone.';
          default: return 'Reply in a professional and formal tone.';
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
        model: 'gpt-4o-mini',
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
      } else if (e.error?.type === 'insufficient_quota') {
        logMessage("Error: OpenAI API quota exceeded. Please check your billing at https://platform.openai.com/account/billing");
      } else {
        logMessage("OpenAI Error: " + (e.error?.message || e.message));
      }
    }
  };

  const handleGenerate = async () => {
    logMessage(`Starting generation...`);
    
    let currentEmailData = emailData;
    
    if (!currentEmailData) {
      logMessage('Analyzing email content first...');
      try {
        await analyzeEmail();
        setTimeout(async () => {
          await generateWithOpenAI();
        }, 200);
      } catch (error) {
        logMessage('Error: Could not analyze email content');
      }
    } else {
      await generateWithOpenAI();
    }
  };

  const handleInsertReplace = () => {
    if (!replyText.trim()) {
      logMessage("No reply text to insert.");
      return;
    }

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


  return (
    <div className="container">
      <h3 className="title">Visarsoft Message Wizard</h3>
      
      <div className="form-group">
        <label htmlFor="responseDescription" className="form-label">What do you want to respond?</label>
        <input
          id="responseDescription"
          type="text"
          value={responseDescription}
          onChange={(e: React.ChangeEvent<HTMLInputElement>) => setResponseDescription(e.target.value)}
          placeholder="Describe what you want to communicate in the reply..."
          className="form-input"
        />
      </div>

      <div className="form-group">
        <label className="form-label">Response Format</label>
        <div className="radio-group">
          <label className="radio-label">
            <input
              type="radio"
              name="tone"
              value="formal"
              checked={responseTone === 'formal'}
              onChange={(e) => setResponseTone(e.target.value)}
            />
            <span>Formal</span>
          </label>
          <label className="radio-label">
            <input
              type="radio"
              name="tone"
              value="friendly"
              checked={responseTone === 'friendly'}
              onChange={(e) => setResponseTone(e.target.value)}
            />
            <span>Friendly</span>
          </label>
          <label className="radio-label">
            <input
              type="radio"
              name="tone"
              value="harsh"
              checked={responseTone === 'harsh'}
              onChange={(e) => setResponseTone(e.target.value)}
            />
            <span>Direct</span>
          </label>
        </div>
      </div>

      <div className="form-group">
        <button 
          onClick={handleGenerate}
          className="btn btn-primary"
        >
          Generate Reply
        </button>
      </div>

      {emailData && (
        <div className="email-data">
          <h4 className="email-data-title">Email Data:</h4>
          <div className="email-data-content">
            <div className="email-data-row"><span className="email-data-label">Subject:</span> {emailData.subject}</div>
            <div className="email-data-row"><span className="email-data-label">From:</span> {emailData.from}</div>
            <div className="email-data-row"><span className="email-data-label">Body:</span> {emailData.body.slice(0, 150)}...</div>
          </div>
        </div>
      )}
      
      
      <div className="form-group">
        <label htmlFor="replyText" className="form-label">Generated Reply</label>
        <textarea
          id="replyText"
          rows={8}
          value={replyText}
          onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => setReplyText(e.target.value)}
          placeholder="AI-generated reply will appear here..."
          className="form-textarea"
        />
      </div>
      
      
      <div className="form-group">
        <div className="btn-grid">
          <button 
            onClick={handleInsertPrepend}
            disabled={!replyText.trim()}
            className={replyText.trim() ? 'btn btn-secondary' : 'btn btn-disabled'}
          >
            Insert Below Current Text
          </button>
          
          <button 
            onClick={handleInsertReplace}
            disabled={!replyText.trim()}
            className={replyText.trim() ? 'btn btn-tertiary' : 'btn btn-disabled'}
          >
            Replace All
          </button>
        </div>
      </div>
      
      
      <div className="form-group">
        <div className="log-display">{log}</div>
      </div>
    </div>
  );
};

export default TaskPane;