/// <reference types="vite/client" />

declare global {
  namespace Office {
    const context: {
      mailbox: {
        item: {
          subject: string;
          itemType: string;
          attachments: Array<{
            name: string;
            attachmentType: string;
            size: number;
          }>;
          body: {
            getAsync(
              coercionType: string,
              options: any,
              callback: (result: any) => void
            ): void;
            setAsync(
              data: string,
              options: any,
              callback: (result: any) => void
            ): void;
            prependAsync(
              data: string,
              options: any,
              callback: (result: any) => void
            ): void;
          };
          from: {
            displayName: string;
            emailAddress: string;
          };
          to: Array<{
            displayName: string;
            emailAddress: string;
          }>;
          dateTimeCreated: Date;
        };
      };
    };
    
    const CoercionType: {
      Text: string;
      Html: string;
    };
    
    const AsyncResultStatus: {
      Succeeded: string;
      Failed: string;
    };
    
    const MailboxEnums: {
      ItemType: {
        Message: string;
      };
    };
    
    function onReady(callback: () => void): void;
  }
}

export {};