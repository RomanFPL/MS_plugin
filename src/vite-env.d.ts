/// <reference types="vite/client" />

declare global {
  namespace Office {
    const context: {
      mailbox: {
        item: {
          subject: string;
          itemType: string;
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
          };
        };
      };
    };
    
    const CoercionType: {
      Text: string;
    };
    
    const AsyncResultStatus: {
      Succeeded: string;
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