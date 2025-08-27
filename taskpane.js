(function () {
  function log(msg){const el=document.getElementById("log");el.textContent=(el.textContent?el.textContent+"\n":"")+msg;}
  Office.onReady(() => {
    log("Office.js ready");
    const item = Office.context.mailbox.item;

    document.getElementById("btnFetch").addEventListener("click", () => {
      const apiUrl = document.getElementById("apiUrl").value.trim();
      if (!apiUrl) { alert("Enter API URL"); return; }
      item.body.getAsync(Office.CoercionType.Text, {}, async (res) => {
        try{
          const subject = item.subject || "";
          const bodyPreview = res.status === Office.AsyncResultStatus.Succeeded ? (res.value||"").slice(0,5000) : "";
          log("Fetching from API...");
          const resp = await fetch(apiUrl, {
            method: "POST",
            headers: {"Content-Type":"application/json"},
            body: JSON.stringify({ subject, bodyPreview })
          });
          if(!resp.ok) throw new Error("API responded "+resp.status);
          const text = await resp.text();
          document.getElementById("replyText").value = text;
          log("Received reply from API.");
        }catch(e){ log("Error: "+e.message); }
      });
    });

    document.getElementById("btnInsert").addEventListener("click", () => {
      const replyText = document.getElementById("replyText").value || "";
      if(!replyText){ alert("Nothing to insert"); return; }
      item.body.setSelectedDataAsync(replyText, { coercionType: Office.CoercionType.Html }, (res) => {
        if(res.status===Office.AsyncResultStatus.Succeeded){ log("Inserted into draft."); }
        else { log("Insert failed: "+res.error.message); }
      });
    });
  });
})();