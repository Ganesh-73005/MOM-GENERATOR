Office.initialize = function () {
  document.getElementById("checkZoom").addEventListener("click", checkForZoomMeeting);
};

async function checkForZoomMeeting() {
  try {
    // Verify we have a valid mail item
    if (!Office.context.mailbox || !Office.context.mailbox.item) {
      throw new Error("No email item selected");
    }

    const item = Office.context.mailbox.item;
    const body = await getBodyText(item);
    const zoomLink = extractZoomLink(body);

    if (zoomLink) {
      const meetingId = extractZoomId(zoomLink);
      const response = await fetch("https://b485-192-193-107-47.ngrok-free.app/get-zoom-details", {
        method: "POST",
        headers: { 
          "Content-Type": "application/json",
          "Authorization": "GANESH73005"
        },
        body: JSON.stringify({ meetingId })
      });

      if (!response.ok) {
        throw new Error(`API request failed: ${response.status}`);
      }

      const data = await response.json();
      
      // Safely update the email body
      if (item.body && item.body.setAsync) {
        return new Promise((resolve, reject) => {
          item.body.setAsync(
            `${body}<hr><h2>Zoom Meeting Details</h2><p>Meeting ID: ${data.id}</p>`,
            { coercionType: Office.CoercionType.Html },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
              } else {
                resolve();
              }
            }
          );
        });
      } else {
        throw new Error("Email body editing not available in this context");
      }
    } else {
      Office.context.ui.message("No Zoom link found in this email.");
    }
  } catch (error) {
    console.error("Error in checkForZoomMeeting:", error);
    Office.context.ui.message("Error processing meeting: " + error.message);
  }
}
