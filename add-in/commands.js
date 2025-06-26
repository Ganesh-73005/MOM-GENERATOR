// Define helper functions first
function getBodyText(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => resolve(result.value));
  });
}

function extractZoomLink(text) {
  const regex = /https:\/\/[\w-]*\.?zoom.us\/(j|my)\/[\d\w?=-]+/i;
  return text.match(regex)?.[0];
}

function extractZoomId(link) {
  return link.split("/j/")[1]?.split("?")[0];
}

function showMessage(message) {
  // Check which message API is available
  if (Office.context.ui && Office.context.ui.displayDialogAsync) {
    // Modern Office versions
    Office.context.ui.displayDialogAsync(
      `https://yourdomain.com/message.html?text=${encodeURIComponent(message)}`,
      { height: 50, width: 300 }
    );
  } else if (Office.context.mailbox && Office.context.mailbox.item) {
    // Fallback for older versions
    Office.context.mailbox.item.notificationMessages.addAsync("status", {
      type: "informationalMessage",
      message: message,
      icon: "icon.16x16",
      persistent: false
    });
  } else {
    // Last resort - use console
    console.log("Message:", message);
  }
}

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
      const response = await fetch("http://localhost:3000/get-zoom-details", {
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
                showMessage("Meeting details added successfully!");
                resolve();
              }
            }
          );
        });
      } else {
        throw new Error("Email body editing not available in this context");
      }
    } else {
      showMessage("No Zoom link found in this email.");
    }
  } catch (error) {
    console.error("Error in checkForZoomMeeting:", error);
    showMessage("Error processing meeting: " + error.message);
  }
}
