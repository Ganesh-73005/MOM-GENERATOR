Office.initialize = function () {
  document.getElementById("checkZoom").addEventListener("click", checkForZoomMeeting);
};

async function checkForZoomMeeting() {
  try {
    const item = Office.context.mailbox.item;
    const body = await getBodyText(item);
    const zoomLink = extractZoomLink(body);

    if (zoomLink) {
      const meetingId = extractZoomId(zoomLink);
      const response = await fetch("http://localhost:3000/get-zoom-details", {
        method: "POST",
        headers: { 
          "Content-Type": "application/json",
          "Authorization": "GANESH73005" // Match your backend secret
        },
        body: JSON.stringify({ meetingId })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      // Handle the response data
    } else {
      Office.context.ui.message("No Zoom link found.");
    }
  } catch (error) {
    console.error("Error:", error);
    Office.context.ui.message("Failed to process meeting. See console for details.");
  }
}
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
