Office.initialize = function () {
  document.getElementById("checkZoom").addEventListener("click", checkForZoomMeeting);
};

async function checkForZoomMeeting() {
  const item = Office.context.mailbox.item;
  const body = await getBodyText(item);
  const zoomLink = extractZoomLink(body);

  if (zoomLink) {
    const meetingId = extractZoomId(zoomLink);
    const response = await fetch("https://your-backend.ngrok.io/get-zoom-details", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ meetingId }),
    });

    const data = await response.json();
    Office.context.mailbox.item.body.setAsync(
      `${body}<hr><h2>Zoom Meeting Details</h2><p>Meeting ID: ${data.id}</p>`,
      { coercionType: Office.CoercionType.Html }
    );
  } else {
    Office.context.ui.message("No Zoom link found.");
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