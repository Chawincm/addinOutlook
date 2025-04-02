Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        // Add a click event handler for the button
        document.getElementById("forwardButton").onclick = forwardEmail;
    }
});

function forwardEmail() {
    Office.context.mailbox.item.forwardAsync({
        toRecipients: ["dlpadmin@chr.co.th"],
        subject: "Forwarded Email",
        htmlBody: "This email has been forwarded to you."
    }, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Email forwarded successfully.");
        } else {
            console.error("Failed to forward email: " + asyncResult.error.message);
        }
    });
}
