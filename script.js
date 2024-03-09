Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("displayEmailDetails").onclick = () => displayEmailDetails();
    }
});

function displayEmailDetails() {
    Office.context.mailbox.item.subject.getAsync((result) => {
        let subject = result.value;
        let from = Office.context.mailbox.item.from;
        
        alert(`Subject: ${subject}\nFrom: ${from ? from.emailAddress : "N/A"}`);
    });
}
