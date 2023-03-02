let mailboxItem;
Office.initialize = function () {
    mailboxItem = Office.context.mailbox.item;
}

function validateBody(event) {
    //mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);

    console.log("validateBody");

    Office.context.ui.displayDialogAsync("https://mail.geoffrey1.msftonlinelab.com/ItemSendWeb/success.html",
        { height: 50, width: 50, displayInIframe: true}, dialogCallback);


}

function dialogCallback(asyncResult) {
    if (asyncResult.status == "failed") {

        // In addition to general system errors, there are 3 specific errors for 
        // displayDialogAsync that you can handle individually.
        switch (asyncResult.error.code) {
            case 12004:
                console.log("Domain is not trusted");
                break;
            case 12005:
                console.log("HTTPS is required");
                break;
            case 12007:
                console.log("A dialog is already opened.");
                break;
            default:
                console.log(asyncResult.error.message);
                break;
        }
    }
    else {


        dialog = asyncResult.value;

        console.log("dialog value");
      

    }

}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}