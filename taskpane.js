Office.initialize = function () {

};


Office.onReady(info =>{
    document.getElementById('address').innerHTML = `<p>${Office.context.mailbox.userProfile.emailAddress}`;
})