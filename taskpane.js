Office.initialize = function () {

};


Office.onReady(async info =>{
    document.getElementById('address').innerHTML = `<p>${Office.context.mailbox.userProfile.emailAddress}`;
    document.getElementById('token').innerHTML = `<p>${await OfficeRuntime.auth.getAccessToken({ allowConsentPrompt: false, allowSignInPrompt: false })}`;
    
})