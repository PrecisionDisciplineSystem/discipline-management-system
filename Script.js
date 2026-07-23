const msalConfig = {
    auth: {
        clientId: "PASTE-YOUR-APPLICATION-ID-HERE",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin
    }
};


const msalInstance = new msal.PublicClientApplication(msalConfig);


async function signIn() {

    try {

        const loginResponse = await msalInstance.loginPopup({

            scopes: [
                "User.Read"
            ]

        });


        console.log("Signed in user:");

        console.log(loginResponse.account);


        alert(
            "Welcome " + loginResponse.account.name
        );


    }

    catch(error){

        console.error(error);

        alert(
            "Login failed. Check console for details."
        );

    }

}
