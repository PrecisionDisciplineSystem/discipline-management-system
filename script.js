const msalConfig = {
    auth: {
        clientId: "8942b408-2646-4bbd-bb4f-7a82226351c2",
        authority: "https://login.microsoftonline.com/9eaa6d01-27d2-4363-9466-ebe0ef39b381",
        redirectUri: "https://precisiondisciplinesystem.github.io/discipline-management-system/"
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
