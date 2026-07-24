const msalConfig = {

    auth: {

        clientId:
        "8942b408-2646-4bbd-bb4f-7a82226351c2",

        authority:
        "https://login.microsoftonline.com/9eaa6d01-27d2-4363-9466-ebe0ef39b381",

        redirectUri:
        "https://precisiondisciplinesystem.github.io/discipline-management-system/"

    }

};


const msalInstance =
new msal.PublicClientApplication(msalConfig);



// Check if user is already signed in

function checkLogin(){


    const accounts =
    msalInstance.getAllAccounts();


    if(accounts.length > 0){

        return accounts[0];

    }


    return null;

}




async function signIn(){


    try{


        const loginResponse =
        await msalInstance.loginPopup({

            scopes:[
                "User.Read"
            ]

        });



        sessionStorage.setItem(
            "user",
            JSON.stringify(loginResponse.account)
        );



        window.location.href =
        "dashboard.html";


    }


    catch(error){


        console.error(error);


        alert(
        "Login failed."
        );


    }

}





function signOut(){


    msalInstance.clearCache();


    sessionStorage.clear();


    window.location.href =
    "login.html";


}
