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


    const loggedOut =
    sessionStorage.getItem("loggedOut");


    // User manually signed out of the application

    if(loggedOut){

        return null;

    }



    const accounts =
    msalInstance.getAllAccounts();



    if(accounts.length > 0){

        return accounts[0];

    }



    return null;

}







// Microsoft Login

async function signIn(){


    try{


        const loginResponse =
        await msalInstance.loginPopup({

            scopes:[
                "User.Read"
            ]

        });



        // Remove logout flag after successful login

        sessionStorage.removeItem(
            "loggedOut"
        );



        // Store user information

        sessionStorage.setItem(

            "user",

            JSON.stringify(loginResponse.account)

        );



        // Send user to dashboard

        window.location.href =
        "dashboard.html";


    }


    catch(error){


        console.error(error);


        alert(
            "Login failed. Check console for details."
        );


    }

}







// Application Logout

function signOut(){


    // Mark application as logged out

    sessionStorage.setItem(

        "loggedOut",

        "true"

    );



    // Clear saved application data

    sessionStorage.removeItem(
        "user"
    );



    // Return to login page

    window.location.href =
    "login.html";


}

function loadContent(page){


    const content =
    document.getElementById("mainContent");



    switch(page){


        case "dashboard":

            content.innerHTML = `

            <h1>
            Dashboard
            </h1>

            <p>
            Welcome to Precision Discipline System.
            </p>


            <h3>
            Action Required
            </h3>


            <p>
            No documents currently require attention.
            </p>

            `;

        break;



        case "generate":

            content.innerHTML = `

            <h1>
            Generate Discipline
            </h1>


            <p>
            Discipline document generation will be built here.
            </p>

            `;

        break;




        case "history":

            content.innerHTML = `

            <h1>
            Employee History
            </h1>


            <p>
            Employee records will appear here.
            </p>

            `;

        break;




        case "upload":

            content.innerHTML = `

            <h1>
            Upload Signed Documents
            </h1>


            <p>
            Signed document uploads will appear here.
            </p>

            `;

        break;




        case "search":

            content.innerHTML = `

            <h1>
            Search Cases
            </h1>


            <p>
            Case search will be built here.
            </p>

            `;

        break;




        case "calculator":

            content.innerHTML = `

            <h1>
            Progression Calculator
            </h1>


            <p>
            Progression calculator will be built here.
            </p>

            `;

        break;




        case "settings":

            content.innerHTML = `

            <h1>
            Settings
            </h1>


            <p>
            Admin settings will be built here.
            </p>

            `;

        break;


    }


}
