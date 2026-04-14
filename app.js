Office.onReady(() => {
    console.log("Add-in pripravený");

    if (info.host === Office.HostType.Word) {
        console.log("Word pripravený");

        await initUser(); // 👉 načítanie usera hneď po štarte
    }

});



function insertTemplate(type) {
    Word.run(async (context) => {
        const body = context.document.body;

        if (type === 'A') {
            body.insertParagraph("=== ŠABLÓNA A ===", Word.InsertLocation.end);
            body.insertParagraph("Toto je obsah A", Word.InsertLocation.end);
        }

        if (type === 'B') {
            body.insertParagraph("=== ŠABLÓNA B ===", Word.InsertLocation.end);
            body.insertParagraph("Toto je obsah B", Word.InsertLocation.end);
        }

        if (type === 'C') {
            body.insertParagraph("=== ŠABLÓNA C ===", Word.InsertLocation.end);
            body.insertParagraph("Toto je obsah C", Word.InsertLocation.end);
        }

        await context.sync();
    });
}

async function initUser() {
    try {
        const user = await getUserProfile();

        // uložíš si ho globálne
        window.currentUser = user;

        // zobrazíš v UI
        document.getElementById("userInfo").innerText =
            `Prihlásený: ${user.displayName}`;

    } catch (error) {
        console.error("Nepodarilo sa načítať usera:", error);

        // document.getElementById("userInfo").innerText =
        //     "Používateľ nenačítaný";

        //  console.warn("SSO nevyšlo, fallback");

        // const fallback = Office.context.userProfile;

        // window.currentUser = fallback;    
    }
}


async function getAccessToken() {
    return await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true
    });
}



async function getUserProfile() {
    const token = await getAccessToken();

    const response = await fetch(
        "https://graph.microsoft.com/v1.0/me",
        {
            headers: {
                Authorization: `Bearer ${token}`
            }
        }
    );

    return await response.json();
}