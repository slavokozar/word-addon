Office.onReady(() => {
    console.log("Add-in pripravený");


    console.log(Office.context.userProfile);

    cosnole.log(OfficeRuntime.auth.getAccessToken());

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