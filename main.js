Office.onReady(() => {
    // Wird aufgerufen, wenn Office bereit ist
});

async function onDocumentOpen(event) {
    try {
        const userInfo = await getUserInfo();

        await Word.run(async (context) => {
            const body = context.document.body;
            body.search("<<Vorname>>", { matchCase: false }).items[0].insertText(userInfo.givenName, "Replace");
            body.search("<<Nachname>>", { matchCase: false }).items[0].insertText(userInfo.surname, "Replace");
            body.search("<<Titel>>", { matchCase: false }).items[0].insertText(userInfo.jobTitle || "", "Replace");
            body.search("<<Mail>>", { matchCase: false }).items[0].insertText(userInfo.mail, "Replace");
            body.search("<<Mobil>>", { matchCase: false }).items[0].insertText(userInfo.mobilePhone || "", "Replace");
            await context.sync();
        });

        event.completed();
    } catch (error) {
        console.error(error);
        event.completed();
    }
}

async function getUserInfo() {
    const token = await Office.auth.getAccessToken({ allowSignInPrompt: true });
    const response = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${token}` }
    });
    return await response.json();
}
