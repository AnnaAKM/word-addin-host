Office.onReady(() => {
    // Office ist bereit
    // Globale Registrierung für Event-Handler (wichtig!)
    if (typeof globalThis !== "undefined") {
        globalThis.onDocumentOpen = onDocumentOpen;
    }
});

async function onDocumentOpen(event) {
    try {
        const userInfo = await getUserInfo();

        await Word.run(async (context) => {
            // Platzhalter und ihre Werte
            const replacements = [
                { placeholder: "<<Vorname>>",  value: userInfo.givenName    || "" },
                { placeholder: "<<Nachname>>",  value: userInfo.surname      || "" },
                { placeholder: "<<Titel>>",     value: userInfo.jobTitle     || "" },
                { placeholder: "<<Mail>>",      value: userInfo.mail         || "" },
                { placeholder: "<<Mobil>>",     value: userInfo.mobilePhone  || "" },
            ];

            for (const { placeholder, value } of replacements) {
                // ✅ Erst laden, dann sync, dann auf .items zugreifen
                const results = context.document.body.search(placeholder, { matchCase: false });
                results.load("items");
                await context.sync();

                for (const result of results.items) {
                    result.insertText(value, "Replace");
                }

                // ✅ Ersetzungen für diesen Platzhalter übernehmen
                await context.sync();
            }
        });

    } catch (error) {
        console.error("Fehler in onDocumentOpen:", error);
    } finally {
        // ✅ event.completed() IMMER aufrufen – auch im Fehlerfall
        event.completed();
    }
}

async function getUserInfo() {
    const token = await Office.auth.getAccessToken({ allowSignInPrompt: true });

    const response = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) {
        throw new Error(`Graph API Fehler: ${response.status} ${response.statusText}`);
    }

    return await response.json();
}
