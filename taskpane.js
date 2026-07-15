Office.onReady(async () => {


    // tambahkan di sini
    const supported =
        Office.context.requirements.isSetSupported(
            "IdentityAPI",
            "1.3"
        );

    alert("IdentityAPI supported = " + supported);


    // Workbook URL
    const workbookUrl =
        Office?.context?.document?.url || "Workbook belum tersimpan di OneDrive / SharePoint";

    document.getElementById("workbookUrl").textContent =
        workbookUrl;

    // User Name
    try {

        const token =
            await OfficeRuntime.auth.getAccessToken({
                allowSignInPrompt: true,
                allowConsentPrompt: true
            });

        const response = await fetch(
            "https://graph.microsoft.com/v1.0/me",
            {
                headers: {
                    Authorization: `Bearer ${token}`
                }
            }
        );

        if (!response.ok) {
            throw new Error(
                `Graph API Error (${response.status})`
            );
        }

        const user = await response.json();

        document.getElementById("userName").textContent =
            user.displayName ||
            user.userPrincipalName ||
            "User tidak ditemukan";

    } catch (error) {

        console.error(error);

        document.getElementById("userName").textContent =
            "SSO belum dikonfigurasi";

        document.getElementById("userName").classList.add("error");
    }

});