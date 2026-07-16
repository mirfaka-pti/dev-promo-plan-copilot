Office.onReady(async () => {
  try {
    const manifest = Office.context.diagnostics.manifest;

    // Buat elemen container untuk info WebApplicationInfo
    const infoContainer = document.createElement("div");
    infoContainer.style.fontFamily = "Segoe UI, sans-serif";
    infoContainer.style.fontSize = "13px";
    infoContainer.style.marginTop = "12px";
    infoContainer.style.padding = "10px";
    infoContainer.style.borderTop = "1px solid #ddd";

    if (manifest && manifest.webApplicationInfo) {
      const webAppInfo = manifest.webApplicationInfo;

      infoContainer.innerHTML = `
        <h3 style="margin-bottom:6px;">WebApplicationInfo</h3>
        <p><strong>Client ID:</strong> ${webAppInfo.id}</p>
        <p><strong>Resource:</strong> ${webAppInfo.resource}</p>
        <p><strong>Redirect URI:</strong> ${webAppInfo.redirectUri}</p>
      `;
    } else {
      infoContainer.innerHTML = `
        <h3 style="margin-bottom:6px;">WebApplicationInfo</h3>
        <p style="color:red;">Tidak ditemukan di manifest.</p>
      `;
    }

    // Tambahkan ke body taskpane
    document.body.appendChild(infoContainer);
  } catch (error) {
    const errorContainer = document.createElement("div");
    errorContainer.style.color = "red";
    errorContainer.style.marginTop = "12px";
    errorContainer.textContent = `Gagal membaca WebApplicationInfo: ${error.message}`;
    document.body.appendChild(errorContainer);
  };
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